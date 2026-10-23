"""
AUTOMAÇÃO DE TEMPERATURA - TERMINAL DE COMBUSTÍVEIS
Baixa DiaryLoadAutoTank do New SAT e roda macro SAP na planilha.
Ciclo: a cada 30 minutos. Reset de log: meia-noite.

v11:
- NOVO: lê no Outlook Web (aba extra do mesmo Chrome) os emails de
  "ABERTURA DE TANQUES", "Virada de ..." e "Atualização de densidade" do Pool
  e preenche os campos AMARELOS da aba INICIO (H = temperatura
  da amostra, I = densidade da amostra x1000). Ex.: GAS 0,7170 / 21,5 -> 717 / 21,5.
  Cada email só é aplicado uma vez; viradas futuras ("às 18h") esperam o horário.
  Se o corpo do email não tiver a tabela em texto, tenta baixar o PDF anexo.
- NOVO: janela de funcionamento 01:00 -> 21:30 (fora disso só mantém sessões).
- NOVO: após cada atualização SAP, grava as 2 linhas do log (Dep. 01 e POOL)
  na planilha LOGs.xlsx da rede; às 23:59 apaga o log da rede.

v10 (remoção do fluxo de Outlook/PDF, que estava órfão — não era mais chamado
     em nenhum ciclo. O que ficou:
- Base: download do New SAT + Excel/macro, com o Chrome em modo --incognito,
  necessário para não ser bloqueado pelo CrowdStrike.
- driver com retry na criação (não mata processo Chrome — só espera e tenta de novo).
- NOVO: flag _macro_rodando — keep-alive do SAP não interfere durante a macro.
- NOVO: verificação de SAP conectado antes de rodar a macro (evita erro seco).
- NOVO: trava de instância única (não deixa abrir o programa duas vezes).
- NOVO: painel Tkinter — Iniciar/Parar/Forçar Ciclo, status, contadores, log ao vivo.
- NOVO: impede suspensão do Windows enquanto o programa está rodando.
"""

import os
import re
import sys
import json
import time
import logging
import shutil
import schedule
import subprocess
import unicodedata
import urllib.parse
from datetime import datetime, date, timedelta, time as dtime
from pathlib import Path

import queue
import threading
import pythoncom
import tkinter as tk
from tkinter import scrolledtext

import win32com.client
import win32gui
import win32con
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# ── Usuário do Windows atual ──────────────────────────────────────────────────
import getpass
import ctypes
import msvcrt
USUARIO_WINDOWS = getpass.getuser()

# ── Carrega configurações ──────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
with open(BASE_DIR / "config.json", encoding="utf-8") as f:
    CFG = json.load(f)

# Substitui {usuario} nos caminhos do config pelo usuário real do Windows
def _substituir_usuario(obj):
    if isinstance(obj, str):
        # {pasta} = pasta onde está este script (deixa tudo junto e portátil)
        return obj.replace("{usuario}", USUARIO_WINDOWS).replace("{pasta}", str(BASE_DIR))
    elif isinstance(obj, dict):
        return {k: _substituir_usuario(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_substituir_usuario(i) for i in obj]
    return obj

CFG = _substituir_usuario(CFG)

# ── Pastas necessárias ─────────────────────────────────────────────────────────
for pasta in [
    CFG["newsat"]["pasta_download"],
    CFG["automacao"]["pasta_log"],
]:
    Path(pasta).mkdir(parents=True, exist_ok=True)

# ── Logger ─────────────────────────────────────────────────────────────────────
def configurar_logger():
    pasta_log = Path(CFG["automacao"]["pasta_log"])
    arquivo_log = pasta_log / f"temperatura_{date.today().strftime('%Y%m%d')}.log"
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.handlers.clear()

    fh = logging.FileHandler(arquivo_log, encoding="utf-8")
    fh.setFormatter(fmt)
    root_logger.addHandler(fh)

    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    root_logger.addHandler(sh)
    # O LogHandler da UI é adicionado depois pela AutoBPAApp._construir_ui()

configurar_logger()
log = logging.getLogger(__name__)

# ── Flag global: impede keep-alive do SAP de interferir durante a macro ───────
_macro_rodando = False

# ── Trava: garante que só um ciclo (agendado OU forçado) rode por vez ─────────
_ciclo_lock = threading.Lock()

# ── Contador de erros consecutivos, pra não floodar o operador de notificação ─
_erros_consecutivos = 0
NOTIFICAR_A_CADA_N_ERROS = 4  # com ciclo de 30min, ~1 notificação a cada 2h de falha contínua

# ── Referência global da UI (preenchida quando roda com painel Tkinter) ───────
_ui_app = None


# ── Trava de instância única — impede rodar o programa duas vezes ─────────────
_lock_file = None

def _garantir_instancia_unica():
    global _lock_file
    lock_path = BASE_DIR / "monitor.lock"
    try:
        _lock_file = open(lock_path, 'w')
        msvcrt.locking(_lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        _lock_file.write(str(os.getpid()))
        _lock_file.flush()
    except Exception:
        print("=" * 50)
        print("  AVISO: O programa já está rodando!")
        print("  Feche a outra janela antes de abrir nova.")
        print("=" * 50)
        import sys
        sys.exit(1)


# ── Impede suspensão do Windows enquanto o programa estiver rodando ───────────
_timeout_original_ac = None
_timeout_original_dc = None

def _ler_timeouts_suspensao_atuais():
    """
    Lê os timeouts de standby (AC/DC) configurados na máquina ANTES de
    zerá-los, pra poder devolver o valor original quando o programa parar.

    Usa `powercfg /q` (resolve o valor efetivo, venha ele de uma customização
    do usuário ou do padrão herdado do esquema) e pega os dois ÚLTIMOS números
    hexadecimais da saída (Current AC / Current DC sempre vêm por último,
    depois de Minimum/Maximum/Increment) — não depende do texto do rótulo,
    que é localizado (varia por idioma do Windows). Registro direto foi
    descartado: a chave só existe se o valor já tiver sido customizado
    manualmente naquele perfil, então falha em qualquer usuário "no padrão".
    """
    global _timeout_original_ac, _timeout_original_dc
    try:
        saida = subprocess.run(
            ["powercfg", "/q", "SCHEME_CURRENT", "SUB_SLEEP", "STANDBYIDLE"],
            capture_output=True, timeout=5, text=True
        ).stdout
        valores_hex = re.findall(r"0x[0-9a-fA-F]+", saida)
        if len(valores_hex) < 2:
            log.warning("  Não foi possível interpretar a saída do powercfg (formato inesperado).")
            return
        ac_seg = int(valores_hex[-2], 16)
        dc_seg = int(valores_hex[-1], 16)
        _timeout_original_ac = ac_seg // 60
        _timeout_original_dc = dc_seg // 60
        log.info(
            f"  Timeout de suspensão original salvo: "
            f"AC={_timeout_original_ac}min DC={_timeout_original_dc}min"
        )
    except Exception as e:
        log.warning(f"  Não foi possível ler o timeout de suspensão original: {e}")


def _impedir_suspensao():
    try:
        _ler_timeouts_suspensao_atuais()
        subprocess.run(
            ["powercfg", "/change", "standby-timeout-ac", "0"],
            capture_output=True, timeout=5
        )
        subprocess.run(
            ["powercfg", "/change", "standby-timeout-dc", "0"],
            capture_output=True, timeout=5
        )
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 | 0x00000001 | 0x00000002
        )
        log.info("  Sistema configurado para não suspender enquanto o programa rodar.")
    except Exception as e:
        log.warning(f"  Não foi possível impedir suspensão: {e}")


def _liberar_suspensao():
    try:
        # Restaura o valor que a máquina tinha antes (ou um fallback razoável
        # se por algum motivo não conseguimos ler o original).
        ac = _timeout_original_ac if _timeout_original_ac is not None else 20
        dc = _timeout_original_dc if _timeout_original_dc is not None else 10
        subprocess.run(
            ["powercfg", "/change", "standby-timeout-ac", str(ac)],
            capture_output=True, timeout=5
        )
        subprocess.run(
            ["powercfg", "/change", "standby-timeout-dc", str(dc)],
            capture_output=True, timeout=5
        )
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
        log.info(f"  Suspensão restaurada: AC={ac}min DC={dc}min")
    except Exception:
        pass


# ── Notificação de erro pro operador (balão na bandeja do Windows) ────────────
_NOTIF_ID = 1

def _notificar_erro_operador(mensagem: str):
    """
    Mostra uma notificação nativa do Windows (balão na bandeja do sistema)
    avisando que o ciclo de temperatura deu erro. Não bloqueia a automação —
    dispara e segue. O ícone some sozinho depois de alguns segundos.
    """
    try:
        hwnd = _ui_app.root.winfo_id() if _ui_app else win32gui.GetDesktopWindow()
        hicon = win32gui.LoadIcon(0, win32con.IDI_ERROR)
        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP | win32gui.NIF_INFO
        nid = (
            hwnd, _NOTIF_ID, flags, win32con.WM_USER + 20, hicon,
            "AutoBPA Monitor",
            mensagem[:255],
            10000,
            "Erro na automação de temperatura",
            win32gui.NIIF_ERROR,
        )
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        log.info("  Notificação de erro enviada ao operador (bandeja do Windows).")

        def _remover_icone():
            try:
                win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (hwnd, _NOTIF_ID))
            except Exception:
                pass
        threading.Timer(15, _remover_icone).start()
    except Exception as e:
        log.warning(f"  Não foi possível notificar o operador: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# 3. NEW SAT — download do DiaryLoadAutoTank
#    Fluxo de navegação/export IDÊNTICO ao script antigo (já funciona).
#    O que mudou: driver com PERFIL FIXO do Chrome (em vez de --incognito),
#    então a sessão sobrevive a reinícios do script — não perde login toda hora.
# ══════════════════════════════════════════════════════════════════════════════
COOKIES_FILE = BASE_DIR / "newsat_cookies.json"
URL_RELATORIO = CFG["newsat"]["url_relatorio"]  # ex.: https://<servidor-newsat>/Report/DiaryLoadAutoTank
_driver_global = None  # Chrome fica aberto entre ciclos
_chromedriver_path_cache = None  # path do chromedriver já resolvido — evita checar de novo toda vez


def _obter_chromedriver_path() -> str:
    """
    Resolve o path do chromedriver uma única vez por execução do programa.
    ChromeDriverManager().install() faz checagem de versão (às vezes
    download) via rede — em rede corporativa isso pode ser lento ou falhar,
    então só chamamos de fato na primeira vez e reaproveitamos o path depois.
    """
    global _chromedriver_path_cache
    if _chromedriver_path_cache is None:
        from webdriver_manager.chrome import ChromeDriverManager
        _chromedriver_path_cache = ChromeDriverManager().install()
        log.info(f"  ChromeDriver resolvido e cacheado: {_chromedriver_path_cache}")
    return _chromedriver_path_cache


# ── Modo "Chrome do robô": em vez de o Selenium abrir (e fechar) um Chrome
#    anônimo, o robô abre um Chrome NORMAL com perfil próprio + porta de
#    depuração e só se CONECTA nele. O Chrome não fecha junto com o programa
#    e o login (Outlook / New SAT) fica salvo no perfil.
#    Liga com "chrome_porta_debug": 9222 no config.json.
CHROME_PORTA = CFG.get("chrome_porta_debug")
CHROME_PERFIL = Path(CFG.get("chrome_perfil", str(BASE_DIR / "chrome_perfil_robo")))


def _porta_aberta(porta: int) -> bool:
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex(("127.0.0.1", int(porta))) == 0


def _caminho_chrome() -> str | None:
    candidatos = [CFG.get("chrome_exe", ""),
                  r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                  r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                  os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")]
    return next((c for c in candidatos if c and Path(c).exists()), None)


def _preparar_preferencias_perfil():
    """
    Mesmas liberações que o Chrome anônimo do script antigo tinha (prefs):
    downloads automáticos/múltiplos permitidos, sem perguntar onde salvar,
    pasta de download = pasta do New SAT. Só é lido quando o Chrome abre.
    """
    try:
        arq = CHROME_PERFIL / "Default" / "Preferences"
        arq.parent.mkdir(parents=True, exist_ok=True)
        try:
            prefs = json.loads(arq.read_text(encoding="utf-8"))
        except Exception:
            prefs = {}
        prefs.setdefault("download", {}).update({
            "default_directory": str(Path(CFG["newsat"]["pasta_download"])),
            "prompt_for_download": False,
            "directory_upgrade": True,
        })
        prefs.setdefault("profile", {}).setdefault("default_content_setting_values", {})[
            "automatic_downloads"] = 1
        prefs.setdefault("safebrowsing", {})["enabled"] = False
        prefs.setdefault("plugins", {})["always_open_pdf_externally"] = True
        arq.write_text(json.dumps(prefs), encoding="utf-8")
    except Exception as e:
        log.warning(f"  Não consegui ajustar as preferências do Chrome do robô: {e}")


def _abrir_chrome_robo() -> bool:
    """Abre o Chrome do robô (se ainda não estiver aberto) — processo separado, sobrevive ao script."""
    if _porta_aberta(CHROME_PORTA):
        return True
    exe = _caminho_chrome()
    if not exe:
        log.error("  chrome.exe não encontrado — coloque o caminho em config.json -> chrome_exe")
        return False
    CHROME_PERFIL.mkdir(parents=True, exist_ok=True)
    _preparar_preferencias_perfil()
    log.info(f"  Abrindo o Chrome do robô (perfil: {CHROME_PERFIL})...")
    subprocess.Popen(
        [exe, f"--remote-debugging-port={CHROME_PORTA}", f"--user-data-dir={CHROME_PERFIL}",
         "--no-first-run", "--no-default-browser-check", "--start-maximized", URL_OUTLOOK],
        creationflags=0x00000008 | 0x00000200,  # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
        close_fds=True,
    )
    for _ in range(20):
        time.sleep(1)
        if _porta_aberta(CHROME_PORTA):
            time.sleep(2)
            return True
    log.error("  O Chrome do robô não abriu a porta de depuração.")
    return False


def _conectar_chrome_robo(pasta_download: Path):
    if not _abrir_chrome_robo():
        raise RuntimeError("Chrome do robô indisponível")
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{CHROME_PORTA}")
    driver = webdriver.Chrome(service=Service(_obter_chromedriver_path()), options=chrome_options)
    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
            "behavior": "allow", "downloadPath": str(pasta_download)})
    except Exception as e:
        log.warning(f"  Não consegui definir a pasta de download: {e}")
    log.info("  Conectado ao Chrome do robô.")
    return driver


def _criar_driver(pasta_download: Path):
    """
    Cria o Chrome em modo --incognito (igual ao script antigo) — necessário
    pra não ser bloqueado pelo CrowdStrike. Sem perfil persistente em disco,
    então o login precisa ser refeito quando a sessão expira (mas o driver
    continua aberto e é reaproveitado entre ciclos, sem recriar toda hora).
    """
    if CHROME_PORTA:
        return _conectar_chrome_robo(pasta_download)
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": str(pasta_download),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "plugins.always_open_pdf_externally": True,  # PDF do email baixa em vez de abrir
    })
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--safebrowsing-disable-download-protection")
    from webdriver_manager.chrome import ChromeDriverManager
    service = Service(_obter_chromedriver_path())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": str(pasta_download)
    })
    return driver


def _criar_driver_com_retry(pasta_download: Path, tentativas: int = 3):
    """Tenta criar o driver; se falhar, espera e tenta de novo (sem mexer em processos)."""
    global _chromedriver_path_cache
    for i in range(tentativas):
        try:
            return _criar_driver(pasta_download)
        except Exception as e:
            log.warning(f"  Falha ao criar Chrome (tentativa {i+1}/{tentativas}): {e}")
            # Invalida o cache do driver — pode ter sido causa da falha (ex:
            # Chrome atualizado no meio da execução, driver ficou desatualizado)
            _chromedriver_path_cache = None
            time.sleep(5)
    log.error("  Não foi possível criar o Chrome após todas as tentativas.")
    return None


_abas = {}  # nome -> window handle ("newsat", "email") dentro do mesmo Chrome


def _obter_driver():
    """Reaproveita o Chrome aberto ou cria um novo (escondido). Retorna None se não der."""
    global _driver_global, _abas
    if _driver_global is not None:
        try:
            _driver_global.title  # testa se ainda está vivo
            return _driver_global
        except Exception:
            _driver_global = None
    driver = _criar_driver_com_retry(Path(CFG["newsat"]["pasta_download"]))
    if driver is None:
        return None
    _driver_global = driver
    if CHROME_PORTA:
        # Chrome do robô já pode ter abas abertas: reaproveita pela URL
        _abas = {}
        for h in driver.window_handles:
            try:
                driver.switch_to.window(h)
                u = driver.current_url.lower()
            except Exception:
                continue
            if "newsat" in u:
                _abas.setdefault("newsat", h)
            elif "outlook" in u or "microsoftonline" in u:
                _abas.setdefault("email", h)
        livres = [h for h in driver.window_handles if h not in _abas.values()]
        if "newsat" not in _abas and livres:
            _abas["newsat"] = livres[0]
    else:
        _abas = {"newsat": driver.current_window_handle}
    _esconder_janela(driver)
    return driver


def _ir_para_aba(driver, nome: str):
    """Troca para a aba do New SAT ou do Outlook (cria a aba se ainda não existir)."""
    h = _abas.get(nome)
    if h and h in driver.window_handles:
        driver.switch_to.window(h)
    else:
        driver.switch_to.new_window("tab")
        _abas[nome] = driver.current_window_handle
    try:
        driver.execute_cdp_cmd("Page.bringToFront", {})  # mostra na tela a aba que o robô está usando
    except Exception:
        pass


def _esconder_janela(driver):
    """
    Joga a janela do Chrome pra fora da área visível da tela (canto invisível).
    Mantém o Chrome visível de verdade (não headless, evita bloqueio do
    CrowdStrike) mas fora do caminho enquanto você trabalha no PC.
    """
    if CHROME_PORTA:
        return  # Chrome do robô: o robô não mexe na janela
    if CFG.get("mostrar_chrome", False):
        _mostrar_janela(driver)  # config "mostrar_chrome": true -> Chrome fica visível na tela
        return
    try:
        driver.set_window_position(-32000, -32000)
    except Exception as e:
        log.warning(f"  Não foi possível esconder a janela do Chrome: {e}")


def _mostrar_janela(driver):
    """Traz a janela do Chrome de volta pra tela — usado só quando precisa de login manual."""
    try:
        driver.set_window_position(0, 0)
        driver.maximize_window()
    except Exception as e:
        log.warning(f"  Não foi possível mostrar a janela do Chrome: {e}")


def _salvar_cookies(driver):
    cookies = driver.get_cookies()
    COOKIES_FILE.write_text(json.dumps(cookies), encoding="utf-8")
    log.info(f"  Cookies salvos em {COOKIES_FILE}")


def _ja_logado(driver) -> bool:
    """Verifica se a página atual é o relatório (não a tela de login)."""
    return "login" not in driver.current_url.lower() and "microsoftonline" not in driver.current_url.lower()


def _aguardar_login_manual(driver, tentativas: int | None = None, intervalo: int = 5) -> bool:
    """
    Mostra a janela do Chrome e espera o operador logar manualmente,
    checando a cada `intervalo` segundos (até `tentativas` vezes). Ao
    detectar o login, salva os cookies e esconde a janela de novo.
    Usado tanto no download quanto no keep-alive — mesma lógica nos dois.
    """
    log.warning("  Sessão expirada. Aguardando login manual no Chrome...")
    if tentativas is None:  # padrão: 10 min pra logar (config "tempo_login_min")
        tentativas = int(CFG.get("tempo_login_min", 10)) * 60 // intervalo
    log.warning(f"  >>> FAÇA O LOGIN NO CHROME QUE ABRIU (até {tentativas * intervalo // 60} min) <<<")
    _mostrar_janela(driver)
    for _ in range(tentativas):
        time.sleep(intervalo)
        if _ja_logado(driver):
            log.info("  Login detectado! Sessão renovada.")
            _salvar_cookies(driver)
            _esconder_janela(driver)
            return True
    log.error("  Timeout no login manual.")
    return False


def baixar_diary_newsat() -> str | None:
    global _driver_global
    log.info("Acessando New SAT para baixar DiaryLoadAutoTank...")
    pasta_dest = Path(CFG["newsat"]["pasta_download"])
    arquivo_esperado = pasta_dest / "DiaryLoadAutoTank.xlsx"

    # Remove arquivo antigo antes de baixar novo
    if arquivo_esperado.exists():
        arquivo_esperado.unlink()

    driver = None
    try:
        # Reutiliza Chrome aberto ou cria novo (com retry)
        driver = _obter_driver()
        if driver is None:
            log.error("  Chrome indisponível. Abortando download.")
            return None
        _ir_para_aba(driver, "newsat")
        wait = WebDriverWait(driver, 60)

        # Acessa relatório direto
        driver.get(URL_RELATORIO)
        time.sleep(3)

        # Verifica se caiu na tela de login Microsoft
        if not _ja_logado(driver):
            if not _aguardar_login_manual(driver):
                return None
            # Navega para o relatório após login confirmado
            driver.get(URL_RELATORIO)
            time.sleep(3)

        log.info("  Logado no New SAT.")

        # Às vezes o New SAT joga pra /Home depois do login — volta pro relatório
        for _ in range(3):
            if "diaryloadautotank" in driver.current_url.lower():
                break
            log.info(f"  New SAT abriu {driver.current_url} — indo de novo para o relatório...")
            driver.get(URL_RELATORIO)
            time.sleep(5)
        if "diaryloadautotank" not in driver.current_url.lower():
            log.error("  New SAT não abriu a página do relatório DiaryLoadAutoTank.")
            driver.save_screenshot(str(BASE_DIR / "newsat_erro.png"))
            return None

        # Aguarda página carregar completamente
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(5)

        # Salva screenshot para diagnóstico
        driver.save_screenshot(str(BASE_DIR / "newsat_apos_login.png"))
        log.info(f"  URL atual: {driver.current_url}")

        # Clica em Pesquisar (sem preencher data — usa hoje automaticamente)
        btn_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "btn-search")))
        btn_pesquisar.click()
        log.info("  Clicou em Pesquisar.")

        # Aguarda iframe do relatório aparecer
        log.info("  Aguardando iframe do relatório...")
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//iframe | //frame")
        ))
        time.sleep(8)

        # Salva screenshot para ver o estado atual
        driver.save_screenshot(str(BASE_DIR / "newsat_resultados.png"))
        log.info("  Resultados carregados.")

        # Entra no iframe e clica nos botões de exportar
        iframes = driver.find_elements(By.XPATH, "//iframe | //frame")
        log.info(f"  Encontrou {len(iframes)} iframe(s).")
        entrou_iframe = False
        for iframe in iframes:
            try:
                driver.switch_to.frame(iframe)
                # Verifica se o botão de exportar está neste iframe
                btns = driver.find_elements(By.ID, "rptRelatorio_ctl05_ctl04_ctl00_ButtonImg")
                if btns:
                    entrou_iframe = True
                    log.info("  Botão exportar encontrado no iframe.")
                    break
                driver.switch_to.default_content()
            except Exception:
                driver.switch_to.default_content()

        if not entrou_iframe:
            # Tenta na página principal
            driver.switch_to.default_content()

        # Clica no botão exportar
        try:
            btn_exportar = wait.until(EC.presence_of_element_located(
                (By.ID, "rptRelatorio_ctl05_ctl04_ctl00_ButtonImg")
            ))
            # Clique por JS: no Chrome do robô a barra do New SAT fica por cima do botão
            driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                                  btn_exportar)
            log.info("  Clicou no botão exportar.")
            time.sleep(2)

            # Clica em Excel
            btn_excel = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//a[@title='Excel']")
            ))
            driver.execute_script("arguments[0].click();", btn_excel)
            log.info("  Clicou em Excel.")
        except Exception as e:
            log.warning(f"  Botão não encontrado, tentando JS: {e}")
            # Se o relatório ainda estiver processando internamente, a chamada
            # de exportação falha com "the report or page is being updated" —
            # não é erro definitivo, é só cedo demais. Tenta de novo algumas vezes.
            exportou = False
            ultimo_erro = None
            for tentativa in range(3):
                try:
                    driver.execute_script("$find('rptRelatorio').exportReport('EXCELOPENXML');")
                    log.info("  Exportação via JS.")
                    exportou = True
                    break
                except Exception as e2:
                    ultimo_erro = e2
                    if "being updated" in str(e2) or "InvalidOperationException" in str(e2):
                        log.warning(f"  Relatório ainda processando, aguardando e tentando de novo ({tentativa+1}/3)...")
                        time.sleep(5)
                    else:
                        break  # erro diferente — não adianta insistir
            if not exportou:
                log.error(f"  Falha total ao exportar: {ultimo_erro}")
                driver.save_screenshot(str(BASE_DIR / "newsat_erro_export.png"))
                driver.switch_to.default_content()
                return None

        log.info("  Aguardando download...")
        # Reconfigura CDP para garantir pasta correta após login
        try:
            driver.switch_to.default_content()
            driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": str(pasta_dest)
            })
        except Exception:
            pass

        # Aguarda download (até 3 minutos)
        log.info("  Aguardando arquivo na pasta...")
        baixando = False
        t_export = time.time() - 5
        pasta_downloads_win = Path.home() / "Downloads"
        for i in range(90):
            time.sleep(2)
            # Se o Chrome salvou na pasta Downloads do Windows, traz pra pasta certa
            try:
                perdidos = [f for f in pasta_downloads_win.glob("DiaryLoadAutoTank*.xlsx")
                            if f.stat().st_mtime >= t_export]
                if perdidos and pasta_downloads_win.resolve() != pasta_dest.resolve():
                    shutil.move(str(perdidos[0]), str(arquivo_esperado))
                    log.info(f"  Arquivo caiu em {pasta_downloads_win} — movido para {pasta_dest}.")
            except Exception:
                pass
            # Verifica se está baixando (.crdownload)
            crdownload = list(pasta_dest.glob("*.crdownload"))
            if crdownload and not baixando:
                baixando = True
                log.info("  Download em andamento...")

            # Verifica arquivo esperado
            if arquivo_esperado.exists():
                log.info(f"  DiaryLoadAutoTank baixado: {arquivo_esperado}")
                return str(arquivo_esperado)

            # Verifica qualquer xlsx novo (exceto a planilha principal)
            xlsx = [f for f in pasta_dest.glob("*.xlsx")
                    if not f.name.startswith(".")
                    and "Atualiza" not in f.name
                    and not str(f).endswith(".crdownload")]
            if xlsx:
                destino = pasta_dest / "DiaryLoadAutoTank.xlsx"
                try:
                    if str(xlsx[0]) != str(destino):
                        shutil.move(str(xlsx[0]), str(destino))
                        log.info("  Arquivo renomeado para DiaryLoadAutoTank.xlsx")
                    log.info(f"  DiaryLoadAutoTank baixado: {destino}")
                    return str(destino)
                except Exception as e:
                    log.warning(f"  Erro ao renomear: {e}")

            if i % 10 == 0 and i > 0:
                log.info(f"  Aguardando download... ({i*2}s)")

        log.warning("  Timeout aguardando download (3 min). Se no Chrome do robô aparecer "
                    "'baixar vários arquivos', clique em PERMITIR.")
        try:
            driver.save_screenshot(str(BASE_DIR / "newsat_timeout.png"))
        except Exception:
            pass
        return None

    except Exception as e:
        log.error(f"Erro no New SAT: {e}")
        try:
            if _driver_global:
                _driver_global.save_screenshot(str(BASE_DIR / "newsat_erro.png"))
        except Exception:
            pass
        _driver_global = None  # Chrome morreu, abre novo na próxima
        return None
    # Chrome não é fechado — reutilizado no próximo ciclo


# ══════════════════════════════════════════════════════════════════════════════
# 3b. EMAIL (Outlook Web) — ABERTURA / VIRADA de tanques -> campos amarelos
#     Usa uma segunda aba do MESMO Chrome do New SAT (mesmo login Microsoft).
# ══════════════════════════════════════════════════════════════════════════════
# MODO TESTE: só lê os emails e mostra no log o que FARIA. Não abre a planilha,
# não roda macro, não mexe no SAP e não grava/apaga o log da rede.
MODO_TESTE = bool(CFG.get("modo_teste", False))

CFG_EMAIL = CFG.get("email", {})
EMAIL_ATIVO = CFG_EMAIL.get("ativo", True)
URL_OUTLOOK = CFG_EMAIL.get("url", "https://outlook.cloud.microsoft/mail/")
# Termos digitados na pesquisa do Outlook (um por vez)
EMAIL_BUSCAS = CFG_EMAIL.get("buscas", ["Virada", "ABERTURA DE TANQUES", "Atualização de densidade"])
# Só aceita o resultado se o ASSUNTO bater com um destes padrões (sem acento, minúsculo)
# O ASSUNTO tem que COMEÇAR assim (aceita "[Externo]", "RES:", "ENC:" na frente) —
# evita pegar email que só cita "virada de" no texto (ex.: pedido de laudo)
_PREFIXO_ASSUNTO = r"^(?:\[externo\]\s*)?(?:(?:res|re|enc|fw|fwd)\s*:\s*)*(?:\[externo\]\s*)?"
EMAIL_PADROES_ASSUNTO = [_PREFIXO_ASSUNTO + p for p in CFG_EMAIL.get("padroes_assunto", [
    r"errata\s+abertura\s+de\s+tanque", r"virada\s+de", r"abertura\s+de\s+tanque",
    r"atualiza\w*\s+de\s+densidade",
])]
EMAIL_MAX_RESULTADOS = CFG_EMAIL.get("max_resultados_por_busca", 20)
# Teste usa um controle separado: o que foi "aplicado" no teste não conta pra produção
# Lê a tabela do PDF anexo (padrão). Se o PDF falhar, tenta o corpo do email.
EMAIL_USAR_PDF = CFG_EMAIL.get("usar_pdf", True)
CACHE_EMAILS = BASE_DIR / ("emails_processados_TESTE.json" if MODO_TESTE else "emails_processados.json")

# Linha da aba INICIO de cada produto (H = temperatura am., I = densidade am.)
LINHAS_AMARELO = {"S10": 7, "S500": 8, "GASOLINA": 9, "B100": 10,
                  "ANIDRO": 11, "HIDRATADO": 12, "JET": 13}

# Código do produto no PDF/email (sem espaço, maiúsculo) -> produto da planilha.
# Produto que não está aqui (ex.: OCB1) é ignorado.
MAPA_PRODUTOS = {
    "GAS": "GASOLINA", "GASA": "GASOLINA", "GASOLINA": "GASOLINA", "GASOLINAA": "GASOLINA",
    "S10": "S10", "ODS10": "S10", "DIESELS10": "S10",
    "S500": "S500", "ODS500": "S500", "DIESE": "S500", "DIESELS500": "S500",
    "B100": "B100", "BIODIESEL": "B100", "BDB100": "B100",
    "ANIDRO": "ANIDRO", "AEAC": "ANIDRO", "EAC": "ANIDRO", "ALCA": "ANIDRO", "ETANOLANIDRO": "ANIDRO",
    "HIDRATADO": "HIDRATADO", "AEHC": "HIDRATADO", "EHC": "HIDRATADO", "ALCH": "HIDRATADO",
    "ETANOLHIDRATADO": "HIDRATADO",
    "JET": "JET", "QAV": "JET", "QAV1": "JET", "JETA1": "JET",
}
MAPA_PRODUTOS.update({k.upper(): v for k, v in CFG_EMAIL.get("mapa_produtos", {}).items()})

# Linha da tabela "Controle de Qualidade":
# Data | Produto | Tanque | Cor | Aspecto | Dens. amostra | Temp. amostra | Dens. 20° | ...
# 20/10/2026 GAS 1 (2)INAM (1)LII 0,7170 21,5 0,7182 - 26,6 ...
_RE_LINHA_QUALIDADE = re.compile(
    r"(\d{2}/\d{2}/\d{2,4})\s*"               # data
    r"([A-Za-z][A-Za-z0-9 .\-]{0,15}?)\s+"   # produto (GAS, S10, OD S500...)
    r"(\d{1,3})\s+"                           # tanque
    r"[^/]{0,80}?"                            # cor + aspecto (não atravessa p/ próxima linha)
    r"(\d[.,]\d{4}|\d[.,]\d{3}(?=\s))\s*"     # densidade da amostra (0,7170)
    r"(\d{1,2}[.,]\d{1,2})\s*"                # temperatura da amostra (21,5)
    r"(\d[.,]\d{3,4})"                        # densidade a 20° (0,7182)
)

_DIAS_SEMANA = {"seg": 0, "ter": 1, "qua": 2, "qui": 3, "sex": 4, "sab": 5, "dom": 6}


def _sem_acento(txt: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", txt)
                   if unicodedata.category(c) != "Mn").lower()


def extrair_dados_qualidade(texto: str) -> list[dict]:
    """
    Lê a(s) linha(s) da tabela Controle de Qualidade num texto (corpo do
    email ou PDF). Retorna [{produto, tanque, temp, dens}], 1 por produto.
    """
    texto = re.sub(r"\s+", " ", texto.replace("\xa0", " "))
    por_produto = {}
    for m in _RE_LINHA_QUALIDADE.finditer(texto):
        codigo = re.sub(r"[\s.\-]", "", m.group(2)).upper()
        produto = MAPA_PRODUTOS.get(codigo)
        if not produto:
            log.info(f"    Produto '{m.group(2)}' não usado na planilha — ignorado.")
            continue
        try:
            fmt = "%d/%m/%Y" if len(m.group(1)) == 10 else "%d/%m/%y"
            data = datetime.strptime(m.group(1), fmt).date().isoformat()
        except ValueError:
            continue
        # Se o painel mostrar uma conversa com emails antigos juntos (mesmo assunto),
        # fica a linha com a DATA mais nova da tabela de cada produto
        if produto in por_produto and por_produto[produto]["data"] >= data:
            continue
        dens = round(float(m.group(4).replace(",", ".")) * 1000)   # 0,7170 -> 717
        temp = float(m.group(5).replace(",", "."))                  # 21,5
        por_produto[produto] = {"produto": produto, "tanque": m.group(3),
                                "temp": temp, "dens": dens, "data": data}
    return list(por_produto.values())


def _data_do_item(linhas: list[str], agora: datetime) -> datetime | None:
    """Converte o horário da lista do Outlook ('05:31', 'Seg, 18:33', '13/10/2026') em datetime."""
    for linha in linhas:
        s = _sem_acento(linha.strip())
        m = re.fullmatch(r"(\d{1,2}):(\d{2})", s)
        if m:
            return agora.replace(hour=int(m[1]), minute=int(m[2]), second=0, microsecond=0)
        m = re.fullmatch(r"(seg|ter|qua|qui|sex|sab|dom)\.?,?\s+(\d{1,2}):(\d{2})", s)
        if m:
            atras = (agora.weekday() - _DIAS_SEMANA[m[1]]) % 7 or 7
            d = agora - timedelta(days=atras)
            return d.replace(hour=int(m[2]), minute=int(m[3]), second=0, microsecond=0)
        m = re.fullmatch(r"(?:[a-z]+\.?,?\s+)?(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?(?:\s+(\d{1,2}):(\d{2}))?", s)
        if m:
            ano = int(m[3]) if m[3] else agora.year
            if ano < 100:
                ano += 2000
            try:
                return datetime(ano, int(m[2]), int(m[1]), int(m[4] or 0), int(m[5] or 0))
            except ValueError:
                continue
    return None


def _horario_efetivo(assunto_norm: str, recebido: datetime) -> datetime:
    """
    'Virada de GAS p/ TQ 01 as 05h30' -> 05:30 (o valor só vale a partir daí).
    Sem horário no assunto (abertura) -> horário de recebimento.
    """
    m = re.search(r"\bas\s+(\d{1,2})\s*(?:h|:)\s*(\d{2})?", assunto_norm)
    if not m or int(m[1]) > 23:
        return recebido
    eff = recebido.replace(hour=int(m[1]), minute=int(m[2] or 0))
    if eff < recebido - timedelta(hours=12):
        eff += timedelta(days=1)       # email 23:50 "às 00h30" -> dia seguinte
    elif eff > recebido + timedelta(hours=12):
        eff -= timedelta(days=1)
    return eff


def _carregar_cache() -> dict:
    try:
        cache = json.loads(CACHE_EMAILS.read_text(encoding="utf-8"))
    except Exception:
        cache = {}
    if cache.get("versao") != 4:   # controle antigo é descartado (reaplica o dia)
        cache = {"versao": 4}
    cache.setdefault("emails", {})
    cache.setdefault("ultimo_aplicado", {})
    return cache


def _salvar_cache(cache: dict):
    # Remove emails com mais de 3 dias, pro arquivo não crescer pra sempre
    limite = (datetime.now() - timedelta(days=3)).isoformat()
    cache["emails"] = {k: v for k, v in cache["emails"].items() if v.get("eff", "") >= limite}
    CACHE_EMAILS.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")


def _achar_primeiro_visivel(driver, seletores: list[str]):
    for sel in seletores:
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                if el.is_displayed():
                    return el
            except Exception:
                continue
    return None


def _campo_busca_outlook(driver):
    return _achar_primeiro_visivel(driver, [
        "#topSearchInput", "input[aria-label*='Pesquis']", "input[aria-label*='Search']",
        "input[placeholder*='Pesquis']",
    ])


def _abrir_outlook(driver) -> bool:
    _ir_para_aba(driver, "email")
    if "/mail" not in driver.current_url.lower():
        driver.get(URL_OUTLOOK)
        time.sleep(6)
    if not _ja_logado(driver):
        if not _aguardar_login_manual(driver):
            return False
        driver.get(URL_OUTLOOK)
        time.sleep(6)
    try:
        WebDriverWait(driver, 60).until(lambda d: _campo_busca_outlook(d) is not None)
        return True
    except Exception:
        log.error("  Outlook Web não carregou (campo de pesquisa não encontrado).")
        driver.save_screenshot(str(BASE_DIR / "outlook_erro.png"))
        return False


def _pesquisar_outlook(driver, termo: str) -> list:
    _fechar_dialogos_outlook(driver)
    campo = _campo_busca_outlook(driver)
    try:
        campo.click()
    except Exception:
        driver.execute_script("arguments[0].focus(); arguments[0].click();", campo)
    campo.send_keys(Keys.CONTROL, "a")
    campo.send_keys(Keys.DELETE)
    campo.send_keys(termo)
    campo.send_keys(Keys.ENTER)
    time.sleep(5)
    try:
        WebDriverWait(driver, 20).until(
            lambda d: d.find_elements(By.CSS_SELECTOR, "div[role='listbox'] div[role='option']"))
    except Exception:
        return []
    return driver.find_elements(By.CSS_SELECTOR, "div[role='listbox'] div[role='option']")


def _texto_corpo_email(driver) -> str:
    """Texto do email aberto no painel de leitura."""
    for _ in range(10):
        docs = [d for d in driver.find_elements(
                    By.CSS_SELECTOR,
                    "div[aria-label='Corpo da mensagem'], div[aria-label='Message body']")
                if d.is_displayed()]
        if not docs:
            docs = [d for d in driver.find_elements(By.CSS_SELECTOR, "div[role='document']")
                    if d.is_displayed()]
        if not docs:
            docs = [d for d in driver.find_elements(By.CSS_SELECTOR, "#ReadingPaneContainerId")
                    if d.is_displayed()]
        texto = "\n".join(d.get_attribute("innerText") or "" for d in docs)
        if texto.strip():
            return texto
        time.sleep(1)
    return ""


_JS_ANEXOS_PDF = r"""
// arguments[0] = borda direita do email clicado na lista de resultados.
// Só vale anexo À DIREITA disso (painel de leitura = email aberto).
const limite = arguments[0];
const visivel = e => e.offsetParent !== null && e.getBoundingClientRect().width > 0;
const ehPdf = e => /\.pdf/i.test((e.getAttribute('title') || '') + ' ' + (e.getAttribute('aria-label') || ''));
const todos = [...document.querySelectorAll('[title],[aria-label]')].filter(e => ehPdf(e) && visivel(e));
let painel = todos.filter(e => e.getBoundingClientRect().left >= limite);
if (!painel.length) {
    // plano B: cartão do anexo pelo nome (PFPL / CONTROLE DE QUALIDADE) no painel
    painel = [...document.querySelectorAll('div,button,span')]
        .filter(e => visivel(e) && e.getBoundingClientRect().left >= limite
                  && /PFPL|CONTROLE DE/i.test(e.innerText || '') && (e.innerText || '').length < 120);
}
painel.sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
return [todos.length, painel];
"""


def _texto_pdf_anexo(driver, limite_x: float = 0) -> str:
    """
    Baixa o PDF anexo DO EMAIL ABERTO (painel de leitura, à direita da lista) e devolve o texto.
    Só aceita PDF que chegou na pasta DEPOIS do clique em Baixar; apaga depois de ler.
    """
    pasta = Path(CFG["newsat"]["pasta_download"])
    try:
        # Espera downloads anteriores terminarem e limpa PDFs velhos da pasta
        for _ in range(15):
            if not list(pasta.glob("*.crdownload")):
                break
            time.sleep(1)
        for f in pasta.glob("PFPL*.pdf"):   # só PDFs do Pool — não mexe em outros PDFs da pasta
            try:
                f.unlink()
            except Exception:
                pass

        total, anexos = driver.execute_script(_JS_ANEXOS_PDF, limite_x)
        log.info(f"    PDFs na tela: {total} | no email aberto: {len(anexos)}")
        if not anexos:
            log.warning("    Nenhum anexo PDF no email aberto.")
            return ""
        anexos[0].click()
        time.sleep(4)
        botoes = [b for b in driver.find_elements(
                      By.XPATH,
                      "//button[@aria-label='Baixar' or @title='Baixar' or @aria-label='Download'"
                      " or normalize-space()='Baixar' or normalize-space()='Download']")
                  if b.is_displayed()]
        if not botoes:
            log.warning("    Botão 'Baixar' do anexo não encontrado.")
            return ""
        t0 = time.time() - 2
        botoes[0].click()
        arquivo = None
        for _ in range(30):
            time.sleep(1)
            novos = [f for f in pasta.glob("*.pdf") if f.stat().st_mtime >= t0]
            if novos and not list(pasta.glob("*.crdownload")):
                arquivo = max(novos, key=lambda f: f.stat().st_mtime)
                break
        if not arquivo:
            log.warning(f"    PDF não chegou em {pasta}.")
            return ""
        time.sleep(1)
        try:
            import pdfplumber
            with pdfplumber.open(str(arquivo)) as pdf:
                # O PDF do Pool tem uma linha "fantasma" (EHC/INCOLOR) desenhada quase na
                # mesma altura da linha real; com tolerância padrão as duas se misturam
                # ("GEHACS", "2200//1100//22002266"). y_tolerance pequeno separa as linhas.
                texto = "\n".join((pg.extract_text(y_tolerance=0.2) or "") for pg in pdf.pages)
                if not extrair_dados_qualidade(texto):
                    texto += "\n" + "\n".join((pg.extract_text() or "") for pg in pdf.pages)
        except ImportError:
            from pypdf import PdfReader
            texto = "\n".join((pg.extract_text() or "") for pg in PdfReader(str(arquivo)).pages)
        if not extrair_dados_qualidade(texto):
            dbg = BASE_DIR / f"debug_pdf_{datetime.now():%H%M%S}.txt"
            dbg.write_text(f"ARQUIVO: {arquivo.name}\n\n{texto}", encoding="utf-8")
            log.warning(f"    Texto do PDF não reconhecido — salvo em {dbg.name} (me mande esse arquivo)")
        try:
            arquivo.unlink()
        except Exception:
            pass
        return texto
    except Exception as e:
        log.warning(f"    Não consegui baixar/ler o PDF anexo: {e}")
        return ""
    finally:
        _fechar_dialogos_outlook(driver)


def _fechar_dialogos_outlook(driver):
    """Fecha visualização de anexo / qualquer diálogo aberto por cima do Outlook."""
    for _ in range(4):
        try:
            abertos = [d for d in driver.find_elements(
                           By.CSS_SELECTOR, ".fui-DialogSurface__backdrop, div[role='dialog']")
                       if d.is_displayed()]
        except Exception:
            abertos = []
        if not abertos:
            return
        try:
            fechar = [b for b in driver.find_elements(
                          By.XPATH, "//div[@role='dialog']//button[@aria-label='Fechar' or @title='Fechar'"
                                    " or @aria-label='Close' or @title='Close']")
                      if b.is_displayed()]
            if fechar:
                fechar[0].click()
            else:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except Exception:
                pass
        time.sleep(1)


def _conferir_com_assunto(assunto_norm: str, dados: list[dict]) -> list[dict]:
    """
    Virada: o PDF tem que ter o MESMO produto e tanque do assunto
    ("Virada de GAS p/ TQ 01" -> GASOLINA no TQ 1). Senão é PDF de outro email.
    Abertura/atualização: sem produto no assunto, passa o que veio (data já foi conferida).
    """
    m = re.search(r"virada\s+de\s+(.+?)\s+p\s*/?\s*tq\.?\s*0*(\d+)", assunto_norm)
    if not m:
        return dados
    produto = MAPA_PRODUTOS.get(re.sub(r"[\s.\-]", "", m[1]).upper())
    if not produto:
        return dados  # produto do assunto desconhecido — não dá pra conferir
    tq = m[2].lstrip("0")
    return [d for d in dados if d["produto"] == produto and d["tanque"].lstrip("0") == tq]


def coletar_emails_qualidade() -> tuple[dict, list]:
    """
    Pesquisa no Outlook Web os emails de abertura/virada/atualização das
    DIA ATUAL, lê os que ainda não foram lidos e devolve:
      pendentes = {produto: {temp, dens, tanque, eff, assunto}} (o mais recente de cada)
      chaves    = emails que entram nessa aplicação (marcados depois de gravar no Excel)
    """
    cache = _carregar_cache()
    agora = datetime.now()
    if not EMAIL_ATIVO:
        return {}, []

    driver = _obter_driver()
    if driver is None:
        log.error("  Chrome indisponível — emails não verificados.")
    else:
        try:
            log.info("Verificando emails de abertura/virada de tanque no Outlook...")
            if _abrir_outlook(driver):
                # SÓ emails do DIA ATUAL (a temperatura é diária: 01:00 -> 21:30).
                # Nada de ontem ou antes entra — nem pela lista, nem pelo assunto, nem pela tabela.
                limite = agora.replace(hour=0, minute=0, second=0, microsecond=0)
                data_minima = agora.date()
                ultimo_texto = ""
                vistos_rodada = set()
                for termo in EMAIL_BUSCAS:
                    itens = _pesquisar_outlook(driver, termo)[:EMAIL_MAX_RESULTADOS]
                    # 1º lê só a lista (os elementos ficam "velhos" depois de clicar)
                    candidatos = []
                    for it in itens:
                        try:
                            txt = it.text or ""
                        except Exception:
                            continue
                        linhas = [l for l in txt.splitlines() if l.strip()]
                        assunto = next((l for l in linhas if any(
                            re.search(p, _sem_acento(l.strip())) for p in EMAIL_PADROES_ASSUNTO)), None)
                        if not assunto:
                            continue
                        # Assunto com data (ex.: "ABERTURA DE TANQUES - 17/10/2026"): se for
                        # antiga, nem abre o email
                        m_data = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", assunto)
                        if m_data:
                            try:
                                d_assunto = date(int(m_data[3]), int(m_data[2]), int(m_data[1]))
                                if d_assunto < data_minima:
                                    continue
                            except ValueError:
                                pass
                        recebido = _data_do_item(linhas, agora)
                        if not recebido or recebido < limite:
                            continue
                        chave = f"{recebido:%Y-%m-%d %H:%M}|{linhas[0].strip()}|{assunto.strip()}"
                        if chave in vistos_rodada:
                            continue  # mesmo email aparece em "Principais" e "Todos os resultados"
                        vistos_rodada.add(chave)
                        reg = cache["emails"].get(chave)
                        if reg and (reg["itens"] or reg.get("tentativas", 0) >= 3):
                            continue  # já lido
                        candidatos.append((txt, chave, assunto, recebido))

                    for txt, chave, assunto, recebido in candidatos:
                        alvo = None
                        for it in driver.find_elements(By.CSS_SELECTOR, "div[role='listbox'] div[role='option']"):
                            try:
                                if it.text == txt:
                                    alvo = it
                                    break
                            except Exception:
                                continue
                        if alvo is None:
                            continue
                        try:
                            r = alvo.rect
                            limite_x = r["x"] + r["width"]
                        except Exception:
                            limite_x = 0
                        try:
                            alvo.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", alvo)
                        time.sleep(3)
                        log.info(f"    Abrindo: {assunto.strip()} (lista diz {recebido:%d/%m %H:%M})")
                        # Espera o painel trocar pro email clicado (não pegar anexo do anterior)
                        texto = ""
                        for _ in range(10):
                            texto = _texto_corpo_email(driver)
                            if texto and texto != ultimo_texto:
                                break
                            time.sleep(1)
                        ultimo_texto = texto
                        time.sleep(2)
                        dados = []
                        if EMAIL_USAR_PDF:
                            # Caminho principal: baixa o PDF anexo DESTE email e lê
                            dados = extrair_dados_qualidade(_texto_pdf_anexo(driver, limite_x))
                            if not dados:
                                log.warning(f"    PDF sem dados reconhecidos — tentando o corpo: {assunto.strip()}")
                        if not dados:
                            dados = extrair_dados_qualidade(texto)
                        if not dados:
                            trecho = re.sub(r"\s+", " ", texto)[:300]
                            log.warning(f"    Tabela não encontrada. Início do texto lido: {trecho!r}")
                        # Virada: PDF tem que bater com produto/tanque do assunto
                        if dados:
                            conferidos = _conferir_com_assunto(_sem_acento(assunto), dados)
                            if not conferidos:
                                veio = ", ".join(f"{d['produto']} TQ{d['tanque']} {d['data'][8:10]}/{d['data'][5:7]}"
                                                 for d in dados)
                                log.warning(f"    PDF NÃO confere com o assunto (veio: {veio}) — descartado: "
                                            f"{assunto.strip()}")
                            dados = conferidos
                        # Data da TABELA manda: linha que não é de hoje é descartada
                        antigos = [d for d in dados if d["data"] < data_minima.isoformat()]
                        dados = [d for d in dados if d["data"] >= data_minima.isoformat()]
                        if antigos and not dados:
                            log.info(f"    Ignorado — tabela é de {antigos[0]['data'][8:10]}/{antigos[0]['data'][5:7]}"
                                     f" (antigo): {assunto.strip()}")
                            cache["emails"][chave] = {"assunto": assunto.strip(), "eff": agora.isoformat(),
                                                      "itens": [], "tentativas": 99, "aplicado": True}
                            continue
                        if dados:
                            # Se a lista do Outlook errou o dia, usa o dia da tabela
                            d_tabela = datetime.fromisoformat(max(d["data"] for d in dados)).date()
                            if d_tabela != recebido.date():
                                recebido = datetime.combine(d_tabela, recebido.time())
                        assunto_norm = _sem_acento(assunto)
                        reg = cache["emails"].get(chave, {"tentativas": 0})
                        reg.update({
                            "assunto": assunto.strip(),
                            "recebido": recebido.isoformat(),
                            "eff": _horario_efetivo(assunto_norm, recebido).isoformat(),
                            "itens": dados,
                            "tentativas": reg.get("tentativas", 0) + 1,
                            "aplicado": reg.get("aplicado", False),
                        })
                        cache["emails"][chave] = reg
                        if dados:
                            resumo = ", ".join(f"{d['produto']} TQ{d['tanque']} {d['dens']} / {d['temp']:.1f}°C"
                                               for d in dados)
                            log.info(f"    Lido: {assunto.strip()} -> {resumo}")
                        else:
                            log.warning(f"    Sem dados reconhecidos em: {assunto.strip()}")
                _salvar_cache(cache)
        except Exception as e:
            log.error(f"  Erro ao ler emails no Outlook: {e}")
            try:
                driver.save_screenshot(str(BASE_DIR / "outlook_erro.png"))
            except Exception:
                pass
        finally:
            try:
                _ir_para_aba(driver, "newsat")
            except Exception:
                pass

    # Monta o que falta aplicar, em ordem cronológica (o mais novo vence)
    pendentes, chaves = {}, []
    ultimo = cache["ultimo_aplicado"]
    for chave, reg in sorted(cache["emails"].items(), key=lambda kv: kv[1].get("eff", "")):
        if reg.get("aplicado") or not reg.get("itens"):
            continue
        eff = datetime.fromisoformat(reg["eff"])
        if eff.date() != agora.date():
            continue  # só vale o que é do dia atual
        if eff > agora:
            log.info(f"    Aguardando horário da virada ({eff:%H:%M}): {reg['assunto']}")
            continue
        chaves.append(chave)
        for d in reg["itens"]:
            if d["produto"] in ultimo and ultimo[d["produto"]] >= reg["eff"]:
                continue  # já foi aplicado algo mais novo pra esse produto
            pendentes[d["produto"]] = {**d, "eff": reg["eff"], "assunto": reg["assunto"]}
    return pendentes, chaves


def marcar_emails_aplicados(pendentes: dict, chaves: list):
    cache = _carregar_cache()
    for chave in chaves:
        if chave in cache["emails"]:
            cache["emails"][chave]["aplicado"] = True
    for produto, d in pendentes.items():
        cache["ultimo_aplicado"][produto] = d["eff"]
    _salvar_cache(cache)


def _preencher_amarelos(wb, pendentes: dict):
    """Escreve temperatura (H) e densidade (I) da amostra na aba INICIO."""
    ws = wb.Worksheets("INICIO")
    for produto, d in pendentes.items():
        linha = LINHAS_AMARELO[produto]
        antes_t, antes_d = ws.Range(f"H{linha}").Value, ws.Range(f"I{linha}").Value
        ws.Range(f"H{linha}").Value = d["temp"]
        ws.Range(f"I{linha}").Value = d["dens"]
        log.info(f"  Amarelo {produto}: temp {antes_t} -> {d['temp']:.2f} | "
                 f"dens {antes_d} -> {d['dens']}  ({d['assunto']})")


# ══════════════════════════════════════════════════════════════════════════════
# 3c. LOG NA REDE (LOGs.xlsx no SharePoint) + JANELA DE FUNCIONAMENTO
# ══════════════════════════════════════════════════════════════════════════════
CFG_LOG_REDE = CFG.get("log_rede", {})
CFG_LOG_REDE.setdefault("caminho", "")  # caminho/URL da LOGs.xlsx compartilhada (config.json)
# TESTE COMPLETO: em modo teste NUNCA toca a rede nem a planilha real.
#  - "rede"     -> cópia local LOGs_TESTE.xlsx (na pasta do teste)
#  - planilha   -> cópia local COPIA_TESTE_Temperatura.xlsm (nome SEM "Atualiza
#                  Temperatura BPA", pra o robô de produção nunca confundir com a real)
TESTE_PLANILHA = BASE_DIR / CFG.get("teste_planilha", "COPIA_TESTE_Temperatura.xlsm")
if MODO_TESTE:
    CFG_LOG_REDE["caminho"] = str(BASE_DIR / CFG.get("teste_log_rede", "LOGs_TESTE.xlsx"))
    CFG_LOG_REDE["limpar_log_local"] = False
CFG_JANELA = CFG.get("janela", {})


def _hora_cfg(chave: str, padrao: str) -> dtime:
    h, m = CFG_JANELA.get(chave, padrao).split(":")
    return dtime(int(h), int(m))


HORA_INICIO = _hora_cfg("inicio", "01:00")
HORA_FIM = _hora_cfg("fim", "21:30")
HORA_LIMPEZA_REDE = CFG_JANELA.get("limpar_log_rede", "23:59")


def dentro_da_janela(agora: datetime | None = None) -> bool:
    t = (agora or datetime.now()).time()
    return HORA_INICIO <= t <= HORA_FIM


def _abrir_log_rede(excel):
    """Abre (ou reaproveita, se já estiver aberta) a LOGs.xlsx da rede. Retorna (wb, abriu_agora)."""
    caminho = CFG_LOG_REDE.get("caminho", "").strip()
    if not caminho:
        raise RuntimeError("config.json -> log_rede.caminho não preenchido")
    nome = urllib.parse.unquote(caminho.split("?")[0].replace("\\", "/").rsplit("/", 1)[-1]).lower()
    for i in range(1, excel.Workbooks.Count + 1):
        wb = excel.Workbooks(i)
        if wb.Name.lower() == nome:
            return wb, False
    return excel.Workbooks.Open(caminho), True


def _aba_log_rede(wb):
    aba = CFG_LOG_REDE.get("aba", "Logs")
    try:
        return wb.Worksheets(aba)
    except Exception:
        return wb.Worksheets(1)


def gravar_log_rede(excel, wb_temp, usar_apoio: bool = False):
    """
    Copia as 2 linhas que a macro ACABOU de gravar na aba LOG da planilha
    (Dep. 01 + POOL) para a LOGs.xlsx da rede — a rede fica igual ao LOG local.
    usar_apoio=True (teste): copia do bloco APOIO!A31:H32 (a macro não rodou).
    """
    if not CFG_LOG_REDE.get("ativo", True):
        return
    if usar_apoio:
        excel.Calculate()
        dados = wb_temp.Worksheets("APOIO").Range("A31:H32").Value2
        return _gravar_linhas_log_rede(excel, wb_temp, dados)
    ws_log = wb_temp.Worksheets("LOG")
    ult = ws_log.Cells(ws_log.Rows.Count, 8).End(-4162).Row   # última linha com DEP. (col. H)
    if ult < 8:
        raise RuntimeError("aba LOG da planilha está vazia — nada para copiar")
    # A macro grava Dep. 01 (com data/hora) e POOL logo abaixo. Se a última linha
    # não for o POOL com a de cima tendo data, cai pro bloco de apoio (mesmos valores).
    if str(ws_log.Range(f"H{ult}").Value).strip().upper() == "POOL" and ws_log.Range(f"A{ult - 1}").Value:
        dados = ws_log.Range(f"A{ult - 1}:H{ult}").Value2
        log.info(f"  Copiando LOG local (linhas {ult - 1}-{ult}) para a rede.")
    else:
        dados = wb_temp.Worksheets("APOIO").Range("A31:H32").Value2
        log.info("  LOG local fora do padrão — copiando do bloco APOIO!A31:H32.")
    _gravar_linhas_log_rede(excel, wb_temp, dados)


def _gravar_linhas_log_rede(excel, wb_temp, dados):
    wb_rede, abriu = _abrir_log_rede(excel)
    try:
        ws = _aba_log_rede(wb_rede)
        linha_ini = int(CFG_LOG_REDE.get("linha_inicial", 4))
        ultima = ws.Cells(ws.Rows.Count, 8).End(-4162).Row   # -4162 = xlUp, coluna H (DEP.)
        prox = max(ultima + 1, linha_ini)
        ws.Range(f"A{prox}:H{prox + 1}").Value2 = dados
        # Formato da data: copia da linha de cima (já certa). Sem linha de cima, usa o
        # formato LOCAL (Excel em português: "aaaa" = ano; "yyyy" vira texto literal)
        try:
            if prox - 2 >= linha_ini:
                ws.Range(f"A{prox}").NumberFormat = ws.Range(f"A{prox - 2}").NumberFormat
            else:
                ws.Range(f"A{prox}").NumberFormatLocal = "dd/mm/aaaa hh:mm"
        except Exception:
            ws.Range(f"A{prox}").NumberFormatLocal = "dd/mm/aaaa hh:mm"
        wb_rede.Save()
        log.info(f"  {'[TESTE] Log gravado na CÓPIA ' + Path(CFG_LOG_REDE['caminho']).name if MODO_TESTE else 'Log gravado na rede'}"
                 f" (linhas {prox}-{prox + 1}).")
    finally:
        if abriu:
            wb_rede.Close(SaveChanges=False)
        try:
            wb_temp.Activate()
        except Exception:
            pass


def limpar_log_rede():
    """Às 23:59 apaga os logs do dia na LOGs.xlsx da rede (mantém título/cabeçalho)."""
    if not CFG_LOG_REDE.get("ativo", True):
        return
    if MODO_TESTE and not Path(CFG_LOG_REDE["caminho"]).exists():
        log.info("[TESTE] 23:59 — sem LOGs_TESTE.xlsx na pasta, nada a limpar.")
        return
    if not _ciclo_lock.acquire(timeout=900):
        log.error("  Limpeza do log da rede: ciclo travado há 15 min — limpeza não feita.")
        return
    try:
        log.info("═══ LIMPEZA DO LOG DA REDE ═══")
        excel = _dispatch_excel()
        wb_rede, abriu = _abrir_log_rede(excel)
        try:
            ws = _aba_log_rede(wb_rede)
            linha_ini = int(CFG_LOG_REDE.get("linha_inicial", 4))
            ultima = max(ws.Cells(ws.Rows.Count, 1).End(-4162).Row,
                         ws.Cells(ws.Rows.Count, 8).End(-4162).Row)
            if ultima >= linha_ini:
                ws.Range(f"A{linha_ini}:H{ultima}").ClearContents()
                wb_rede.Save()
                log.info(f"  Log da rede apagado (linhas {linha_ini}-{ultima}).")
            else:
                log.info("  Log da rede já estava vazio.")
        finally:
            if abriu:
                wb_rede.Close(SaveChanges=False)

        if CFG_LOG_REDE.get("limpar_log_local", False):
            wb = _selecionar_planilha_correta(excel)
            ws_local = wb.Worksheets("LOG")
            ultima = ws_local.Cells(ws_local.Rows.Count, 8).End(-4162).Row
            if ultima >= 7:
                ws_local.Range(f"A7:H{ultima}").ClearContents()
                wb.Save()
                log.info("  Aba LOG da planilha local também apagada.")
    except Exception as e:
        log.error(f"  Erro ao limpar o log da rede: {e}")
    finally:
        _ciclo_lock.release()


# ══════════════════════════════════════════════════════════════════════════════
# 4. EXCEL — preencher campos amarelos e rodar macro (idêntico ao antigo,
#    + verificação de SAP conectado antes de rodar, pra falhar rápido e claro)
# ══════════════════════════════════════════════════════════════════════════════
def _aceitar_popup_sap():
    """Pressiona Enter para aceitar popup do SAP que pede permissão de acesso."""
    try:
        def callback(hwnd, extra):
            titulo = win32gui.GetWindowText(hwnd)
            if any(x in titulo for x in ["SAP", "Script", "Allow", "Permitir", "Macro"]):
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        win32gui.EnumWindows(callback, None)
    except Exception as e:
        log.warning(f"  Não foi possível tratar popup SAP: {e}")


def _fechar_popup_sap_gui():
    """Fecha qualquer popup/dialog aberto no SAP antes de rodar a macro."""
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        conn = app.Children(0)
        session = conn.Children(0)
        try:
            wnd1 = session.findById("wnd[1]")
            if wnd1:
                session.findById("wnd[1]").sendVKey(0)
                time.sleep(0.5)
        except Exception:
            pass
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
    except Exception:
        pass


def _verificar_sap_conectado() -> bool:
    """Verifica se o SAP GUI está aberto e com sessão ativa antes de tentar a macro."""
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        if app.Children.Count == 0:
            log.warning("  SAP: nenhuma conexão ativa.")
            return False
        conn = app.Children(0)
        if conn.Children.Count == 0:
            log.warning("  SAP: conexão sem sessões ativas.")
            return False
        session = conn.Children(0)
        _ = session.findById("wnd[0]")
        return True
    except Exception as e:
        log.warning(f"  SAP não está conectado: {e}")
        return False


def _nome_chave_planilha() -> str:
    """
    Deriva a 'palavra-chave' de busca a partir do nome configurado em config.json,
    removendo sufixo de versão tipo ' rev2' — assim continua reconhecendo a
    planilha mesmo se o número da revisão mudar (rev2, rev3, etc.).
    Ex.: 'Atualiza Temperatura BPA rev2.xlsm' -> 'atualiza temperatura bpa'
    """
    caminho = CFG["planilha"]["caminho"]
    stem = Path(caminho).stem
    chave = re.sub(r"\s*rev\s*\d*\s*$", "", stem, flags=re.IGNORECASE).strip()
    return chave.lower()


def _selecionar_planilha_correta(excel):
    """
    Localiza, entre TODAS as planilhas já abertas no Excel, a que tem nome
    parecido com a planilha de temperatura configurada (ex: 'Atualiza
    Temperatura BPA'). Isso evita que a macro rode ou salve em cima de outra
    planilha do usuário que esteja aberta ao mesmo tempo.
    Se nenhuma planilha aberta bater com o nome, abre do caminho configurado.
    """
    chave = _nome_chave_planilha()
    caminho = CFG["planilha"]["caminho"]

    for i in range(1, excel.Workbooks.Count + 1):
        try:
            wb_candidata = excel.Workbooks(i)
            if chave in wb_candidata.Name.lower():
                log.info(f"  Planilha correta encontrada aberta: {wb_candidata.Name}")
                wb_candidata.Activate()
                return wb_candidata
        except Exception:
            continue

    log.info(f"  Nenhuma planilha aberta com nome parecido a '{chave}'. Abrindo: {caminho}")
    wb = excel.Workbooks.Open(caminho)
    wb.Activate()
    return wb


def _abrir_planilha_teste(excel):
    """Abre (ou reaproveita) a CÓPIA de teste pelo nome EXATO — nunca a planilha real."""
    nome = TESTE_PLANILHA.name.lower()
    for i in range(1, excel.Workbooks.Count + 1):
        wb = excel.Workbooks(i)
        if wb.Name.lower() == nome:
            return wb
    return excel.Workbooks.Open(str(TESTE_PLANILHA))


def _dispatch_excel():
    """
    Cria o objeto Excel.Application sem passar pelo cache gen_py do pywin32.

    win32com.client.Dispatch() tenta gerar/usar wrappers "makepy" em disco
    (win32com.gen_py). Se esse cache ficar corrompido — script morto no meio
    de uma geração, duas instâncias gerando ao mesmo tempo, Office atualizado
    no meio do caminho — toda chamada seguinte quebra com:
        module 'win32com.gen_py.<clsid>' has no attribute 'CLSIDToPackageMap'
    dynamic.Dispatch fala com o Excel via IDispatch puro (late binding), sem
    gerar nem ler nada em win32com.gen_py — não tem cache pra corromper, então
    esse erro deixa de poder acontecer por essa via. Mantém um fallback de
    autolimpeza só por segurança, caso alguma outra chamada no processo já
    tenha deixado o cache num estado ruim.
    """
    try:
        return win32com.client.dynamic.Dispatch("Excel.Application")
    except AttributeError as e:
        if "CLSIDToPackageMap" not in str(e):
            raise
        log.warning("  Cache gen_py corrompido detectado — limpando e tentando de novo...")
        for nome_modulo in [k for k in list(sys.modules) if k.startswith("win32com.gen_py")]:
            del sys.modules[nome_modulo]
        shutil.rmtree(win32com.__gen_path__, ignore_errors=True)
        return win32com.client.dynamic.Dispatch("Excel.Application")


def rodar_macro_somente(pendentes: dict | None = None, chaves: list | None = None,
                        gravar_rede: bool = False) -> bool:
    """
    1. Se veio dado novo de email, preenche os campos amarelos (e salva).
    2. Roda a macro (SAP) com os dados da planilha.
    3. Se deu certo e gravar_rede=True, copia o log para a LOGs.xlsx da rede.
    """
    global _macro_rodando
    macro = CFG["planilha"]["macro"]

    excel = None
    wb = None
    try:
        excel = _dispatch_excel()
        excel.Visible = True
        excel.DisplayAlerts = False
        wb = _selecionar_planilha_correta(excel)
    except Exception as e:
        log.error(f"Erro ao abrir a planilha: {e}")
        return False

    # Campos amarelos — independe do SAP estar no ar
    if pendentes:
        try:
            _preencher_amarelos(wb, pendentes)
            wb.Save()
            marcar_emails_aplicados(pendentes, chaves or [])
        except Exception as e:
            log.error(f"  Erro ao preencher campos amarelos: {e} — tenta de novo no próximo ciclo.")

    log.info(f"Rodando macro com dados atuais: {macro}")
    if not _verificar_sap_conectado():
        log.error("  SAP indisponível. Macro abortada — será tentada no próximo ciclo.")
        return False

    _macro_rodando = True
    try:
        # Fecha popups e tenta rodar com 1 retry
        _fechar_popup_sap_gui()
        _aceitar_popup_sap()
        for tentativa in range(2):
            try:
                excel.Application.Run(f"'{wb.Name}'!{macro}")
                wb.Save()
                log.info("  Macro executada com sucesso.")
                if gravar_rede:
                    try:
                        gravar_log_rede(excel, wb)
                    except Exception as e:
                        log.warning(f"  Não foi possível gravar o log na rede: {e}")
                return True
            except Exception as e:
                if tentativa == 0:
                    log.warning(f"  Macro falhou, aguardando 10s e tentando novamente... ({e})")
                    _fechar_popup_sap_gui()
                    time.sleep(10)
                else:
                    log.error(f"  Macro falhou na segunda tentativa: {e}")
                    return False
    except Exception as e:
        log.error(f"Erro ao rodar macro: {e}")
        return False
    finally:
        _macro_rodando = False
    # Não fecha o Excel — mantém aberto para o próximo ciclo


# ══════════════════════════════════════════════════════════════════════════════
# 5. RESET DIÁRIO — limpa log às 00:00
# ══════════════════════════════════════════════════════════════════════════════
def reset_diario():
    log.info("═══ RESET DIÁRIO ═══ Iniciando novo ciclo de 24h")

    # Reconfigura logger com data nova
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    configurar_logger()
    log.info("Log reiniciado para o novo dia.")


# ══════════════════════════════════════════════════════════════════════════════
# 6. CICLO PRINCIPAL (idêntico ao antigo + atualiza o painel Tkinter no fim)
# ══════════════════════════════════════════════════════════════════════════════
def ciclo_hora(forcado: bool = False):
    global _erros_consecutivos

    no_horario = dentro_da_janela()
    if not no_horario and not forcado:
        log.info(f"Fora do horário de funcionamento ({HORA_INICIO:%H:%M}–{HORA_FIM:%H:%M}) — ciclo pulado.")
        if _ui_app:
            _ui_app.root.after(0, _ui_app.mostrar_fora_janela)
        return

    if not _ciclo_lock.acquire(blocking=False):
        log.warning("  Já existe um ciclo em andamento (agendado ou forçado) — ignorando esta chamada.")
        return

    try:
        log.info("━━━ INÍCIO DO CICLO HORÁRIO ━━━")
        inicio = datetime.now()

        # 1. Emails de abertura/virada -> valores dos campos amarelos
        pendentes, chaves = {}, []
        try:
            pendentes, chaves = coletar_emails_qualidade()
            if not pendentes:
                log.info("  Nenhuma abertura/virada nova para aplicar.")
        except Exception as e:
            log.error(f"  Erro na verificação de emails: {e}")

        if MODO_TESTE:
            # Nada de planilha real, macro, SAP ou rede — só cópias na pasta do teste
            if pendentes:
                log.info("  [TESTE] Colocaria no AMARELO (aba INICIO):")
                for produto, d in sorted(pendentes.items(), key=lambda kv: LINHAS_AMARELO[kv[0]]):
                    log.info(f"  [TESTE]   {produto:<9} temp {d['temp']:.2f} | dens {d['dens']} "
                             f"(TQ {d['tanque']}, vale desde {d['eff'][11:16]}) — {d['assunto']}")
            sucesso = True
            if TESTE_PLANILHA.exists():
                try:
                    excel = _dispatch_excel()
                    excel.Visible = True
                    excel.DisplayAlerts = False
                    wb = _abrir_planilha_teste(excel)
                    if pendentes:
                        _preencher_amarelos(wb, pendentes)
                        wb.Save()
                        log.info(f"  [TESTE] Amarelo preenchido na CÓPIA {TESTE_PLANILHA.name}.")
                    if no_horario:
                        gravar_log_rede(excel, wb, usar_apoio=True)
                except Exception as e:
                    log.error(f"  [TESTE] Erro na cópia da planilha/log: {e}")
                    sucesso = False
            if pendentes and sucesso:
                marcar_emails_aplicados(pendentes, chaves)   # só no controle de TESTE
            log.info("  [TESTE] Macro/SAP NÃO executados. Planilha real e rede NÃO tocadas.")
        else:
            # 2. Baixa DiaryLoadAutoTank do New SAT
            arquivo_diary = baixar_diary_newsat()

            # 3. Preenche amarelos + roda macro + log na rede (só dentro do horário)
            if arquivo_diary:
                sucesso = rodar_macro_somente(pendentes, chaves, gravar_rede=no_horario)
            else:
                # Sem o DiaryLoadAutoTank a macro abre um MsgBox que TRAVA o Excel —
                # então não roda. O amarelo ainda é preenchido (se houver email novo).
                log.error("  DiaryLoadAutoTank não baixou — macro NÃO executada neste ciclo.")
                if pendentes:
                    try:
                        excel = _dispatch_excel()
                        wb = _selecionar_planilha_correta(excel)
                        _preencher_amarelos(wb, pendentes)
                        wb.Save()
                        marcar_emails_aplicados(pendentes, chaves)
                    except Exception as e:
                        log.error(f"  Erro ao preencher campos amarelos: {e}")
                sucesso = False
        status = "OK" if sucesso else "ERRO"

        duracao = (datetime.now() - inicio).seconds
        log.info(f"━━━ FIM DO CICLO | Status: {status} | Duração: {duracao}s ━━━\n")

        # Atualiza painel da interface, se estiver rodando com Tkinter
        if _ui_app:
            _ui_app.root.after(0, lambda: _ui_app.atualizar_status(status, duracao))

        # Avisa o operador via notificação do Windows se o ciclo falhou —
        # notifica no 1º erro e depois só a cada N erros seguidos, pra não
        # floodar o operador se o SAP ficar fora do ar por horas.
        if status == "ERRO":
            _erros_consecutivos += 1
            if _erros_consecutivos == 1 or _erros_consecutivos % NOTIFICAR_A_CADA_N_ERROS == 0:
                _notificar_erro_operador(
                    f"A temperatura não foi atualizada ({_erros_consecutivos}ª falha seguida). "
                    "Verifique o SAP/painel."
                )
        else:
            _erros_consecutivos = 0
    finally:
        _ciclo_lock.release()


# ══════════════════════════════════════════════════════════════════════════════
# 7. MANTER SESSÕES VIVAS (a cada 5 minutos)
#    Igual ao antigo, só pulando o SAP se a macro estiver rodando naquele
#    momento (evita interferir e derrubar a macro em execução).
# ══════════════════════════════════════════════════════════════════════════════
def manter_sessoes_vivas():
    """Acessa New SAT e SAP rapidamente só para evitar timeout por inatividade."""
    global _driver_global
    if MODO_TESTE:
        return  # teste só usa o Outlook (que se mantém sozinho) — não encosta em New SAT/SAP
    # Ciclo (agendado ou forçado) usando o Chrome/Excel agora — não disputa com ele
    if _ciclo_lock.locked():
        log.info("── Keep-alive pulado — ciclo em andamento ──")
        return
    log.info("── Keep-alive: mantendo sessões ativas ──")

    # New SAT — abre o relatório para manter sessão ativa
    try:
        if _driver_global is not None:
            try:
                _driver_global.title  # confere se o Chrome ainda está vivo
            except Exception:
                log.warning("  New SAT: Chrome morreu. Recriando no próximo ciclo...")
                _driver_global = None

        if _driver_global is not None:
            _ir_para_aba(_driver_global, "newsat")
            _driver_global.get(URL_RELATORIO)
            time.sleep(2)
            # Se caiu na tela de login, marca para refazer login no próximo ciclo
            if not _ja_logado(_driver_global):
                if not _aguardar_login_manual(_driver_global):
                    log.error("  Chrome será recriado no próximo ciclo.")
                    _driver_global = None
            else:
                log.info("  New SAT: sessão mantida.")
        else:
            log.info("  New SAT: Chrome não aberto, pulando keep-alive.")
    except Exception as e:
        log.warning(f"  New SAT keep-alive falhou: {e}")
        _driver_global = None

    # SAP — pula se a macro estiver rodando agora
    if _macro_rodando:
        log.info("  SAP keep-alive pulado — macro em execução.")
        return
    if MODO_TESTE:
        return  # teste não encosta no SAP, nem pra leitura

    # SAP — leitura passiva do texto da janela principal. Não navega pra
    # nenhuma transação (evita tirar o usuário da tela em que está trabalhando).
    # Só o fato de acessar uma propriedade da sessão já mantém a conexão viva.
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        conn = app.Children(0)
        sessao = conn.Children(0)
        try:
            titulo = sessao.findById("wnd[0]").Text
        except Exception:
            titulo = "(sem titulo)"
        log.info(f"  SAP: keep-alive OK (leitura passiva — tela: {titulo}).")
    except Exception as e:
        log.warning(f"  SAP keep-alive falhou: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# 8. AGENDAMENTO
# ══════════════════════════════════════════════════════════════════════════════
def _registrar_jobs():
    schedule.clear()
    intervalo = CFG["automacao"]["intervalo_minutos"]
    schedule.every(intervalo).minutes.do(ciclo_hora)
    schedule.every(5).minutes.do(manter_sessoes_vivas)
    schedule.every().day.at("00:00").do(reset_diario)
    schedule.every().day.at(HORA_LIMPEZA_REDE).do(limpar_log_rede)
    # Primeiro ciclo do dia logo na abertura da janela (01:00), sem esperar o intervalo
    schedule.every().day.at(HORA_INICIO.strftime("%H:%M")).do(ciclo_hora)


# ══════════════════════════════════════════════════════════════════════════════
# INTERFACE GRÁFICA — Tkinter (sem conexão externa)
# ══════════════════════════════════════════════════════════════════════════════
_log_queue = queue.Queue()

class LogHandler(logging.Handler):
    """Coloca logs numa fila thread-safe — a UI lê e exibe a cada 100ms."""
    def emit(self, record):
        try:
            _log_queue.put_nowait((record.levelname, self.format(record)))
        except Exception:
            pass


class AutoBPAApp:
    def __init__(self, root):
        global _ui_app
        _ui_app = self
        self.root = root
        self.root.title("AutoBPA Monitor — Terminal" + ("  [MODO TESTE]" if MODO_TESTE else ""))
        self.root.geometry("900x600")
        self.root.configure(bg="#1a1a2e")
        self.root.resizable(True, True)

        self._rodando = False
        self._thread = None
        self._ciclos_ok = 0
        self._ciclos_erro = 0
        self._proximo_ciclo = None

        self._construir_ui()
        self._atualizar_relogio()

    def _construir_ui(self):
        # Cabeçalho
        frame_topo = tk.Frame(self.root, bg="#16213e", pady=10)
        frame_topo.pack(fill="x")
        tk.Label(
            frame_topo, text="AutoBPA Monitor" + ("  — MODO TESTE (não mexe no SAP)" if MODO_TESTE else ""),
            font=("Arial", 18, "bold"), fg="#f39c12" if MODO_TESTE else "#00d4ff", bg="#16213e"
        ).pack(side="left", padx=20)
        self.lbl_relogio = tk.Label(
            frame_topo, text="", font=("Arial", 12), fg="#aaaaaa", bg="#16213e"
        )
        self.lbl_relogio.pack(side="right", padx=20)

        # Status + Botões
        frame_controle = tk.Frame(self.root, bg="#1a1a2e", pady=10)
        frame_controle.pack(fill="x", padx=20)
        self.lbl_status = tk.Label(
            frame_controle, text="Aguardando inicio",
            font=("Arial", 13, "bold"), fg="#aaaaaa", bg="#1a1a2e"
        )
        self.lbl_status.pack(side="left")
        self.btn_parar = tk.Button(
            frame_controle, text="Parar", command=self._parar,
            font=("Arial", 11, "bold"), bg="#c0392b", fg="white",
            relief="flat", padx=15, pady=5, cursor="hand2", state="disabled"
        )
        self.btn_parar.pack(side="right", padx=5)
        self.btn_iniciar = tk.Button(
            frame_controle, text="Iniciar", command=self._iniciar,
            font=("Arial", 11, "bold"), bg="#27ae60", fg="white",
            relief="flat", padx=15, pady=5, cursor="hand2"
        )
        self.btn_iniciar.pack(side="right", padx=5)
        self.btn_forcar = tk.Button(
            frame_controle, text="Forçar Ciclo", command=self._forcar_ciclo,
            font=("Arial", 11, "bold"), bg="#2980b9", fg="white",
            relief="flat", padx=15, pady=5, cursor="hand2", state="disabled"
        )
        self.btn_forcar.pack(side="right", padx=5)

        # Métricas
        frame_metricas = tk.Frame(self.root, bg="#16213e", pady=8)
        frame_metricas.pack(fill="x", padx=20, pady=(0, 5))
        self.lbl_ok = tk.Label(
            frame_metricas, text="Ciclos OK: 0",
            font=("Arial", 10), fg="#2ecc71", bg="#16213e", padx=15
        )
        self.lbl_ok.pack(side="left")
        self.lbl_erro = tk.Label(
            frame_metricas, text="Erros: 0",
            font=("Arial", 10), fg="#e74c3c", bg="#16213e", padx=15
        )
        self.lbl_erro.pack(side="left")
        self.lbl_proximo = tk.Label(
            frame_metricas, text="Proximo ciclo: --:--",
            font=("Arial", 10), fg="#f39c12", bg="#16213e", padx=15
        )
        self.lbl_proximo.pack(side="left")
        self.lbl_ultimo = tk.Label(
            frame_metricas, text="Ultimo ciclo: --",
            font=("Arial", 10), fg="#aaaaaa", bg="#16213e", padx=15
        )
        self.lbl_ultimo.pack(side="left")

        # Caixa de log
        frame_log = tk.Frame(self.root, bg="#1a1a2e")
        frame_log.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        tk.Label(
            frame_log, text="Log em tempo real",
            font=("Arial", 10, "bold"), fg="#00d4ff", bg="#1a1a2e"
        ).pack(anchor="w", pady=(0, 3))
        self.log_box = scrolledtext.ScrolledText(
            frame_log, height=20, font=("Courier New", 9),
            bg="#0d0d0d", fg="#00ff41", insertbackground="white",
            state="disabled", relief="flat", padx=8, pady=8
        )
        self.log_box.pack(fill="both", expand=True)
        self.log_box.tag_config("info",    foreground="#00ff41")
        self.log_box.tag_config("warning", foreground="#f39c12")
        self.log_box.tag_config("error",   foreground="#e74c3c")

        # Conecta handler de log à fila thread-safe
        handler = LogHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logging.getLogger().addHandler(handler)

        # Inicia loop de leitura da fila de logs (100ms) — thread-safe
        self._processar_fila_log()

        # Rodapé
        tk.Label(
            self.root, text="Automação de Temperatura — Terminal de Combustíveis",
            font=("Arial", 8), fg="#555555", bg="#1a1a2e"
        ).pack(pady=(0, 5))

    def _processar_fila_log(self):
        """Lê a fila de logs e insere na caixa de texto — roda na thread principal a cada 100ms."""
        try:
            while True:
                nivel, msg = _log_queue.get_nowait()
                self.log_box.config(state="normal")
                tag = "info"
                if nivel == "WARNING":
                    tag = "warning"
                elif nivel == "ERROR":
                    tag = "error"
                self.log_box.insert(tk.END, msg + "\n", tag)
                self.log_box.see(tk.END)
                self.log_box.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._processar_fila_log)

    def _atualizar_relogio(self):
        self.lbl_relogio.config(text=datetime.now().strftime("%d/%m/%Y  %H:%M:%S"))
        if self._rodando and self._proximo_ciclo:
            restante = max(0, int((self._proximo_ciclo - datetime.now()).total_seconds()))
            mins = restante // 60
            segs = restante % 60
            self.lbl_proximo.config(text=f"Proximo ciclo: {mins:02d}:{segs:02d}")
        self.root.after(1000, self._atualizar_relogio)

    def atualizar_status(self, status: str, duracao: int):
        hora = datetime.now().strftime("%H:%M")
        if status == "OK":
            self._ciclos_ok += 1
            self.lbl_status.config(
                text=f"RODANDO — Ultimo ciclo OK as {hora} ({duracao}s)",
                fg="#2ecc71"
            )
        else:
            self._ciclos_erro += 1
            self.lbl_status.config(
                text=f"ERRO no ultimo ciclo as {hora}",
                fg="#e74c3c"
            )
        self.lbl_ok.config(text=f"Ciclos OK: {self._ciclos_ok}")
        self.lbl_erro.config(text=f"Erros: {self._ciclos_erro}")
        self.lbl_ultimo.config(text=f"Ultimo ciclo: {hora}")

        intervalo = CFG["automacao"]["intervalo_minutos"]
        self._proximo_ciclo = datetime.now() + __import__("datetime").timedelta(minutes=intervalo)

    def mostrar_fora_janela(self):
        self.lbl_status.config(
            text=f"AGUARDANDO — fora do horário ({HORA_INICIO:%H:%M}–{HORA_FIM:%H:%M})",
            fg="#f39c12"
        )
        intervalo = CFG["automacao"]["intervalo_minutos"]
        self._proximo_ciclo = datetime.now() + timedelta(minutes=intervalo)

    def _loop_automacao(self):
        # Inicializa COM para esta thread (necessário para win32com em threads)
        pythoncom.CoInitialize()
        log.info("╔══════════════════════════════════════════╗")
        log.info("║  AUTOMAÇÃO TEMPERATURA BPA - INICIANDO   ║")
        log.info("╚══════════════════════════════════════════╝")

        _impedir_suspensao()
        _registrar_jobs()

        intervalo = CFG["automacao"]["intervalo_minutos"]
        self._proximo_ciclo = datetime.now() + __import__("datetime").timedelta(minutes=intervalo)

        try:
            ciclo_hora()
        except Exception as e:
            log.error(f"Erro no ciclo inicial: {e}")

        while self._rodando:
            try:
                schedule.run_pending()
            except Exception as e:
                log.error(f"  Erro no loop: {e}. Continuando...")
            time.sleep(30)

        log.info("Automacao pausada.")

    def _iniciar(self):
        if not self._rodando:
            self._rodando = True
            self.btn_iniciar.config(state="disabled")
            self.btn_parar.config(state="normal")
            self.btn_forcar.config(state="normal")
            self.lbl_status.config(text="Iniciando...", fg="#f39c12")
            self._thread = threading.Thread(target=self._loop_automacao, daemon=True)
            self._thread.start()

    def _parar(self):
        if self._rodando:
            self._rodando = False
            self.btn_iniciar.config(state="normal")
            self.btn_parar.config(state="disabled")
            self.btn_forcar.config(state="disabled")
            self.lbl_status.config(text="Parado", fg="#aaaaaa")
            _liberar_suspensao()

    def _forcar_ciclo(self):
        """Dispara um ciclo imediatamente fora do agendamento normal."""
        if self._rodando:
            def _executar():
                pythoncom.CoInitialize()  # necessário: esta thread é diferente da do loop principal
                try:
                    log.info("=== CICLO FORÇADO MANUALMENTE ===")
                    ciclo_hora(forcado=True)
                finally:
                    pythoncom.CoUninitialize()
            threading.Thread(target=_executar, daemon=True).start()


if __name__ == "__main__":
    _garantir_instancia_unica()
    root = tk.Tk()
    app = AutoBPAApp(root)
    root.after(100, app._iniciar)
    root.mainloop()
