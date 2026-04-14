"""
AUTOMAÇÃO DE TEMPERATURA - RAÍZEN BPA
Monitora email, baixa PDF e DiaryLoadAutoTank, preenche planilha e roda macro SAP.
Ciclo: a cada hora. Reset de log: meia-noite.
"""

import os
import re
import json
import time
import logging
import shutil
import schedule
from datetime import datetime, date
from pathlib import Path

import pdfplumber
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# ── Usuário do Windows atual ──────────────────────────────────────────────────
import getpass
USUARIO_WINDOWS = getpass.getuser()

# ── Carrega configurações ──────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
with open(BASE_DIR / "config.json", encoding="utf-8") as f:
    CFG = json.load(f)

# Substitui {usuario} nos caminhos do config pelo usuário real do Windows
def _substituir_usuario(obj):
    if isinstance(obj, str):
        return obj.replace("{usuario}", USUARIO_WINDOWS)
    elif isinstance(obj, dict):
        return {k: _substituir_usuario(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_substituir_usuario(i) for i in obj]
    return obj

CFG = _substituir_usuario(CFG)
log_init = logging.getLogger(__name__)

# ── Pastas necessárias ─────────────────────────────────────────────────────────
for pasta in [
    CFG["email"]["pasta_download"],
    CFG["newsat"]["pasta_download"],
    CFG["automacao"]["pasta_log"],
]:
    Path(pasta).mkdir(parents=True, exist_ok=True)

# ── Logger ─────────────────────────────────────────────────────────────────────
def configurar_logger():
    pasta_log = Path(CFG["automacao"]["pasta_log"])
    arquivo_log = pasta_log / f"temperatura_{date.today().strftime('%Y%m%d')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(arquivo_log, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )

configurar_logger()
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# 1. OUTLOOK — baixar PDF do email pfpl@ipiranga.com.br
# ══════════════════════════════════════════════════════════════════════════════
# Padrões de assunto que indicam email com dados de temperatura
ASSUNTOS_TEMPERATURA = ["ABERTURA", "Virada", "Atualiza"]

def _buscar_pdfs_em_pasta(pasta_outlook, pasta_dest: Path) -> list[str]:
    """Varre uma pasta do Outlook e baixa PDFs de emails não lidos com assunto relevante."""
    pdfs_baixados = []
    try:
        mensagens = pasta_outlook.Items
        mensagens.Sort("[ReceivedTime]", True)
        for msg in mensagens:
            try:
                if msg.UnRead is False:
                    continue
                assunto = getattr(msg, "Subject", "")
                if not any(p.lower() in assunto.lower() for p in ASSUNTOS_TEMPERATURA):
                    continue
                for anexo in msg.Attachments:
                    nome = anexo.FileName
                    if nome.lower().endswith(".pdf"):
                        destino = pasta_dest / nome
                        anexo.SaveAsFile(str(destino))
                        log.info(f"  PDF baixado: {nome} (assunto: {assunto[:50]})")
                        pdfs_baixados.append(str(destino))
                msg.UnRead = False
                msg.Save()
            except Exception as e:
                log.warning(f"  Erro ao processar mensagem: {e}")
    except Exception as e:
        log.warning(f"  Erro ao varrer pasta: {e}")
    return pdfs_baixados


def _encontrar_pasta_outlook(namespace, nome_parcial: str):
    """Busca pasta pelo nome parcial em todas as contas."""
    def buscar_recursivo(pasta):
        try:
            if nome_parcial.lower() in pasta.Name.lower():
                return pasta
            for sub in pasta.Folders:
                resultado = buscar_recursivo(sub)
                if resultado:
                    return resultado
        except Exception:
            pass
        return None

    for conta in namespace.Folders:
        resultado = buscar_recursivo(conta)
        if resultado:
            return resultado
    return None


def baixar_pdf_outlook() -> list[str]:
    """
    Varre Caixa de Entrada e pasta VIRADA/ABERTURA procurando PDFs com dados de temperatura.
    Retorna lista de caminhos dos PDFs baixados.
    """
    log.info("Verificando emails de controle de qualidade...")
    pdfs_baixados = []
    pasta_dest = Path(CFG["email"]["pasta_download"])

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # 1. Caixa de Entrada — remetente pfpl@ipiranga.com.br
        try:
            inbox = namespace.GetDefaultFolder(6)
            mensagens = inbox.Items
            mensagens.Sort("[ReceivedTime]", True)
            remetente_cfg = CFG["email"]["remetente"].lower()
            for msg in mensagens:
                try:
                    if msg.UnRead is False:
                        continue
                    remetente = getattr(msg, "SenderEmailAddress", "").lower()
                    if remetente_cfg not in remetente:
                        continue
                    for anexo in msg.Attachments:
                        nome = anexo.FileName
                        if nome.lower().endswith(".pdf"):
                            destino = pasta_dest / nome
                            anexo.SaveAsFile(str(destino))
                            log.info(f"  PDF baixado (pfpl): {nome}")
                            pdfs_baixados.append(str(destino))
                    msg.UnRead = False
                    msg.Save()
                except Exception as e:
                    log.warning(f"  Erro mensagem pfpl: {e}")
        except Exception as e:
            log.warning(f"  Erro Caixa de Entrada: {e}")

        # 2. Pasta VIRADA/ABERTURA — qualquer remetente, assunto relevante
        pasta_virada = _encontrar_pasta_outlook(namespace, "VIRADA")
        if pasta_virada:
            log.info(f"  Varrendo pasta: {pasta_virada.Name}")
            novos = _buscar_pdfs_em_pasta(pasta_virada, pasta_dest)
            pdfs_baixados.extend(novos)
        else:
            log.warning("  Pasta VIRADA/ABERTURA não encontrada no Outlook.")

    except Exception as e:
        log.error(f"Erro ao acessar Outlook: {e}")

    return pdfs_baixados


# ══════════════════════════════════════════════════════════════════════════════
# 2. PDF — extrair temperatura e densidade por produto
# ══════════════════════════════════════════════════════════════════════════════

# Mapeamento de nomes do PDF para nomes da planilha
MAPA_PRODUTO = {
    "S10":      "S10",
    "S500":     "S500",
    "GASOLINA": "GASOLINA",
    "GAS":      "GASOLINA",
}

def extrair_dados_pdf(caminho_pdf: str) -> dict:
    """
    Extrai {produto: {temp, dens}} do PDF de Controle de Qualidade.
    Colunas esperadas: [Data, Produto, Tanque, Cor, Aspecto,
                        Densidade, Temperatura(ºC), ...]
    """
    dados = {}
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                tabelas = page.extract_tables()
                for tabela in tabelas:
                    for linha in tabela:
                        if not linha or len(linha) < 7:
                            continue
                        produto_raw = str(linha[1]).strip().upper() if linha[1] else ""
                        produto = MAPA_PRODUTO.get(produto_raw)
                        if not produto:
                            continue
                        try:
                            temp = float(str(linha[6]).replace(",", ".").strip())
                            dens = float(str(linha[5]).replace(",", ".").strip())
                            dados[produto] = {"temp": temp, "dens": dens}
                            log.info(f"  PDF → {produto}: temp={temp} dens={dens}")
                        except (ValueError, TypeError):
                            log.warning(f"  Não foi possível converter dados de {produto_raw}")
    except Exception as e:
        log.error(f"Erro ao ler PDF {caminho_pdf}: {e}")
    return dados


def processar_todos_pdfs(lista_pdfs: list[str]) -> dict:
    """Consolida dados de múltiplos PDFs. Último PDF vence por produto."""
    consolidado = {}
    for pdf in lista_pdfs:
        dados = extrair_dados_pdf(pdf)
        consolidado.update(dados)
    return consolidado


# ══════════════════════════════════════════════════════════════════════════════
# 3. NEW SAT — download do DiaryLoadAutoTank
# ══════════════════════════════════════════════════════════════════════════════

COOKIES_FILE = BASE_DIR / "newsat_cookies.json"
URL_RELATORIO = "https://bpa.newsat.cosan.rede/Report/DiaryLoadAutoTank"
_driver_global = None  # Chrome fica aberto entre ciclos
_driver_global = None  # Chrome fica aberto entre ciclos

def _criar_driver(pasta_download: Path):
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": str(pasta_download),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
    })
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--safebrowsing-disable-download-protection")
    from webdriver_manager.chrome import ChromeDriverManager
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    # Configura download via CDP
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": str(pasta_download)
    })
    return driver

def _salvar_cookies(driver):
    import json as _json
    cookies = driver.get_cookies()
    COOKIES_FILE.write_text(_json.dumps(cookies), encoding="utf-8")
    log.info(f"  Cookies salvos em {COOKIES_FILE}")

def _carregar_cookies(driver):
    import json as _json
    if not COOKIES_FILE.exists():
        return False
    cookies = _json.loads(COOKIES_FILE.read_text(encoding="utf-8"))
    for cookie in cookies:
        try:
            driver.add_cookie(cookie)
        except Exception:
            pass
    return True

def _ja_logado(driver) -> bool:
    """Verifica se a página atual é o relatório (não a tela de login)."""
    return "login" not in driver.current_url.lower() and "microsoftonline" not in driver.current_url.lower()

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
        # Reutiliza Chrome aberto ou cria novo
        if _driver_global is not None:
            try:
                _driver_global.title  # testa se ainda está vivo
                driver = _driver_global
                log.info("  Reutilizando Chrome aberto.")
            except Exception:
                _driver_global = None
        if driver is None:
            driver = _criar_driver(pasta_dest)
            _driver_global = driver
        wait = WebDriverWait(driver, 60)

        # Acessa relatório direto
        driver.get(URL_RELATORIO)
        time.sleep(3)

        # Carrega cookies salvos se existirem
        if COOKIES_FILE.exists():
            _carregar_cookies(driver)
            driver.get(URL_RELATORIO)
            time.sleep(3)

        # Verifica se caiu na tela de login Microsoft
        if not _ja_logado(driver):
            log.warning("  Sessão expirada. Aguardando login manual no Chrome...")
            log.warning("  >>> FAÇA O LOGIN NO CHROME QUE ABRIU E AGUARDE <<<")
            for _ in range(36):
                time.sleep(5)
                if _ja_logado(driver):
                    log.info("  Login detectado!")
                    _salvar_cookies(driver)
                    # Navega para o relatório após login
                    driver.get(URL_RELATORIO)
                    time.sleep(3)
                    break
            else:
                log.error("  Timeout no login manual.")
                return None

        log.info("  Logado no New SAT.")

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
            btn_exportar = wait.until(EC.element_to_be_clickable(
                (By.ID, "rptRelatorio_ctl05_ctl04_ctl00_ButtonImg")
            ))
            btn_exportar.click()
            log.info("  Clicou no botão exportar.")
            time.sleep(2)

            # Clica em Excel
            btn_excel = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[@title='Excel']")
            ))
            btn_excel.click()
            log.info("  Clicou em Excel.")
        except Exception as e:
            log.warning(f"  Botão não encontrado, tentando JS: {e}")
            try:
                driver.execute_script("$find('rptRelatorio').exportReport('EXCELOPENXML');")
                log.info("  Exportação via JS.")
            except Exception as e2:
                log.error(f"  Falha total ao exportar: {e2}")
                driver.save_screenshot(str(BASE_DIR / "newsat_erro_export.png"))
                driver.switch_to.default_content()
                return None

        log.info("  Aguardando download...")

        # Aguarda download (até 10 minutos)
        log.info("  Aguardando arquivo na pasta...")
        for i in range(300):
            time.sleep(2)
            # Verifica arquivo esperado
            if arquivo_esperado.exists():
                log.info(f"  DiaryLoadAutoTank baixado: {arquivo_esperado}")
                return str(arquivo_esperado)
            # Verifica qualquer xlsx novo (exceto a planilha principal)
            xlsx = [f for f in pasta_dest.glob("*.xlsx")
                    if not f.name.startswith(".")
                    and "Atualiza_Temperatura" not in f.name
                    and not f.name.endswith(".crdownload")]
            if xlsx:
                destino = pasta_dest / "DiaryLoadAutoTank.xlsx"
                if xlsx[0] != destino:
                    shutil.move(str(xlsx[0]), str(destino))
                    log.info(f"  Arquivo renomeado para DiaryLoadAutoTank.xlsx")
                return str(destino)
            if i % 10 == 0:
                log.info(f"  Aguardando download... ({i*2}s)")

        log.warning("  Timeout aguardando download.")
        driver.save_screenshot(str(BASE_DIR / "newsat_timeout.png"))
        return None

    except Exception as e:
        log.error(f"Erro no New SAT: {e}")
        try:
            if _driver_global:
                _driver_global.save_screenshot(str(BASE_DIR / "newsat_erro.png"))
        except Exception:
            _driver_global = None  # Chrome morreu, abre novo na próxima
        return None
    # Chrome não é fechado — reutilizado no próximo ciclo


# ══════════════════════════════════════════════════════════════════════════════
# 4. EXCEL — preencher campos amarelos e rodar macro
# ══════════════════════════════════════════════════════════════════════════════
def _aceitar_popup_sap():
    """Pressiona Enter para aceitar popup do SAP que pede permissão de acesso."""
    try:
        import win32gui
        import win32con
        # Procura janela do SAP com título de permissão
        def callback(hwnd, extra):
            titulo = win32gui.GetWindowText(hwnd)
            if any(x in titulo for x in ["SAP", "Script", "Allow", "Permitir", "Macro"]):
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        win32gui.EnumWindows(callback, None)
    except Exception as e:
        log.warning(f"  Não foi possível tratar popup SAP: {e}")


def rodar_macro_somente() -> bool:
    """Roda a macro sem alterar nenhum campo — usa os dados já existentes na planilha."""
    caminho = CFG["planilha"]["caminho"]
    macro = CFG["planilha"]["macro"]
    log.info(f"Rodando macro com dados atuais: {macro}")
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True   # Visível para que o SAP GUI consiga interagir
        excel.DisplayAlerts = False
        # Verifica se a planilha já está aberta
        try:
            wb = excel.Workbooks(os.path.basename(caminho))
            log.info("  Planilha já estava aberta, reutilizando.")
        except Exception:
            wb = excel.Workbooks.Open(caminho)
        # Aceita popup SAP antes de rodar
        _aceitar_popup_sap()
        excel.Application.Run(f"'{os.path.basename(caminho)}'!{macro}")
        wb.Save()
        log.info("  Macro executada com sucesso.")
        return True
    except Exception as e:
        log.error(f"Erro ao rodar macro: {e}")
        return False
    # Não fecha o Excel — mantém aberto para o próximo ciclo


def atualizar_planilha_e_rodar_macro(dados_pdf: dict) -> bool:
    """
    Abre a planilha via COM, preenche H7:I9 com dados do PDF
    e executa a macro Extrai_Atualiza.
    """
    if not dados_pdf:
        log.warning("Nenhum dado de PDF para preencher na planilha.")
        return False

    caminho = CFG["planilha"]["caminho"]
    macro = CFG["planilha"]["macro"]
    celulas = CFG["planilha"]["celulas"]

    log.info(f"Abrindo planilha: {caminho}")
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True   # Visível para que o SAP GUI consiga interagir
        excel.DisplayAlerts = False

        # Verifica se planilha já está aberta
        try:
            wb = excel.Workbooks(os.path.basename(caminho))
            log.info("  Planilha já estava aberta, reutilizando.")
        except Exception:
            wb = excel.Workbooks.Open(caminho)
        ws = wb.Sheets("INICIO")

        for produto, vals in dados_pdf.items():
            if produto not in celulas:
                continue
            cel_temp = celulas[produto]["temp"]
            cel_dens = celulas[produto]["dens"]
            ws.Range(cel_temp).Value = vals["temp"]
            ws.Range(cel_dens).Value = vals["dens"]
            log.info(f"  Preenchido {produto}: Temp={vals['temp']} Dens={vals['dens']}")

        wb.Save()

        # Aceita popup SAP e roda a macro
        _aceitar_popup_sap()
        log.info(f"  Executando macro: {macro}")
        excel.Application.Run(f"'{os.path.basename(caminho)}'!{macro}")
        wb.Save()
        log.info("  Macro executada com sucesso.")
        return True

    except Exception as e:
        log.error(f"Erro ao manipular planilha: {e}")
        return False
    # Não fecha o Excel — mantém aberto para o próximo ciclo


# ══════════════════════════════════════════════════════════════════════════════
# 5. RESET DIÁRIO — limpa log e PDFs às 00:00
# ══════════════════════════════════════════════════════════════════════════════
def reset_diario():
    log.info("═══ RESET DIÁRIO ═══ Iniciando novo ciclo de 24h")

    # Arquiva PDFs do dia anterior
    pasta_pdfs = Path(CFG["email"]["pasta_download"])
    pasta_arquivo = pasta_pdfs / "Arquivo" / date.today().strftime("%Y%m%d")
    pasta_arquivo.mkdir(parents=True, exist_ok=True)
    for pdf in pasta_pdfs.glob("*.pdf"):
        shutil.move(str(pdf), str(pasta_arquivo / pdf.name))

    # Reconfigura logger com data nova
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    configurar_logger()
    log.info("Log reiniciado para o novo dia.")


# ══════════════════════════════════════════════════════════════════════════════
# 6. CICLO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def ciclo_hora():
    log.info("━━━ INÍCIO DO CICLO HORÁRIO ━━━")
    inicio = datetime.now()

    # 1. Baixa DiaryLoadAutoTank do New SAT
    baixar_diary_newsat()

    # 2. Roda macro com dados atuais da planilha
    log.info("Rodando macro com dados atuais da planilha.")
    sucesso = rodar_macro_somente()
    status = "OK" if sucesso else "ERRO"

    duracao = (datetime.now() - inicio).seconds
    log.info(f"━━━ FIM DO CICLO | Status: {status} | Duração: {duracao}s ━━━\n")


# ══════════════════════════════════════════════════════════════════════════════
# 7. MANTER SESSÕES VIVAS (a cada 5 minutos)
# ══════════════════════════════════════════════════════════════════════════════
def manter_sessoes_vivas():
    """Acessa New SAT e SAP rapidamente só para evitar timeout por inatividade."""
    global _driver_global
    log.info("── Keep-alive: mantendo sessões ativas ──")

    # New SAT — abre o relatório para manter sessão ativa
    try:
        if _driver_global is not None:
            _driver_global.get(URL_RELATORIO)
            time.sleep(2)
            # Se caiu na tela de login, marca para refazer login no próximo ciclo
            if not _ja_logado(_driver_global):
                log.warning("  New SAT: sessão expirou. Aguardando login manual no Chrome...")
                log.warning("  >>> FAÇA O LOGIN NO CHROME QUE ABRIU E AGUARDE <<<")
                for _ in range(36):
                    time.sleep(5)
                    if _ja_logado(_driver_global):
                        log.info("  Login detectado! Sessão renovada.")
                        _salvar_cookies(_driver_global)
                        break
                else:
                    log.error("  Timeout no login. Chrome será recriado no próximo ciclo.")
                    _driver_global = None
            else:
                log.info("  New SAT: sessão mantida.")
        else:
            log.info("  New SAT: Chrome não aberto, pulando keep-alive.")
    except Exception as e:
        log.warning(f"  New SAT keep-alive falhou: {e}")
        _driver_global = None

    # SAP — abre ZV04 e volta ao menu para evitar timeout
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        conn = app.Children(0)
        session = conn.Children(0)
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZV04"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        log.info("  SAP: ZV04 executado, sessão mantida.")
    except Exception as e:
        log.warning(f"  SAP keep-alive falhou: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# 8. AGENDAMENTO
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    log.info("╔══════════════════════════════════════════╗")
    log.info("║  AUTOMAÇÃO TEMPERATURA BPA - INICIANDO   ║")
    log.info("╚══════════════════════════════════════════╝")

    intervalo = CFG["automacao"]["intervalo_minutos"]

    # Ciclo completo a cada 30 minutos
    schedule.every(intervalo).minutes.do(ciclo_hora)
    # Keep-alive a cada 5 minutos
    schedule.every(5).minutes.do(manter_sessoes_vivas)
    # Reset diário à meia-noite
    schedule.every().day.at("00:00").do(reset_diario)

    # Roda imediatamente na primeira vez
    ciclo_hora()

    # Loop principal
    while True:
        schedule.run_pending()
        time.sleep(30)
