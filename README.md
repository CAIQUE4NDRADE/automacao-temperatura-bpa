# 🌡️ AutoBPA Monitor — Automação de Temperatura · Raízen BPA

Automação desenvolvida para eliminar o processo manual de atualização de temperatura no SAP, integrando New SAT, Excel e SAP GUI em um ciclo totalmente automático, com interface gráfica e alertas via Telegram.

---

## 📌 Sobre o Projeto

No terminal de combustíveis da Raízen — Base de Paulínia (BPA), a atualização de temperatura no SAP era realizada manualmente pelos operadores de gate a cada ciclo, consumindo em média 10 minutos por atualização e exigindo atenção constante durante o turno.

Este projeto automatiza 100% desse processo, permitindo que os operadores foquem em atividades de maior valor operacional.

---

## ⚙️ O que o sistema faz

| Frequência | Ação |
|---|---|
| A cada 30 minutos | Baixa o `DiaryLoadAutoTank.xlsx` do New SAT |
| A cada 30 minutos | Fecha e reabre a planilha limpa, prepara o SAP na ZO3C2 e roda a macro `Extrai_Atualiza` |
| A cada 5 minutos | Keep-alive: executa ZV04 no SAP e acessa o New SAT para evitar timeout por inatividade |
| Ao voltar da suspensão | Detecta automaticamente e roda ciclo imediato se necessário |
| À meia-noite | Reseta o log diário automaticamente |

---

## 🖥️ Interface Gráfica

O sistema possui interface visual desenvolvida em Tkinter:

- **Botões Iniciar / Parar** — controle total da automação
- **Status em tempo real** — 🟢 Rodando / 🔴 Erro / ⚫ Parado
- **Contagem regressiva** para o próximo ciclo
- **Contadores** de ciclos OK e erros
- **Log colorido** em tempo real na tela (verde = normal, laranja = aviso, vermelho = erro)

---

## 📲 Alertas via Telegram

O sistema envia notificações automáticas para o celular do operador:

| Evento | Mensagem |
|---|---|
| Programa iniciado | ✅ AutoBPA Monitor iniciado |
| Programa parado | ⏹ AutoBPA Monitor pausado |
| Erro no ciclo | ⚠️ Erro no ciclo — verificar SAP e log |
| 70 min sem ciclo | ⚠️ Sistema parado! Último ciclo há X minutos |

---

## 🛡️ Proteções e Resiliências

- **Chrome com perfil fixo** — login salvo, sem necessidade de autenticar a cada ciclo
- **Kill de Chrome órfão** — mata processos travados antes de recriar sessão
- **Reconexão automática** do Chrome ao detectar queda
- **Flag `_macro_rodando`** — impede que o keep-alive SAP interrompa a execução da macro
- **Verificação de tela SAP** — keep-alive pula se o operador estiver em qualquer transação
- **Detecção de suspensão** — ao voltar, decide se roda ciclo imediato ou aguarda
- **Retry imediato** se o download do Diary falhou antes da suspensão
- **Planilha fixada por nome** — evita conflito com outras planilhas abertas
- **Loop principal blindado** — nunca fecha, qualquer erro é logado e o programa continua
- **Instância única** — impede abertura dupla acidental
- **Watchdog** — alerta Telegram se ficar mais de 70 minutos sem ciclo completo
- **Impedimento de suspensão** via `powercfg` ao iniciar

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.11+** — orquestração do fluxo completo
- **Tkinter** — interface gráfica nativa
- **Selenium + ChromeDriver** — automação do New SAT (login SSO Microsoft, navegação, download)
- **pywin32 (win32com)** — integração com Excel e SAP GUI Scripting
- **schedule** — agendamento dos ciclos
- **requests** — envio de alertas via Telegram Bot API
- **webdriver-manager** — gerenciamento automático do ChromeDriver

---

## 🗂️ Estrutura do Projeto

```
Atualização de Temperatura/
├── monitor_temperatura.py          # Script principal (automação + interface)
├── config.json                     # Configurações (caminhos, macro, intervalos)
├── iniciar.bat                     # Inicializa a automação
├── Atualiza Temperatura BPA rev2.xlsm  # Planilha com macro SAP
├── chrome_profile/                 # Perfil fixo do Chrome (login salvo)
└── Logs/                           # Logs diários gerados automaticamente
```

---

## 🚀 Como Instalar

### Pré-requisitos

- Python 3.11+
- Google Chrome instalado
- Microsoft Excel com a planilha `.xlsm` configurada
- SAP GUI com scripting habilitado e logado

### Instalação das dependências

```bash
pip install selenium webdriver-manager schedule pywin32 pdfplumber requests pyautogui pyperclip
```

### Configuração

O arquivo `config.json` usa `{usuario}` como variável dinâmica — o sistema detecta automaticamente o usuário Windows logado, sem necessidade de edição manual.

```json
{
    "planilha": {
        "caminho": "C:\\Users\\{usuario}\\Downloads\\Atualização de Temperatura\\Atualiza Temperatura BPA rev2.xlsm",
        "macro": "Extrai_Atualiza"
    },
    "newsat": {
        "pasta_download": "C:\\Users\\{usuario}\\Downloads\\Atualização de Temperatura"
    },
    "automacao": {
        "intervalo_minutos": 30,
        "pasta_log": "C:\\Users\\{usuario}\\Downloads\\Atualização de Temperatura\\Logs"
    }
}
```

### Execução

```bash
python monitor_temperatura.py
```

Ou clique duas vezes em `iniciar.bat`.

---

## 📊 Ganhos Operacionais

| Indicador | Antes | Depois |
|---|---|---|
| Tempo por atualização | ~10 minutos | 0 minutos (automático) |
| Intervenção manual no ciclo | Obrigatória | Apenas em viradas/aberturas |
| Risco de esquecimento | Alto | Eliminado |
| Disponibilidade do operador | Parcialmente comprometida | Total |
| Monitoramento remoto | Inexistente | Alertas no celular via Telegram |

---

## ⚠️ Intervenções Manuais Necessárias

O sistema não substitui a intervenção humana nos seguintes momentos — que requerem julgamento operacional:

- **Virada de tanque** — o operador deve alimentar os dados na planilha
- **Atualização de densidade** — inserção manual dos laudos de qualidade
- **Início do dia (abertura)** — preenchimento dos dados iniciais do turno

---

## 📋 Exemplo de Log

```
14:00:01 [INFO] ━━━ INÍCIO DO CICLO HORÁRIO ━━━
14:00:05 [INFO]   Acessando New SAT para baixar DiaryLoadAutoTank...
14:02:30 [INFO]   DiaryLoadAutoTank baixado com sucesso
14:02:31 [INFO]   Fechando planilha para reabrir limpa...
14:02:33 [INFO]   SAP preparado na ZO3C2 para execução da macro.
14:02:55 [INFO]   Macro executada com sucesso.
14:02:55 [INFO] ━━━ FIM DO CICLO | Status: OK | Duração: 174s ━━━

14:05:00 [INFO] ── Keep-alive: mantendo sessões ativas ──
14:05:02 [INFO]   New SAT: sessão mantida.
14:05:04 [INFO]   SAP: ZV04 executado, sessão mantida.
```

---

## 👤 Autor

**Caique Bezerra de Andrade**  
Operações de Terminal — Raízen BPA · Paulínia/SP  
Transição de carreira para Tecnologia | Front End & Análise e Desenvolvimento de Sistemas

[![GitHub](https://img.shields.io/badge/GitHub-CAIQUE4NDRADE-black?logo=github)](https://github.com/CAIQUE4NDRADE)

---

## 📄 Licença

Projeto desenvolvido internamente para uso operacional na Raízen BPA.  
Código disponibilizado para fins de portfólio e aprendizado.
