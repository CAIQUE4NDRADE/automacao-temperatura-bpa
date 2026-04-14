# 🌡️ Automação de Atualização de Temperatura — Raízen BPA

> Automação desenvolvida para eliminar o processo manual de atualização de temperatura no SAP, integrando New SAT, Outlook e Excel em um ciclo totalmente automático.

---

## 📌 Sobre o Projeto

No terminal de combustíveis da Raízen — Base de Paulínia (BPA), a atualização de temperatura no SAP era realizada manualmente pelos operadores de gate a cada ciclo, consumindo em média **10 minutos por atualização** e exigindo atenção constante durante o turno.

Este projeto automatiza 100% desse processo, permitindo que os operadores foquem em atividades de maior valor operacional.

---

## ⚙️ O que o sistema faz

| Frequência | Ação |
|---|---|
| A cada **30 minutos** | Baixa o `DiaryLoadAutoTank.xlsx` do New SAT |
| A cada **30 minutos** | Preenche a planilha e roda a macro de atualização no SAP |
| A cada **5 minutos** | Keep-alive: executa `ZV04` no SAP e acessa o New SAT para evitar timeout por inatividade |
| À **meia-noite** | Reseta o log diário automaticamente |

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.11+** — orquestração do fluxo completo
- **Selenium + ChromeDriver** — automação do New SAT (login SSO Microsoft, navegação, download)
- **pywin32 (win32com)** — integração com Excel e SAP GUI Scripting
- **pdfplumber** — extração de dados de PDFs de Controle de Qualidade
- **schedule** — agendamento dos ciclos
- **openpyxl** — manipulação da planilha Excel

---

## 🗂️ Estrutura do Projeto

```
automacao_temperatura/
├── monitor_temperatura.py   # Script principal
├── config.json              # Configurações (caminhos, macro, intervalos)
├── instalar.bat             # Instalador de dependências
├── iniciar.bat              # Inicializa a automação
└── Logs/                    # Logs diários gerados automaticamente
```

---

## 🚀 Como Instalar

### Pré-requisitos
- Python 3.11+
- Google Chrome instalado
- Microsoft Excel com a planilha `.xlsm` configurada
- SAP GUI com scripting habilitado
- Microsoft Outlook aberto e logado

### Instalação

```bash
# 1. Clone o repositório
git clone https://github.com/CAIQUE4NDRADE/automacao-temperatura-bpa.git

# 2. Entre na pasta
cd automacao-temperatura-bpa

# 3. Instale as dependências
python -m pip install pdfplumber selenium schedule pywin32 openpyxl webdriver-manager
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
        "url": "https://bpa.newsat.cosan.rede/Home"
    },
    "automacao": {
        "intervalo_minutos": 30
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

---

## ⚠️ Intervenções Manuais Necessárias

O sistema **não substitui** a intervenção humana nos seguintes momentos — que requerem julgamento operacional:

- **Virada de tanque** — o operador deve alimentar os dados na planilha
- **Atualização de densidade** — inserção manual dos laudos de qualidade
- **Início do dia (abertura)** — preenchimento dos dados iniciais do turno

---

## 📋 Exemplo de Log

```
14:00:01 [INFO] ━━━ INÍCIO DO CICLO HORÁRIO ━━━
14:00:05 [INFO]   Acessando New SAT para baixar DiaryLoadAutoTank...
14:02:30 [INFO]   DiaryLoadAutoTank baixado com sucesso
14:02:31 [INFO]   Rodando macro com dados atuais: Extrai_Atualiza
14:02:35 [INFO]   Planilha já estava aberta, reutilizando.
14:02:55 [INFO]   Macro executada com sucesso.
14:02:55 [INFO] ━━━ FIM DO CICLO | Status: OK | Duração: 174s ━━━

14:05:00 [INFO] ── Keep-alive: mantendo sessões ativas ──
14:05:02 [INFO]   New SAT: sessão mantida.
14:05:04 [INFO]   SAP: ZV04 executado, sessão mantida.
```

---

## 👤 Autor

**Caique Bezerra de Andrade**
Operações de Terminal — Raízen BPA Paulínia/SP
Transição de carreira para Tecnologia | Front End & Análise e Desenvolvimento de Sistemas

[![GitHub](https://img.shields.io/badge/GitHub-CAIQUE4NDRADE-181717?style=flat&logo=github)](https://github.com/CAIQUE4NDRADE)

---

## 📄 Licença

Projeto desenvolvido internamente para uso operacional na Raízen BPA.
Código disponibilizado para fins de portfólio e aprendizado.
