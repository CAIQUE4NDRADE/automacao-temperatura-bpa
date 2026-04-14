# AUTOMAÇÃO DE TEMPERATURA - RAÍZEN BPA
## Como funciona

A cada 1 hora o sistema:
1. Verifica email de pfpl@ipiranga.com.br e baixa PDFs novos
2. Acessa New SAT e baixa o DiaryLoadAutoTank.xlsx
3. Lê os dados do PDF (S10, S500, Gasolina) e preenche a planilha
4. Executa a macro Extrai_Atualiza automaticamente
5. Registra tudo no log do dia

À meia-noite: arquiva os PDFs do dia e começa log novo.

---

## INSTALAÇÃO (apenas uma vez)

1. Copie esta pasta para: `C:\Users\cs338104\Downloads\Atualização de Temperatura\`
2. Clique com botão direito em `instalar.bat` → **Executar como administrador**
3. Abra `config.json` e preencha:
   - `newsat.usuario` → seu login do New SAT
   - `newsat.senha` → sua senha do New SAT

---

## ARQUIVOS

| Arquivo | Função |
|---|---|
| `monitor_temperatura.py` | Script principal |
| `config.json` | Configurações (edite usuário/senha) |
| `instalar.bat` | Instala dependências e cria tarefa no Windows |
| `iniciar.bat` | Inicia manualmente se necessário |

---

## LOGS

Ficam em: `Logs\temperatura_AAAAMMDD.log`

Exemplo de log saudável:
```
08:00:01 [INFO] ━━━ INÍCIO DO CICLO HORÁRIO ━━━
08:00:02 [INFO] Verificando emails de controle de qualidade...
08:00:05 [INFO] PDF baixado: PFPL_1001_CONTROLE_DE_QUALIDADE.pdf
08:00:06 [INFO]   PDF → S500: temp=31.0 dens=0.835
08:00:08 [INFO] Acessando New SAT para baixar DiaryLoadAutoTank...
08:00:25 [INFO]   DiaryLoadAutoTank baixado com sucesso
08:00:26 [INFO] Abrindo planilha...
08:00:30 [INFO]   Preenchido S500: Temp=31.0 Dens=0.835
08:00:35 [INFO]   Macro executada com sucesso.
08:00:35 [INFO] ━━━ FIM DO CICLO | Status: OK | Duração: 34s ━━━
```

---

## ATENÇÃO

- O Chrome precisa estar instalado no servidor
- O Outlook precisa estar aberto e logado
- O Excel precisa estar instalado (licença ativa)
- **Não feche o Outlook nem o Excel** enquanto a automação estiver rodando

---

## AJUSTE DE SELETORES DO NEW SAT

Se a navegação no New SAT falhar, será necessário ajustar os seletores CSS/XPath
no arquivo `monitor_temperatura.py` na função `baixar_diary_newsat()`.
Isso depende do HTML exato do site. Os seletores atuais são genéricos e podem
precisar de ajuste fino na primeira execução — verifique o log para identificar
onde parou.
