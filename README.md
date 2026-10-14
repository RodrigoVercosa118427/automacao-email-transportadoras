# 📧 Automação de e-mails para transportadoras

Script em Python que envia, pelo Outlook, os e-mails mensais de cobrança de fatura para cada transportadora. Ele lê uma planilha de contatos e personaliza a mensagem para cada uma.

Feito para o fechamento mensal: em vez de escrever um e-mail por vez, o script envia todos de uma vez e registra o status de cada envio numa planilha.

## O que ele faz

- Lê a lista de transportadoras em `email.xlsx`
- Monta a saudação conforme a hora (bom dia, boa tarde ou boa noite)
- Personaliza a mensagem pelo tipo de cobrança: **Mensal** ou **Quinzenal** (2ª quinzena)
- Escreve o mês em português (ex.: `Outubro/2026`) e define o prazo como o dia 1º do mês seguinte
- Coloca em cópia o e-mail da coluna `COPIA`, quando preenchida
- Gera `emails_enviados.xlsx` com a coluna **Status Envio** para cada linha

**Exemplo de assunto:** `Fatura Mensal de Outubro/2026 Transportes Exemplo Ltda`

## Requisitos

- Windows com o **Microsoft Outlook** instalado e configurado
- Python 3.10 ou mais novo

```bash
pip install -r requirements.txt
```

## Como usar

1. Copie `email_exemplo.xlsx` para `email.xlsx` e preencha com as transportadoras:

   | TRANSPORTADORA | EMAIL | COPIA | TIPO |
   |---|---|---|---|
   | Transportes Exemplo Ltda | financeiro@transportes-exemplo.com.br | comercial@transportes-exemplo.com.br | Mensal |
   | Logística Modelo S.A. | faturamento@logistica-modelo.com.br | | Quinzenal |

2. Rode em **modo teste** (padrão). Os e-mails ficam na pasta **Rascunhos** do Outlook para você conferir:

   ```bash
   python envio_emails.py
   ```

3. Se estiver tudo certo, abra `envio_emails.py`, troque `MODO_TESTE = True` por `MODO_TESTE = False` e rode de novo para enviar.

4. Confira o resultado em `emails_enviados.xlsx`.

> O `email.xlsx` e o `emails_enviados.xlsx` estão no `.gitignore`, para os contatos reais nunca irem para o GitHub.

## Tecnologias

- **pandas** e **openpyxl**: leitura e escrita das planilhas Excel
- **pywin32** (`win32com`): automação do Outlook

## Autor

**Rodrigo Verçosa**: projeto desenvolvido em aula de Python.
