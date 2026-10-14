# Rodrigo Firmino Verçosa Dos Santos
# Envio automático de e-mails de cobrança de fatura para transportadoras
#
# Lê a planilha email.xlsx, monta um e-mail personalizado para cada
# transportadora (mensal ou quinzenal) e envia pelo Outlook instalado no Windows.
# No fim, gera emails_enviados.xlsx com o status de cada envio.

import os
from datetime import datetime

import pandas as pd
import win32com.client as win32

# True  = só cria rascunhos no Outlook (bom para testar)
# False = envia de verdade
MODO_TESTE = True

PASTA = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_ENTRADA = os.path.join(PASTA, "email.xlsx")
ARQUIVO_SAIDA = os.path.join(PASTA, "emails_enviados.xlsx")

MESES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}


def saudacao(hora):
    """Bom dia, boa tarde ou boa noite, de acordo com a hora."""
    if hora <= 11:
        return "bom dia!"
    if hora <= 17:
        return "boa tarde!"
    return "boa noite!"


def primeiro_dia_proximo_mes(data):
    """Prazo para a transportadora enviar os valores."""
    if data.month == 12:
        return datetime(data.year + 1, 1, 1)
    return datetime(data.year, data.month + 1, 1)


def montar_email(row, agora):
    """Devolve (assunto, corpo_html) para uma transportadora."""
    mes_ano = f"{MESES[agora.month]}/{agora.year}"
    data_limite = primeiro_dia_proximo_mes(agora)

    if row.TIPO == "Quinzenal":
        periodo = "2ª Quinzena"
        mensagem_principal = (
            "Por gentileza, encaminhar os custos referente aos serviços "
            f"prestados na 2ª quinzena de {mes_ano}."
        )
    else:
        periodo = "Mensal"
        mensagem_principal = (
            "Por gentileza, encaminhar os custos referente aos serviços "
            f"prestados em {mes_ano} para validação."
        )

    assunto = f"Fatura {periodo} de {mes_ano} {row.TRANSPORTADORA}"
    corpo = f"""
    <p>Olá {row.TRANSPORTADORA}, {saudacao(agora.hour)}</p>
    <p>{mensagem_principal}</p>
    <p>Devido ao processo de encerramento mensal precisamos que os valores
    sejam enviados até {data_limite.strftime("%d/%m/%Y")}, tudo bem?</p>
    """
    return assunto, corpo


def main():
    planilha = pd.read_excel(
        ARQUIVO_ENTRADA,
        dtype={"TRANSPORTADORA": "str", "EMAIL": "str", "COPIA": "str", "TIPO": "str"},
    )

    outlook = win32.Dispatch("outlook.application")
    agora = datetime.now()

    for idx, row in planilha.iterrows():
        assunto, corpo = montar_email(row, agora)

        email = outlook.CreateItem(0)
        email.To = row.EMAIL
        if isinstance(row.COPIA, str) and row.COPIA.strip():
            email.CC = row.COPIA
        email.Subject = assunto
        email.HTMLBody = corpo

        try:
            if MODO_TESTE:
                email.Save()  # fica em Rascunhos
                planilha.at[idx, "Status Envio"] = "Rascunho criado (modo teste)"
            else:
                email.Send()
                planilha.at[idx, "Status Envio"] = "E-mail enviado com sucesso!"
        except Exception as erro:
            planilha.at[idx, "Status Envio"] = f"E-mail não enviado! ({erro})"

    planilha.to_excel(ARQUIVO_SAIDA, index=False)
    print(f"Concluído. Status salvo em {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    main()
