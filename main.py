import pandas as pd
import os
from datetime import datetime
import win32com.client as win32


def main():
    # Localizando os arquivos csv no diretório

    caminho = "bases"
    arquivos = os.listdir(caminho)

    tabela_estruturada = pd.DataFrame()

    # Iteração sobre cada arquivo csv e consolidando no Dataframe principal

    for nome_arquivo in arquivos:
        tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo))
        tabela_estruturada = pd.concat([tabela_estruturada, tabela_vendas])

    tabela_estruturada = tabela_estruturada.sort_values(by="Data de Venda")
    tabela_estruturada = tabela_estruturada.reset_index(drop=True)

    # Convertendo o Dataframe em xlsx (Excel)

    diretorio_output = "output"
    if not os.path.exists(diretorio_output):
        os.makedirs(diretorio_output)

    tabela_estruturada.to_excel(os.path.join(f"{diretorio_output}/Vendas.xlsx"), index=False)

    # Envio de Email

    enviar_email("Vendas.xlsx")


def enviar_email(anexo):
    outlook = win32.Dispatch("Outlook.Application")
    email = outlook.CreateItem(0)
    email.To = "" # Conta que receberá o email
    data = datetime.today().strftime("%d/%m/%Y")
    email.Subject = f"Relatório de Vendas do dia {data}"
    email.body = f"""
    Prezados,

    Segue em anexo o relatório de vendas do dia {data} atualizado.
    Caso haja alguma dúvida em relação ao relatório, estarei a disposição.
    """
    caminho = os.getcwd()
    relatorio = os.path.join(f"{caminho}/output/{anexo}")

    # Inserção do anexo
    email.Attachments.Add(relatorio)

    # Envio do Email
    email.Send()



if __name__ == "__main__":
    main()
