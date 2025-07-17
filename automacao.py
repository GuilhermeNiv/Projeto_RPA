import os
from datetime import datetime
import pandas as pd
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()

EMAIL_REMETENTE = os.getenv('EMAIL_REMETENTE')
SENHA_REMETENTE = os.getenv('SENHA_REMETENTE')
EMAIL_DESTINATARIO = os.getenv('EMAIL_DESTINATARIO')

PLANILHAS_PATH = 'planilhas'

def validar_env():
    variaveis = {
        'EMAIL_REMETENTE': EMAIL_REMETENTE,
        'SENHA_REMETENTE': SENHA_REMETENTE,
        'EMAIL_DESTINATARIO': EMAIL_DESTINATARIO
    }
    faltando = [k for k, v in variaveis.items() if not v]
    if faltando:
        raise EnvironmentError(f"Variáveis .env faltando: {', '.join(faltando)}")

def listar_planilhas():
    arquivos = []
    for arquivo in os.listdir(PLANILHAS_PATH):
        if arquivo.endswith('.xlsx') and arquivo.lower() in ['vendas_01.xlsx', 'vendas_02.xlsx', 'vendas_03.xlsx']:
            arquivos.append(os.path.join(PLANILHAS_PATH, arquivo))
    if not arquivos:
        raise FileNotFoundError(f"Nenhuma planilha válida encontrada na pasta '{PLANILHAS_PATH}'")
    return arquivos

def carregar_planilhas(arquivos):
    dfs = []
    for caminho in arquivos:
        print(f"Lendo planilha: {caminho}")
        df = pd.read_excel(caminho)
        dfs.append(df)
    combinado = pd.concat(dfs, ignore_index=True)
    print(f"Planilhas carregadas e combinadas, total de linhas: {len(combinado)}")
    return combinado

def gerar_resumo(df):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_arquivo = f'resumo_{timestamp}.xlsx'
    df.to_excel(nome_arquivo, index=False)
    print(f"Arquivo resumo salvo: {nome_arquivo}")
    return nome_arquivo

def enviar_email_com_anexo(arquivo):
    msg = EmailMessage()
    msg['Subject'] = 'Resumo Consolidado de Planilhas'
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINATARIO
    msg.set_content(
        f"Olá,\n\nSegue em anexo o arquivo de resumo consolidado gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}.\n\nAtenciosamente,\nSistema de Automação"
    )

    with open(arquivo, 'rb') as f:
        dados = f.read()

    msg.add_attachment(
        dados,
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=os.path.basename(arquivo)
    )

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(EMAIL_REMETENTE, SENHA_REMETENTE)
            smtp.send_message(msg)
        print(f"E-mail enviado com sucesso para {EMAIL_DESTINATARIO}")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

def main():
    try:
        print("Validando variáveis de ambiente...")
        validar_env()

        print("Listando planilhas para processar...")
        arquivos = listar_planilhas()

        print("Carregando e consolidando planilhas...")
        df_resumo = carregar_planilhas(arquivos)

        print("Gerando arquivo de consolidado...")
        arquivo_resumo = gerar_resumo(df_resumo)

        print("Enviando e-mail com o consolidado...")
        enviar_email_com_anexo(arquivo_resumo)

        print("Processo concluído com sucesso!")
    except Exception as error:
        print(f"Erro no processo: {error}")

if __name__ == '__main__':
    main()