import os
import pandas as pd
import unicodedata
from Email import enviar_email

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# ID da planilha e o espaço a ser armazenado
PLANILHA_ID = "1xDPGHt6cKMOO2iyGf_VkDBNVbdWrTujbVF-E"
NOME_ESPACO = "Página1"

# --------------------------------------------

def acessar_planilha(creds, PLANILHA_ID, NOME_ESPACO):
  try:
    # Acessando a planilha
    service = build("sheets", "v4", credentials=creds)

    # Chamando a API
    sheet = service.spreadsheets()
    result = (sheet.values().get(spreadsheetId=PLANILHA_ID, range=NOME_ESPACO).execute())
    values = result.get("values", [])
    if not values:
        print("Dados não encontrados.")
        return None
    # Criando dataframe
    return pd.DataFrame(values[1:], columns=values[0])
  
  except HttpError as err:
      print(err)
      return None

# --------------------------------------------

def remover_acentos(texto):
    if isinstance(texto, str):
        return ''.join(
            char for char in unicodedata.normalize('NFD', texto)
            if unicodedata.category(char) != 'Mn'
        )

    return texto  

# --------------------------------------------

def processar_datas(dataframe):
  # Converte colunas de hora de string para datetime e adiciona coluna de finalização.
  dataframe['Início do Atendimento'] = pd.to_datetime(dataframe['Início do Atendimento'], format='%H:%M:%S', errors='coerce')
  dataframe['Final do Atendimento'] = pd.to_datetime(dataframe['Final do Atendimento'], format='%H:%M:%S', errors='coerce')
  dataframe['Atendimento Finalizado'] = ~dataframe['Final do Atendimento'].isna()

  return dataframe

# --------------------------------------------

def calc_estatisticas(dataframe):
    # Cálculo do tempo total de atendimento em segundos
    df_finalizados = dataframe[dataframe['Atendimento Finalizado'] == True].copy()
    df_finalizados['Duração'] = df_finalizados['Final do Atendimento'] - df_finalizados['Início do Atendimento']
    df_finalizados['Duração (minutos)'] = df_finalizados['Duração'].dt.total_seconds() / 60

    # Média agrupada por atendente
    media_atendentes = df_finalizados.groupby('Atendentes').agg(
        Media_Duracao=('Duração (minutos)', 'mean'),
    )

    # Quantidade de atendimentos agrupados por atendente
    estatisticas_totais = dataframe.groupby('Atendentes').agg(
        Quantidade_Atendimentos=('Atendentes', 'size'),
        Quantidade_Nao_Finalizados=('Atendimento Finalizado', lambda x: (~x).sum())
    )

    # Junção dos dois agrupamentos
    estatisticas = pd.merge(media_atendentes, estatisticas_totais, on='Atendentes', how='right')
    estatisticas['Media_Duracao'] = estatisticas['Media_Duracao'].fillna(0)

    # Exibição das estatísticas por atendente
    for index, row in estatisticas.iterrows():
        print(f"Atendente: {index}")
        print(f"  Média de Duração: {row['Media_Duracao']:.2f} minutos")
        print(f"  Quantidade de Atendimentos: {row['Quantidade_Atendimentos']}")
        print(f"  Quantidade de Atendimentos Não Finalizados: {row['Quantidade_Nao_Finalizados']}")
        print("-" * 50)

    return estatisticas

# --------------------------------------------

def salvar_resultados(dataframe_estatisticas, dataframe_pendentes):
    # Criar diretório para os arquivos, se não existir
    os.makedirs("relatorios", exist_ok=True)

    # Exportar atendimentos pendentes para Excel
    pendentes_path = "relatorios/atendimentos_pendentes.xlsx"
    dataframe_pendentes.to_excel(pendentes_path, index=False)
    print(f"\nAtendimentos pendentes salvos em: {pendentes_path}")

    # Exportar estatísticas por atendente para arquivos Excel
    arquivos_gerados = {}
    for atendente, dados in dataframe_estatisticas.iterrows():
        arquivo_path = os.path.join("relatorios", f"relatorio_{atendente}.xlsx")
        dados.to_frame().transpose().to_excel(arquivo_path, index=False)
        arquivos_gerados[atendente] = arquivo_path
    
    print(f"\nRelatórios de atendimentos individuais foram salvos em: relatorios/\n")

    return arquivos_gerados

# --------------------------------------------

def obter_destinatario(nome_atendente):
    # Mapear os atendentes para os e-mails dos gestores
    emails_gestores = {
        "Fabio": "Eva@email",
        "Julia": "Douglas@email", 
        "Patricia": "Ana@email", 
        "Rodrigo": "Ana@email" 
    }
    
    return emails_gestores.get(nome_atendente, "default@email")

# --------------------------------------------

def enviar_relatorios(arquivos_gerados, email_remetente, senha):
    # Laço de repetição para o envio dos arquivos para os gestores responsáveis
    for atendente, arquivo_path in arquivos_gerados.items():
        destinatario = obter_destinatario(atendente)
        assunto = f"Relatório de Atendimentos - {atendente}"
        corpo = f"Olá, segue em anexo o relatório de atendimentos para {atendente}."
        enviar_email(
            assunto=assunto,
            corpo=corpo,
            destinatario=destinatario,
            remetente=email_remetente,
            senha=senha,
            arquivo_anexo=arquivo_path
        )

# --------------------------------------------

def main():
  creds = None
  # Obtendo as credenciais para o acesso
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  # Se não há credenciais válidas, o usuário terá de autorizar
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      print("Credenciais inválidas ou inexistentes. Será necessário gerar um novo token.")
      flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
      creds = flow.run_local_server(port=0)

    # Salvando as credenciais para o próximo acesso
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  # Criação do dataframe para o Pandas 
  dataframe = acessar_planilha(creds, PLANILHA_ID, NOME_ESPACO)
  if dataframe is None:
    return

  # Modificação de alguns campos do dataframe
  dataframe['Atendentes'] = dataframe['Atendentes'].apply(remover_acentos)
  dataframe = processar_datas(dataframe)

  # Cálculo das estatísticas dos atendentes
  estatisticas = calc_estatisticas(dataframe)

  # Número total de atendimentos
  total_atendimentos = dataframe['Atendentes'].count()
  at_nao_finalizados = dataframe[dataframe['Atendimento Finalizado'] == False]['Atendentes'].count()

  # Exibição da quantidade de atendimentos
  print(f"\nNúmero total de atendimentos: {total_atendimentos}")
  print(f"\nAtendimentos não finalizados: {at_nao_finalizados}\n")

  # Formatando as colunas de data para exibir apenas a hora
  dataframe['Início do Atendimento'] = dataframe['Início do Atendimento'].dt.strftime('%H:%M:%S')
  dataframe['Final do Atendimento'] = dataframe['Final do Atendimento'].dt.strftime('%H:%M:%S')

  # Atendimentos não finalizados
  pendentes = dataframe[dataframe['Atendimento Finalizado'] == False]
  print(pendentes)

  return estatisticas, pendentes

# --------------------------------------------
  
if __name__ == "__main__":
  # Obtenção das informações
  estatisticas, pendentes = main()

  # Exportação dos dados para arquivos xlsx
  arquivos_gerados = salvar_resultados(estatisticas, pendentes)
  
  # Envio dos relatórios para os devidos emails
  enviar_relatorios(arquivos_gerados, "email_remetente", "senha")
  