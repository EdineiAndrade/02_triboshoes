import os
import gspread
from google.oauth2.service_account import Credentials

def get_credentials_file():  

    return os.path.join("C:\\sh", "credentials.json")

def authenticate_google_sheets():  
    #Autentica no Google Sheets e retorna o cliente gspread.    
    credentials_file = get_credentials_file()
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    # Verifica se o arquivo de credenciais existe antes de autenticar
    if not os.path.exists(credentials_file):
        raise FileNotFoundError(f"❌ Arquivo de credenciais não encontrado: {credentials_file}")

    credentials = Credentials.from_service_account_file(credentials_file, scopes=scopes)
    return gspread.authorize(credentials)

def save_to_google_sheets(data):    
    sheet_id = "1TFmVRbu-pt2qQCfq7qisU2gWPZm3pJ1q48NeL4DMVFs"  # ID da planilha
    client = authenticate_google_sheets()
    sheet = client.open_by_key(sheet_id).get_worksheet(2)  # Abre a terceira aba
    
    # Apagar os dados atuais
    sheet.clear()
    # Salvar DF atualizado
    # Converte o DataFrame para uma lista de listas
    data_to_upload = [data.columns.tolist()] + data.values.tolist()

    # Insere os dados no Google Sheets
    sheet.append_rows(data_to_upload)
    print("✅ Dados salvos no Google Sheets.")


