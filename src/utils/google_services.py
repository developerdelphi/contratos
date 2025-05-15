# utils/google_services.py
import os
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from flask import flash, current_app # Para usar logger e flash
from config import Config


CREDENTIALS_FILE = Config.CREDENTIALS_FILE


# Defina SCOPES_GSPREAD e SCOPES_DRIVE aqui também se forem usados apenas por estas funções
SCOPES_GSPREAD = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
SCOPES_DRIVE = ['https://www.googleapis.com/auth/drive.readonly']


def get_google_sheets_client(logger, credentials_file_path=CREDENTIALS_FILE):
    logger.debug(f"Tentando autenticar Google Sheets com credenciais: {credentials_file_path}")
    try:
        if not os.path.exists(credentials_file_path):
            logger.error(f"Arquivo de credenciais '{credentials_file_path}' não encontrado.")
            # Não chame flash aqui, deixe a rota tratar. Retorne None ou levante uma exceção.
            return None # Exemplo
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file_path, SCOPES_GSPREAD)
        client = gspread.authorize(creds)
        logger.info("Cliente Google Sheets autenticado com sucesso.")
        return client
    except Exception as e:
        logger.error(f"Erro ao autenticar com Google Sheets: {e}", exc_info=True)
        return None
    

def get_sheet_data(sheet_url_or_id, logger):
    """Busca dados de uma Planilha Google e retorna como DataFrame."""
    logger.info(f"Tentando obter cliente Google Sheets para: {sheet_url_or_id}")
    client = get_google_sheets_client(logger=logger, credentials_file_path=CREDENTIALS_FILE)
    if not client:
        logger.error("Cliente autenticação não fornecido para get_sheet_data.")
        return None
    
    try:
        logger.info(f"Tentando abrir planilha com identificador: '{sheet_url_or_id}'")
        # gspread.open_by_url() ou gspread.open_by_key() são mais explícitos
        # Se o usuário sempre fornecer o ID, open_by_key é melhor.
        # Se for URL, open_by_url. gspread.open() tenta adivinhar.
        
        # Vamos assumir que o usuário está fornecendo o ID diretamente
        # Se for uma URL completa, você pode precisar extrair o ID dela primeiro.
        # Ex: "https://docs.google.com/spreadsheets/d/ESTE_EH_O_ID/edit#gid=0" -> ID: "ESTE_EH_O_ID"
        
        # Simplificação: Se for uma URL, tente abrir por URL, senão, por chave (ID)
        actual_id_to_open = sheet_url_or_id
        
        if "docs.google.com/spreadsheets/d/" in sheet_url_or_id:
            actual_id_to_open = sheet_url_or_id.split('/d/')[1].split('/')[0]
            logger.info(f"ID extraído da URL: {actual_id_to_open}")
            sheet = client.open_by_key(actual_id_to_open) # Usar open_by_key com o ID extraído
        else: # Assume que é um ID direto
            logger.info(f"Assumindo que '{sheet_url_or_id}' é um ID de planilha.")
            sheet = client.open_by_key(sheet_url_or_id)

        logger.info(f"Planilha '{sheet.title}' aberta com sucesso.")
        
        worksheet = sheet.sheet1 # Assume a primeira aba
        logger.info(f"Acessando primeira aba (worksheet): '{worksheet.title}'.")
        
        data = worksheet.get_all_records() # Espera cabeçalhos na primeira linha
        logger.info(f"Dados lidos da planilha: {len(data)} registros.")
        
        if not data:
            logger.warning(f"Nenhum dado encontrado na planilha '{sheet.title}' (worksheet: '{worksheet.title}'). Verifique se há dados e cabeçalhos.")

        return pd.DataFrame(data)

    except gspread.exceptions.SpreadsheetNotFound:
        logger.error(f"SpreadsheetNotFound para o identificador: '{sheet_url_or_id}' (ID tentado: '{actual_id_to_open if 'actual_id_to_open' in locals() else sheet_url_or_id}').")
        return None
    except gspread.exceptions.APIError as api_e:
        logger.error(f"gspread.exceptions.APIError ao acessar planilha '{sheet_url_or_id}': {api_e}", exc_info=True)
        return None
    except Exception as e:
        logger.error(f"Erro inesperado ao ler planilha '{sheet_url_or_id}': {e}", exc_info=True)
        return None
