import gspread
import os
import sys
import json
from google.oauth2.service_account import Credentials
from gspread.exceptions import SpreadsheetNotFound,WorksheetNotFound

# ===== CONFIGURAÇÕES =====
SPREADSHEET_NAME = "REGISTRO_LIBERACOES"
ABA_NOME = "Página1"  # ou o nome real da aba

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def conectar_planilha():
    creds = carregar_credenciais()

    cliente = gspread.authorize(creds)
    try:
        planilha = cliente.open(SPREADSHEET_NAME)
    except SpreadsheetNotFound:
        raise Exception(
            f"Planilha '{SPREADSHEET_NAME}' não encontrada "
            "ou sem permissão de acesso."
        )

    try:
        aba = planilha.worksheet(ABA_NOME)
    except WorksheetNotFound:
        raise Exception(
            f"A aba '{ABA_NOME}' não existe na planilha."
        )

    return aba

def registrar_liberacao(registro):
    aba = conectar_planilha()

    aba.append_row([
        registro["LOTE"],
        registro["DATA_APREENSAO"],
        registro["TIPO_AGENTE"],
        registro["RECOLHA"],
        registro["DIAS"],
        registro["PLACA"],
        registro["MODELO"],
        registro["ATENDENTE"],
        registro["DATA_LIBERACAO"]
    ])

def caminho_recurso(nome):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, nome)
    return nome

def carregar_credenciais():
    with open(caminho_recurso("datasec"), "r", encoding="utf-8") as f:
        info = json.load(f)

    return Credentials.from_service_account_info(
        info,
        scopes=SCOPES
    )
