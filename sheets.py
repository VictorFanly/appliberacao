import gspread
import os
import sys
from google.oauth2 import service_account
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound

# ===== CONFIGURAÇÕES =====
SPREADSHEET_NAME = "REGISTRO_LIBERACOES"
ABA_NOME = "Página1"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ================= UTIL =================

def caminho_recurso(nome):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, nome)
    return nome

# ================= GOOGLE SHEETS =================

def conectar_planilha():
    cred_path = caminho_recurso("datasec")

    if not os.path.exists(cred_path):
        raise FileNotFoundError(
            f"Arquivo de credenciais não encontrado: {cred_path}"
        )

    credentials = service_account.Credentials.from_service_account_file(
        cred_path,
        scopes=SCOPES
    )

    cliente = gspread.authorize(credentials)

    try:
        planilha = cliente.open(SPREADSHEET_NAME)
    except SpreadsheetNotFound:
        raise Exception(
            f"Planilha '{SPREADSHEET_NAME}' não encontrada "
            "ou sem permissão para este e-mail de serviço."
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

    aba.append_row(
        [
            registro.get("LOTE", ""),
            registro.get("DATA_APREENSAO", ""),
            registro.get("TIPO_AGENTE", ""),
            registro.get("RECOLHA", ""),
            registro.get("DIAS", ""),
            registro.get("PLACA", ""),
            registro.get("MODELO", ""),
            registro.get("ATENDENTE", ""),
            registro.get("DATA_LIBERACAO", "")
        ],
        value_input_option="USER_ENTERED"
    )
