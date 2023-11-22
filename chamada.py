import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime

# Função para obter a instância da planilha
def obter_planilha():
    # Se estiver usando credenciais armazenadas localmente
    creds = Credentials.from_authorized_user_file("token.json")

    # Se não houver credenciais válidas, solicitar autorização
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json",
                ["https://www.googleapis.com/auth/spreadsheets"]
            )
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    # Autorizar o acesso ao Google Sheets
    gc = gspread.authorize(creds)

    # Abrir a primeira planilha
    spreadsheet_id = "1b3T7gfqZPJwHgXQyDe1a2q__IXLGH2rpzXlzZG_-TJg"
    worksheet = gc.open_by_key(spreadsheet_id).get_worksheet(0)

    return worksheet

# Função para verificar se a matrícula já existe
def matricula_existente(worksheet, matricula):
    matriculas = worksheet.col_values(2)  # Coluna B contém as matrículas
    return matricula in matriculas

# Função para registrar o ponto
def registrar_ponto(matricula, nome):
    try:
        # Obter a instância da planilha
        worksheet = obter_planilha()

        # Verificar se a matrícula já está na planilha
        if matricula_existente(worksheet, matricula):
            print(f'Matricula já encontrada: {matricula}')
        else:
            # Se a matrícula não foi encontrada, adicioná-la
            nova_linha = [nome, matricula, "", "", ""]
            worksheet.append_row(nova_linha)
            print(f"Matrícula {matricula} adicionada à planilha para {nome}")

        # Obter a linha correspondente
        row_number = worksheet.find(matricula).row

        # Verificar se a entrada já foi registrada
        entrada = worksheet.cell(row_number, 4).value
        saida = worksheet.cell(row_number, 5).value

        if not entrada:
            # Registra a entrada
            agora = datetime.now()
            entrada = agora.strftime("%H:%M:%S")
            worksheet.update_cell(row_number, 3, agora.strftime("%d/%m/%Y"))
            worksheet.update_cell(row_number, 4, entrada)
            print(f"Entrada registrada para {nome} ({matricula}): {entrada}")
        elif not saida:
            # Registra a saída
            agora = datetime.now()
            saida = agora.strftime("%H:%M:%S")
            worksheet.update_cell(row_number, 5, saida)
            print(f"Saida registrada para {nome} ({matricula}): {saida}")
        else:
            print("Não é possível registrar novamente. Já foram registradas entrada e saída.")

    except HttpError as err:
        print(err)

if __name__ == "__main__":
    # Solicitar entrada do usuário
    matricula = input("Digite a matrícula: ")
    nome = input("Digite o nome: ")

    # Chamar a função para registrar ponto
    registrar_ponto(matricula, nome)