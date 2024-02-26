import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
import json
import tkinter as tk  

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

# Função para obter o nome correspondente à matrícula do arquivo JSON
def obter_nome_por_matricula(matricula):
    with open("nomes_matriculas.json", "r") as file:
        dados = json.load(file)
        return dados.get(matricula, None)

# Função para registrar o ponto
def registrar_ponto(matricula, nome, rotulo_retorno):
    try:
        # Obter a instância da planilha
        worksheet = obter_planilha()

        # Se a matrícula não foi encontrada, adicioná-la
        if not matricula_existente(worksheet, matricula):
            nova_linha = [nome, matricula, "", "", ""]
            worksheet.append_row(nova_linha)
            rotulo_retorno["text"] = f"Matrícula {matricula} adicionada à planilha para {nome}"

        # Obter a última linha preenchida
        ultima_linha = worksheet.row_count

        # Verificar se já existe uma entrada e uma saída registradas
        entrada = worksheet.cell(ultima_linha, 4).value if ultima_linha > 0 else None
        saida = worksheet.cell(ultima_linha, 5).value if ultima_linha > 0 else None

        if not entrada:
            # Registra a entrada
            agora = datetime.now().date()
            entrada = datetime.now().strftime("%H:%M:%S")
            worksheet.update_cell(ultima_linha, 3, agora.strftime("%d/%m/%Y"))
            worksheet.update_cell(ultima_linha, 4, entrada)
            rotulo_retorno["text"] = f"Entrada registrada para {nome} ({matricula}): {entrada}"
        elif not saida:
            # Registra a saída
            agora = datetime.now().strftime("%H:%M:%S")
            worksheet.update_cell(ultima_linha, 5, agora)
            rotulo_retorno["text"] = f"Saida registrada para {nome} ({matricula}): {agora}"
        else:
            # Adicionar uma nova linha abaixo
            nova_linha = [nome, matricula, "", "", ""]
            worksheet.append_row(nova_linha)
            nova_row_number = ultima_linha + 1
            agora = datetime.now().strftime("%H:%M:%S")
            worksheet.update_cell(nova_row_number, 3, datetime.now().strftime("%d/%m/%Y"))
            worksheet.update_cell(nova_row_number, 4, agora)
            rotulo_retorno["text"] = f"Nova linha adicionada para {nome} ({matricula}) com entrada registrada: {agora}"

    except HttpError as err:
        rotulo_retorno["text"] = str(err)

    # Limpar o campo de entrada após o processamento
    entrada_matricula.delete(0, tk.END)

# Função para lidar com a entrada da matrícula
def obter_matricula_e_registrar_ponto(entrada_matricula, rotulo_retorno):
    matricula = entrada_matricula.get()

    # Obter o nome correspondente à matrícula
    nome = obter_nome_por_matricula(matricula)

    if nome:
        # Chamar a função para registrar ponto
        registrar_ponto(matricula, nome, rotulo_retorno)
    else:
        rotulo_retorno["text"] = f"Matrícula {matricula} não encontrada nos dados."

# Configurar a interface gráfica com Tkinter
def centralizar_janela(janela):
    janela.update_idletasks()
    largura = janela.winfo_width()
    altura = janela.winfo_height()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry('{}x{}+{}+{}'.format(largura, altura, x, y))

root = tk.Tk()
root.title("Registro de Ponto")

# Aumentar o tamanho da janela
root.geometry("600x200")

# Centralizar a janela na tela
centralizar_janela(root)

# Adicionar uma entrada para a matrícula
tk.Label(root, text="Digite a matrícula:").pack(pady=10)
entrada_matricula = tk.Entry(root)
entrada_matricula.pack(pady=10)

# Adicionar um botão para enviar a matrícula
botao_enviar = tk.Button(root, text="Registrar Ponto", command=lambda: obter_matricula_e_registrar_ponto(entrada_matricula, rotulo_retorno))
botao_enviar.pack(pady=10)

# Adicionar um rótulo para exibir mensagens de retorno
rotulo_retorno = tk.Label(root, text="")
rotulo_retorno.pack(pady=10)

# Iniciar o loop principal da interface gráfica
root.mainloop()
