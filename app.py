import PySimpleGUI as sg
import imaplib
import email
import os
import re
from pdfminer.high_level import extract_text
from openpyxl import Workbook

# Define o tema da interface
sg.theme('reddit')

# Função para criar a janela principal


def criar_janela_principal():
    layout = [
        [sg.Text('Email'), sg.Input(key='email')],
        [sg.Text('Senha de Aplicativo'), sg.Input(
            key='senha', password_char='*')],
        [sg.Text('Servidor IMAP'), sg.Input(
            default_text='imap.gmail.com', key='imap_server')],
        [sg.FolderBrowse('Escolher Pasta para Salvar Anexos',
                         target='save_folder'), sg.Input(key='save_folder')],
        [sg.FileSaveAs('Escolher Planilha', target='planilha', file_types=(
            ("Excel Files", "*.xlsx"),)), sg.Input(key='planilha')],
        [sg.Button('Iniciar'), sg.Button('Cancelar')]
    ]
    return sg.Window('Configuração do Robô', layout=layout)

# Função principal


def main():
    janela = criar_janela_principal()

    while True:
        event, values = janela.read()
        if event == sg.WIN_CLOSED or event == 'Cancelar':
            break
        elif event == 'Iniciar':
            email_user = values['email']
            email_pass = values['senha']
            imap_server = values['imap_server']
            save_folder = values['save_folder']
            planilha_path = values['planilha']

            if not email_user or not email_pass or not imap_server or not save_folder or not planilha_path:
                sg.popup('Erro', 'Todos os campos devem ser preenchidos!')
            else:
                sg.popup('Processo iniciado...')
                try:
                    processar_emails(email_user, email_pass,
                                     imap_server, save_folder, planilha_path)
                    sg.popup('Processo concluído com sucesso!')
                except imaplib.IMAP4.error as e:
                    sg.popup('Erro de autenticação', f'Erro de IMAP: {e}')
                except Exception as e:
                    sg.popup('Erro', f'Ocorreu um erro: {e}')

    janela.close()

# Função para processar emails


def processar_emails(email_user, email_pass, imap_server, save_folder, planilha_path):
    # Conectar ao servidor de email
    try:
        print(f"Conectando ao servidor IMAP: {imap_server}")
        mail = imaplib.IMAP4_SSL(imap_server)
        print("Fazendo login...")
        mail.login(email_user, email_pass)
        print("Login bem-sucedido")
    except imaplib.IMAP4.error as e:
        sg.popup('Erro de autenticação', f'Não foi possível fazer login: {e}')
        return
    except Exception as e:
        sg.popup('Erro de conexão', f'Erro ao conectar ao servidor IMAP: {e}')
        return

    mail.select("inbox")

    status, messages = mail.search(None, 'ALL')
    email_ids = messages[0].split()

    # Criar a planilha
    wb = Workbook()
    ws = wb.active
    ws.append(["Código de Barras", "Nome do Arquivo",
              "Data de Vencimento", "Valor"])

    # Processar cada email
    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                if msg.is_multipart():
                    for part in msg.walk():
                        content_disposition = str(
                            part.get("Content-Disposition"))
                        if "attachment" in content_disposition:
                            filename = part.get_filename()
                            if filename:
                                filepath = os.path.join(save_folder, filename)
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                if filename.lower().endswith(".pdf"):
                                    processar_pdf(filepath, ws)
    # Salvar a planilha
    wb.save(planilha_path)
    mail.logout()

# Função para processar PDFs e extrair informações


def processar_pdf(pdf_path, worksheet):
    text = extract_text(pdf_path)

    # Exemplo de regex para encontrar o código de barras, data de vencimento e valor
    codigo_de_barras = re.search(r'\d{44}', text)
    data_vencimento = re.search(r'\d{2}/\d{2}/\d{4}', text)
    valor = re.search(r'R\$\s*\d+,\d{2}', text)

    if codigo_de_barras and data_vencimento and valor:
        worksheet.append([
            codigo_de_barras.group(0),
            os.path.basename(pdf_path),
            data_vencimento.group(0),
            valor.group(0)
        ])


if __name__ == '__main__':
    main()
