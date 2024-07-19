# Robô de Processamento de Emails

Este é um script Python que processa emails, extrai anexos e cria uma planilha Excel com informações relevantes.

## Pré-requisitos

- Python 3.x instalado
- Bibliotecas: `PySimpleGUI`, `imaplib`, `email`, `os`, `re`, `pdfminer`, `openpyxl`

## Como Usar

1. Clone este repositório.
2. Execute o script `processar_emails.py`.
3. Preencha os campos na interface gráfica:
   - **Email**: Insira o seu endereço de email.
   - **Senha de Aplicativo**: Digite a senha de aplicativo (caso use autenticação em duas etapas).
   - **Servidor IMAP**: O servidor IMAP (padrão: `imap.gmail.com`).
   - **Pasta para salvar anexos**: Escolha a pasta onde deseja salvar os anexos dos emails.
   - **Caminho para a planilha Excel**: Defina o caminho para a planilha onde os dados serão armazenados.
4. Clique no botão "Iniciar" para processar os emails.

## Funcionalidades

- Conecta-se ao servidor de email via IMAP.
- Percorre os emails na caixa de entrada.
- Extrai anexos (PDFs) e salva na pasta especificada.
- Cria uma planilha Excel com as seguintes informações:
  - Código de Barras
  - Nome do Arquivo
  - Data de Vencimento
  - Valor

## Observações

- Certifique-se de preencher todos os campos corretamente.
- Caso ocorra algum erro, uma mensagem de erro será exibida na interface gráfica.

## Licença

Este projeto está sob a licença MIT.
