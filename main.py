from botcity.plugins.email import BotEmailPlugin
import tabula
import os
import shutil
from dotenv import load_dotenv



def extrair_dados_pdf():
    """
    Extrai dados de uma tabela em um arquivo PDF e exporta os dados para um arquivo Excel.

    1. Lê o arquivo PDF localizado como 'criacao_valor.pdf', especificamente a página 10.
    2. Captura a primeira tabela encontrada na página especificada.
    3. Remove linhas e colunas totalmente vazias da tabela.
    4. Define as colunas da tabela com os nomes: 'RUBRICAS', '2009', '2010', '2011', '2012'.
    5. Reseta os índices da tabela.
    6. Exporta a tabela resultante para um arquivo Excel chamado 'dados.xlsx', sem incluir os índices.
    """
    #Lendo o pdf 
    tabelas = tabula.read_pdf("criacao_valor.pdf", pages=10)

    #Capturando o primeiro resultado
    df_resultado = tabelas[0]

    #Excluir linhas totalmente vazias
    df_resultado = df_resultado.dropna(how='all', axis=0)
    #Excluir colunas totalmente vazias
    df_resultado = df_resultado.dropna(how='all', axis=1)

    #Criando o cabeçalho da tabela
    df_resultado.columns = ['RUBRICAS', '2009', '2010', '2011', '2012']
    
    #Rezetando os index da tabela
    df_resultado = df_resultado.reset_index(drop=True)

    #Exportando para o excel
    df_resultado.to_excel("dados.xlsx", index=False)

def move_arquivos():
    """
    Move arquivos de acordo com suas extensões para diretórios específicos.

    1. Verifica se os diretórios 'Arquivos PDF' e 'Arquivos Excel' existem. Se não existirem, cria-os.
    2. Move arquivos com extensão '.xlsx' para o diretório 'Arquivos Excel'.
    3. Move arquivos com extensão '.pdf' para o diretório 'Arquivos PDF'.
    """

    diretorio = os.listdir()

    if 'Arquivos PDF' not in diretorio:
        os.mkdir('Arquivos PDF')

    if 'Arquivos Excel' not in diretorio:
        os.mkdir('Arquivos Excel')


    for arquivo in diretorio:

        if arquivo.endswith('.xlsx'):
            shutil.move(arquivo, 'Arquivos Excel')

        if arquivo.endswith('.pdf'):
            shutil.move(arquivo, 'Arquivos PDF')
            
def enviar_email():
    """
    Envia um e-mail com um arquivo em anexo e busca por e-mails específicos.

    1. Carrega as variáveis de ambiente para o e-mail e senha do aplicativo.
    2. Configura o plugin de e-mail para IMAP e SMTP com o servidor do Gmail.
    3. Realiza o login com o e-mail e senha fornecidos.
    4. Busca e imprime e-mails com o assunto 'Test Message'.
    5. Envia um e-mail para os destinatários especificados com um anexo e conteúdo HTML.
    6. Desconecta dos servidores IMAP e SMTP após o envio.
    """
    
    load_dotenv()

    email_usuario = os.getenv('USER_EMAIL')
    app_password = os.getenv('USER_PASSWORD')

    # Instantiate the plugin
    email = BotEmailPlugin()
    
    # Configure IMAP with the gmail server
    email.configure_imap("imap.gmail.com", 993)

    # Configure SMTP with the gmail server
    email.configure_smtp("imap.gmail.com", 587)

    # Login with a valid email account
    email.login(email_usuario, app_password)

    # Search for all emails with subject: Test Message
    messages = email.search('SUBJECT "Test Message"')

    # For each email found: prints the date, sender address and text content of the email
    for msg in messages:
        print("\n---------------------------")
        print("Date => " + msg.date_str)
        print("From => " + msg.from_)
        print("Msg => " + msg.text)

    # Defining the attributes that will compose the message
    to = ["matheusinicial@gmail.com", "ceramarpinheiro1@gmail.com"]
    subject = "Esse é um teste para o e-mail"
    body = """<h1>Olá Matheus</h1>

            <p>É com muito orgulho que conseguimos extrair os dados do PDF e mandar por e-mail</p>

            <h3>Parabéns 🚀</h3>"""
    files = ["dados.xlsx"]

    # Sending the email message
    email.send_message(subject, body, to, attachments=files, use_html=True)

    # Close the conection with the IMAP and SMTP servers
    email.disconnect()

def main():

    try:
        extrair_dados_pdf()
        enviar_email()

    except Exception as ex:
        print(ex)
    
    finally:       
        move_arquivos()
        print('Finalizando o processo...')


if __name__ == '__main__':
    main()