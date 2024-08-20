
from botcity.plugins.email import BotEmailPlugin
import tabula
import os
import shutil
from dotenv import load_dotenv



def extrair_dados_pdf():
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
            


def enviar_email(user, password, arquivo, destinatarios, assunto, corpo):
    email = BotEmailPlugin()

    # Configuração do servidor IMAP para leitura de e-mails
    email.configure_imap('imap.gmail.com', 993)

    # Configuração do servidor SMTP para envio de e-mails
    email.configure_smtp("imap.gmail.com", 587) 

    # Login com uma conta de e-mail válida
    email.login(user, password)

    # Pesquisa todos os e-mails com o assunto específico (opcional, você pode remover essa parte se não precisar)
    messages = email.search(f'SUBJECT "{assunto}"')

    # Para cada e-mail encontrado: imprime a data, o endereço do remetente e o conteúdo do e-mail
    for msg in messages:
        print("\n---------------------------")
        print("Date => " + msg.date_str)
        print("From => " + msg.from_)
        print("Msg => " + msg.text)

    # # Defining the attributes that will compose the message
    # to = ["<RECEIVER_ADDRESS_1>", "<RECEIVER_ADDRESS_2>"]
    # subject = "Hello World"
    # body = "<h1>Hello!</h1> This is a test message!"
    # files = [arquivo]

    # Enviando a mensagem de e-mail
    email.send_message(assunto, corpo, destinatarios, attachments=[arquivo], use_html=True)

    # Fecha a conexão com os servidores IMAP e SMTP
    email.disconnect()


def main():
    try:

        load_dotenv()

        email_usuario = os.getenv('EMAIL_USER')
        app_password = os.getenv('USER_PASSWORD')

        print(app_password)

      



        # extrair_dados_pdf()

        # assunto = 'Dados extraidos do PDF'
        # body = f"""
        #     <h1>Olá Matheus</h1>

        #     <p>É com muito orgulho que conseguimos extrair os dados do PDF e mandar por e-mail</p>

        #     <h3>Parabéns lindão 🚀</h3>

        # """

        # enviar_email(usuario_email, usuario_email_senha, 'dados.xlsx', 'matheusinicial@gmail.com',assunto, body)


        
   
            

    except Exception as ex:
        print(ex)
    
    finally:       
        move_arquivos()









if __name__ == '__main__':
    main()