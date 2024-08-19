
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
            


def enviar_email():
    pass


def main():
    try:

        load_dotenv()

        usuario_email = os.getenv('USER_EMAIL')
        usuario_email_senha = os.getenv('USER_PASSWORD')

        print(usuario_email)


        extrair_dados_pdf()

        


        # move_arquivos()
        
   
            

    except Exception as ex:
        print(ex)
    
    finally:
        pass









if __name__ == '__main__':
    main()