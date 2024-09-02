#### SHAREPOINT #########################
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.creation_information import FileCreationInformation
import os

def autenticacao_sharepoint(site_url):
    
    # Credenciais de login
    username = 'user@dominio'  # Substitua pelo seu usuário do SharePoint
    password = 'suasenha'  # Substitua pela sua senha do SharePoint

    # Contexto de autenticação
    ctx_auth = AuthenticationContext(site_url)
    ctx_auth.acquire_token_for_user(username, password)
    ctx = ClientContext(site_url, ctx_auth)
    
    return ctx


def log_erro(lista_erro,file_name,target_folder_url,ctx):
    '''lista_erro = tipo Lista, é o conteudo q será escrito\n
       file_name = tipo string, é o Nome do arq\n
       target_folder_url = tipo string, é caminho de destino\n
       ctx = é o contexto sharepoint, obter da função de autenticação'''
       
    file_info = FileCreationInformation()
    # Conteúdo do arquivo a ser enviado
    file_content = "\n".join(lista_erro)
    file_info.content = file_content.encode('utf-8')  # Codificar o conteúdo do arquivo
    file_info.url = target_folder_url + '/' + file_name

    # Upload do arquivo
    target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
    uploaded_file = target_folder.files.add(file_info.url,file_info.content)
    ctx.execute_query()
    
    
def upload_file(original_file,target_folder_url,ctx):
    '''
    original_file = tipo string, é o caminho do arq que será feito upload\n
    target_folder_url = tipo string, é caminho de destino\n
    ctx = é o contexto sharepoint, obter da função de autenticação'''
    # Convertendo o dataframe para um arquivo Excel em memória
    # abre o arquivo localmente
    with open(original_file,"rb") as content_file:
        file_content = content_file.read()
    # Define nome do arquivo e path (target)
    dir, name = os.path.split(target_folder_url)
    # Escreve arquivo no sharepoint
    file = ctx.web.get_folder_by_server_relative_url(dir).upload_file(name, file_content).execute_query()
    
    