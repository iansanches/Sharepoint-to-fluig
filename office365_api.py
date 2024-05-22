from urllib import response
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import datetime
import yaml
import logging
import time

data_hoje = datetime.datetime.now().strftime(r"%d-%m-%Y")
# Configuração básica do logger
logging.basicConfig(
    filename=f'log/app-{data_hoje}.log',  # Especifica o arquivo de log
    level=logging.INFO,  # Define o nível mínimo de gravidade das mensagens a serem registradas
    format='%(asctime)s - %(levelname)s - %(message)s'  # Formato das mensagens de log
)

with open(r'C:\Users\rpa\Documents\POC-SHAREPOINT\config.yaml', 'r', encoding='utf=8') as params:
    config = yaml.safe_load(params)
    USERNAME = config['sharepoint']['sharepoint_email']
    PASSWORD = config['sharepoint']['sharepoint_password']
    SHAREPOINT_SITE = config['sharepoint']['sharepoint_url_site']
    SHAREPOINT_SITE_NAME = config['sharepoint']['sharepoint_site_name']
    SHAREPOINT_DOC = config['sharepoint']['sharepoint_doc_library']
    PASTA_DOWNLOAD = config['sharepoint']['pasta_local_download']

class SharePoint:
        
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(
                USERNAME,
                PASSWORD
            )
        )
        return conn
    
    def _get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{folder_name}'
        for _ in range(3):  # Tenta 3 vezes
            try:
                root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
                root_folder.expand(["Files", "Folders"]).get().execute_query()
                logging.info(f"Consulta de arquivos executada com sucesso")
                return root_folder.files
            except Exception as e:
                logging.error(f"Erro ao consultar arquivos da pasta: {e}")
                logging.info("Tentando baixar novamente em 10 segundos...")
                time.sleep(10)  # Espera 5 segundos antes de tentar novamente
        # Se todas as tentativas falharem, levanta uma exceção ou retorna None
        logging.error("Não foi possível consultar os arquivos após várias tentativas.")
        return e
        
    
    def get_folder_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{folder_name}'
        
        for _ in range(3):  # Tenta 3 vezes
            try:
                root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
                root_folder.expand(["Folders"]).get().execute_query()
                return root_folder.folders
            except Exception as e:
                logging.error(f"Erro ao consultar lista de pastas: {e}")
                logging.info("Tentando consultar novamente em 10 segundos...")
                time.sleep(10)  # Espera 5 segundos antes de tentar novamente
        # Se todas as tentativas falharem, levanta uma exceção ou retorna None
        logging.error("Não foi possível consultar as pastas após várias tentativas.")
        return e

    def download_file(self, file_name, folder_name):  
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{folder_name}/{file_name}'

        for _ in range(3):  # Tenta 3 vezes
            try:
                file = File.open_binary(conn, file_url)
                logging.info(f"Download feito com sucesso, arquivo: {file_name}")
                return file.content
            except Exception as e:
                logging.error(f"Erro ao baixar arquivo: {e}")
                logging.info("Tentando baixar novamente em 10 segundos...")
                time.sleep(10)  # Espera 5 segundos antes de tentar novamente
        
        # Se todas as tentativas falharem, levanta uma exceção ou retorna None
        logging.error("Não foi possível baixar o arquivo após várias tentativas.")
        return e
    
    def download_latest_file(self, folder_name):
        date_format = "%Y-%m-%dT%H:%M:%SZ"
        files_list = self._get_files_list(folder_name)
        file_dict = {}
        for file in files_list:
            dt_obj = datetime.datetime.strptime(file.time_last_modified, date_format)
            file_dict[file.name] = dt_obj
        # sort dict object to get the latest file
        file_dict_sorted = {key:value for key, value in sorted(file_dict.items(), key=lambda item:item[1], reverse=True)}    
        latest_file_name = next(iter(file_dict_sorted))
        content = self.download_file(latest_file_name, folder_name)
        return latest_file_name, content
        

    def upload_file(self, file_name, folder_name, content):
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.upload_file(file_name, content).execute_query()
        return response
    
    def upload_file_in_chunks(self, file_path, folder_name, chunk_size, chunk_uploaded=None, **kwargs):
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.files.create_upload_session(
            source_path=file_path,
            chunk_size=chunk_size,
            chunk_uploaded=chunk_uploaded,
            **kwargs
        ).execute_query()
        return response
    
    def get_list(self, list_name):
        conn = self._auth()
        target_list = conn.web.lists.get_by_title(list_name)
        items = target_list.items.get().execute_query()
        return items
        
    def get_file_properties_from_folder(self, folder_name):
        files_list = self._get_files_list(folder_name)
        properties_list = []
        for file in files_list:
            file_dict = {
                'file_id': file.unique_id,
                'file_name': file.name,
                'major_version': file.major_version,
                'minor_version': file.minor_version,
                'file_size': file.length,
                'time_created': file.time_created,
                'time_last_modified': file.time_last_modified
            }
            properties_list.append(file_dict)
            file_dict = {}
        return properties_list
    

