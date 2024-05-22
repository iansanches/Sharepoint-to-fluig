from office365_api import SharePoint
import sys, os
from pathlib import PurePath
import time
from datetime import datetime, timedelta
import modulo_fluig
import shutil
import logging
import yaml

data_hoje = datetime.now().strftime(r"%d-%m-%Y")
# Configuração básica do logger
logging.basicConfig(
    filename=f'log/app-{data_hoje}.log',  # Especifica o arquivo de log
    level=logging.INFO,  # Define o nível mínimo de gravidade das mensagens a serem registradas
    format='%(asctime)s - %(levelname)s - %(message)s'  # Formato das mensagens de log
)

with open(r'C:\Users\rpa\Documents\POC-SHAREPOINT\config.yaml', 'r', encoding='utf=8') as params:
    config = yaml.safe_load(params)
    sharepoint_doc = config['sharepoint']['sharepoint_doc_library']
    pasta_download = config['sharepoint']['pasta_local_download']

#Pasta do sharepoint
FOLDER_NAME = sharepoint_doc #r'COMUNICAO' #SHAREPOINT_DOC_LIBRARY no yaml
#Pasta destino download
FOLDER_DEST = pasta_download #r'C:\Users\rpa\Documents\POC-SHAREPOINT\download'
# Determina se são pastas e subpastas
CRAWL_FOLDERS = "Yes"

def limpar_pasta_download():
    """Função para limpar a pasta de download do projeto"""
    #Verifica se a pasta existe
    if os.path.exists(FOLDER_DEST):
        shutil.rmtree(FOLDER_DEST)
        logging.info("A pasta de download e todo o seu conteúdo foram excluídos.")
    else:
        logging.info("A pasta de download não existe.")

def save_file(file_n, file_obj, subfolder):
    """Função para salvar arquivo localmente e no fluig"""
    dir_path = PurePath(FOLDER_DEST, subfolder)
    file_dir_path = PurePath(dir_path, file_n)
    with open(file_dir_path, 'wb') as f:
        f.write(file_obj)
    try:
        modulo_fluig.main(file_dir_path)
        logging.info(f"Arquivo salvo {file_dir_path}")
    except Exception as e:
        logging.error(f"Erro na hora de gravar no fluig: {e}")

def create_dir(path):
    """Função para criar pasta se não existir"""
    dir_path = PurePath(FOLDER_DEST, path)
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

def get_file(file_n, folder):
    "Função para obter o arquivo"
    file_obj = SharePoint().download_file(file_n, folder)
    save_file(file_n, file_obj, folder)
 
def get_files(folder):
    """Função para pegar os arquivos apenas pelos recentes, considerando 2 dias, configurado no timedelta"""
    files_list = SharePoint()._get_files_list(folder)
    logging.info(f"Lista de arquivos: {files_list}")
    data_atual = datetime.now()
    data_menos_2 = data_atual - timedelta(days=2)

    for file in files_list:
        #Condição verificando a propriedade do Sharepoint de Tempo da ultima modificação
        if file.time_last_modified >= data_menos_2:
            get_file(file.name, folder)
        else:
            logging.info(f"Arquivo não é recente {file.name}")
                
def get_folders(folder):
    """Função para pegar uma lista de subpastas de uma pasta"""
    l = []
    folder_obj = SharePoint().get_folder_list(folder)
    for subfolder_obj in folder_obj:
        subfolder = '/'.join([folder, subfolder_obj.name])
        l.append(subfolder)
    return l

def main():
    limpar_pasta_download()
    start_time = time.time()
    logging.info(f"Inicio: {datetime.now()}")

    try:
        if CRAWL_FOLDERS == 'Yes':
            folder_list = get_folders(FOLDER_NAME)
            for folder in folder_list:
                for subfolder in get_folders(folder):
                    folder_list.append(subfolder)
                    
            folder_list[:0] = [FOLDER_NAME]
            logging.info(f"Pastas: {folder_list}")

            for folder in folder_list:
                # Cria a pasta se ela não existir
                create_dir(folder)
                # Pega os arquivos especificos dessa pasta
                logging.info(f"PASTA CRIADA {folder}, tentar criar os arquivos")
                get_files(folder)
        else:
            get_files(FOLDER_NAME)
    except Exception as e:
        logging.error(f"Erro: {e}")
        return e
    
    
    elapsed_time = time.time() - start_time
    logging.info(f"Tempo decorrido: {elapsed_time}")
    logging.info(f"Fim: {datetime.now()}")

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logging.error(f"Erro: {e}")

    
