from requests_oauthlib import OAuth1Session
import yaml
import logging
from time import sleep
from datetime import datetime

data_hoje = datetime.now().strftime(r"%d-%m-%Y")
# Configuração básica do logger
logging.basicConfig(
    filename=f'log/app-{data_hoje}.log',  # Especifica o arquivo de log
    level=logging.INFO,  # Define o nível mínimo de gravidade das mensagens a serem registradas
    format='%(asctime)s - %(levelname)s - %(message)s'  # Formato das mensagens de log
)

PESQUISA_PASTA = '1'
PESQUISA_ARQUIVO = '2'
PARENT_ID_PASTA_ENGETEC = '1553' # PASTA 8 é a correta, 1553 é a pasta de teste

with open('config.yaml', 'r', encoding='utf=8') as params:
    config = yaml.safe_load(params)
    CLIENT_KEY = config['fluig']['client_key']
    CLIENT_SECRET = config['fluig']['client_secret']
    RESOURCE_OWNER_KEY = config['fluig']['resource_owner_key']
    RESOURCE_OWNER_SECRET = config['fluig']['resource_owner_secret']
    DOMINIO = config['fluig']['dominio']


def envia_arquivo(nome_arquivo, id_pasta, caminho_arquivo):
    """Função para enviar uma arquivo para o fluig"""
    # Inicialize a sessão OAuth1
    oauth = OAuth1Session(CLIENT_KEY, CLIENT_SECRET, RESOURCE_OWNER_KEY, RESOURCE_OWNER_SECRET) 
    #Url da api do Fluig para criar arquivo
    url = fr'{DOMINIO}/content-management/api/v2/documents/upload/{nome_arquivo}/{id_pasta}/publish'
    #Abrir o arquivo em formato binario
    with open(caminho_arquivo, 'rb') as file:
        file_content = file.read()
    #Chamada a api
    response = oauth.post(url, files={'file': file_content})
    return response
    

def cria_pasta(nome_pasta, parent_id):
    """Função para criar uma pasta no fluig"""
    # Inicialize a sessão OAuth1
    oauth = OAuth1Session(CLIENT_KEY, CLIENT_SECRET, RESOURCE_OWNER_KEY, RESOURCE_OWNER_SECRET)
    #Url da api do Fluig para criar pasta
    url = fr'{DOMINIO}/content-management/api/v2/folders/{parent_id}'
    body = {"alias": nome_pasta}
    #Chamada a api
    logging.info(f"Chamada a api para criar pasta ({nome_pasta}), url: {url}")
    response = oauth.post(url, json=body)
    return response

def verifica_existencia_arquivo(nome_arquivo):
    """Função para verificar a existência do arquivo pelo nome"""
    # Inicialize a sessão OAuth1
    oauth = OAuth1Session(CLIENT_KEY, CLIENT_SECRET, RESOURCE_OWNER_KEY, RESOURCE_OWNER_SECRET)
    #Url da api do Fluig para procurar arquivo
    url= f'{DOMINIO}/dataset/api/v2/dataset-handle/search?datasetId=document&field=documentPK.documentId&constraintsField=documentDescription&constraintsField=documentType&constraintsField=deleted&constraintsInitialValue={nome_arquivo}&constraintsInitialValue={PESQUISA_ARQUIVO}&constraintsInitialValue=false&constraintsFinalValue={nome_arquivo}&constraintsFinalValue={PESQUISA_ARQUIVO}&constraintsFinalValue=false&constraintsType=MUST&constraintsType=MUST'
    #Chamada a api
    logging.info(f"Chamada a api para verificar existencia de arquivo ({nome_arquivo}), url: {url}")
    response = oauth.get(url)
    return response

def verifica_existencia_pasta(item_lista, parent_id):
    """Função para verificar a existência de uma pasta pelo nome, partindo da origem como sempre sendo '3013 - ENGETEC OPERAÇÃO' """
    # Inicialize a sessão OAuth1
    oauth = OAuth1Session(CLIENT_KEY, CLIENT_SECRET, RESOURCE_OWNER_KEY, RESOURCE_OWNER_SECRET)
    #Url da api do Fluig para procurar pasta
    url= f'{DOMINIO}/dataset/api/v2/dataset-handle/search?datasetId=document&field=documentPK.documentId&constraintsField=documentDescription&constraintsField=documentType&constraintsField=deleted&constraintsField=parentDocumentId&constraintsInitialValue={item_lista}&constraintsInitialValue={PESQUISA_PASTA}&constraintsInitialValue=false&constraintsInitialValue={parent_id}&constraintsFinalValue={item_lista}&constraintsFinalValue={PESQUISA_PASTA}&constraintsFinalValue=false&constraintsFinalValue={parent_id}&constraintsType=MUST&constraintsType=MUST&constraintsType=MUST'
    #Chamada a api
    response = oauth.get(url)
    return response

def get_documento(documento_id):
    """Pega dados do documento"""
    # Inicialize a sessão OAuth1
    oauth = OAuth1Session(CLIENT_KEY, CLIENT_SECRET, RESOURCE_OWNER_KEY, RESOURCE_OWNER_SECRET)
    #Url da api do Fluig para pegar dados do documento
    url = f'{DOMINIO}/content-management/api/v2/documents/{documento_id}'
    #Chamada a api    
    logging.info(f"Chamada a api para verificar documento ({documento_id}), url: {url}")
    response = oauth.get(url)
    return response 

def update_arquivo(documento_id, caminho_arquivo, id_pasta, nome_arquivo):
    """Atualiza o arquivo"""
    # Inicialize a sessão OAuth1
    oauth = OAuth1Session(CLIENT_KEY, CLIENT_SECRET, RESOURCE_OWNER_KEY, RESOURCE_OWNER_SECRET)
    #Url da api do Fluig para deletar o documento
    url = f'{DOMINIO}/content-management/api/v2/documents/{documento_id}' 
    #Chamada a api de deletar
    logging.info(f"Chamada a api para deletar documento ({documento_id}), url: {url}")
    response = oauth.delete(url)
    if response.status_code == 204:
        logging.info(f"Arquivo deletado: ({documento_id})")
        logging.info(f"Chamada a api para deletar documento ({documento_id}), url: {url}")
        #Chamada a api para gravar arquivo
        logging.info(f"Chamada a api para enviar documento ({documento_id}), url: {url}")
        response = envia_arquivo(nome_arquivo, id_pasta, caminho_arquivo)
        return response
    else:
        raise Exception(f"Arquivo não foi deletado, status_code:{response.status_code}")

def verifica_pasta_anterior(lista:list, parent_id):
    """Função para verificar se as pastas anteriores são do arquivo em questão, devido ao FLUIG cada pasta ter um ID """
    lista = lista[:-1]
    lista.reverse()
    pasta_correta = True
    logging.info(f"Lista com as pastas invertida {lista}")

    for index, item in enumerate(lista):
        #Verifica se não é a ultima pasta
        if index != len(lista) - 1:
            #Primeira pasta
            response = get_documento(parent_id)
            if response.status_code == 200:
                resposta = response.json()
                description = resposta['description']
                parent_id = resposta['parentId']
                #Segunda Pasta
                proximo_item = lista[index + 1]
                response = get_documento(parent_id)
                if response.status_code == 200:
                    resposta = response.json()
                    description_pasta_anterior = resposta['description']
                    #Verifica se a descrição da pasta atual bate com a descrição, e se a proxima pasta bate com a proxima descrição
                    if description == item and description_pasta_anterior == proximo_item:
                        logging.info("Pasta Correta")
                        pasta_correta = True
                    else:
                        logging.info("Pasta errada")
                        pasta_correta = False
                        break
                else:
                    raise Exception(f"Erro na consulta de get_documento, status{response.status_code}, body: {response.text}")
            else:
                raise Exception(f"Erro na consulta de get_documento, status{response.status_code}, body: {response.text}")
        else:
            break
    return pasta_correta       

def main(diretorio_arquivo):
    logging.info("Iniciando a verificação no fluig")
    diretorio_str = str(diretorio_arquivo)
    lista = []
    #Lista criada a partir do nome das pastas, pegando a 6 pq ta seguindo essa maquina a 5 é download
    #ex: C:\Users\rpa\Documents\POC-SHAREPOINT\download\COMUNICAO
    lista = diretorio_str.split("\\")[6:] 
    #Nome Arquivo sempre será o ultimo elemento da lista
    nome_arquivo = lista[-1]
    #Procurar o arquivo
    response = verifica_existencia_arquivo(nome_arquivo)
    existe_arquivo = False

    if response.status_code == 200:
        resposta_values = response.json()
        logging.info(f"Requisição de verificar existencia do arquivo bem sucedida, resposta: {resposta_values}")
        if resposta_values['values']:
            values = resposta_values['values']
            #Percorrendo cada valor pois pode ter nomes de arquivos iguais em pastas diferentes
            for valor in values:
                documento_id = valor['documentPK.documentId']
                response = get_documento(documento_id)
                #Verifica o status 
                if response.status_code == 200:
                    #Resposta do arquivo
                    resposta = response.json()
                    logging.info(f"Requisição de verificar o arquivo bem sucedida, resposta: {resposta} ")
                    parent_id = resposta['parentId'] 
                    pasta_correta = verifica_pasta_anterior(lista, parent_id)
                    #Se o arquivo estiver na Pasta correta ele irá excluir o arquivo e subir o novo
                    if pasta_correta:
                        logging.info(f"Existe o arquivo na pasta correta: ({nome_arquivo})")
                        update_arquivo(documento_id, diretorio_arquivo, parent_id, nome_arquivo)
                        existe_arquivo = True
                        break  
                else:
                    raise f"Erro na consulta do arquivo {response.status_code} - {response.text}"
        #Se não retornar valor pra consulta de arquivo, ou se o arquivo nao existe na pasta correta (pode existir em outra pasta)  
        if not resposta_values['values'] or (resposta_values and not existe_arquivo):
            parent_id = PARENT_ID_PASTA_ENGETEC #TODO esta seguindo atualmente a pasta "Teste"
            #Verificar se cada pasta existe a partir da pasta da ENGETEC
            for item in lista[0:-1]:
                response = verifica_existencia_pasta(item, parent_id)
                resposta = response.json()

                if resposta['values']:
                    #Pega o primeiro resultado como parte sempre da mesma pasta, no máximo só pode ter um resultado, levando em consideração que está vindo do sharepoint, pois no fluig tem como sim ter duas pastas com mesmo nome
                    documento_id = resposta['values'][0]['documentPK.documentId']
                    #Se a Pasta existir ele atribui o valor dela como parent_id, pra procurar a próxima
                    parent_id = documento_id
                else:
                    logging.info(f"Criando a pasta {item}")
                    response = cria_pasta(item, parent_id)
                    resposta = response.json()
                    documento_id = resposta['documentId']
                    parent_id = documento_id
            
            #Após criar pastas criar o arquivo
            logging.info(f"Criando arquivo: {nome_arquivo} no diretorio {diretorio_arquivo}")
            response = envia_arquivo(nome_arquivo, parent_id, diretorio_arquivo)

            if response.status_code == 200:
                logging.info(f"Arquivo ({nome_arquivo}) gravado com sucesso")
            else:
                logging.error(f"Arquivo não gravado")
                logging.error(response.status_code)
                logging.error(response.text)

    else:
        logging.error("Requisição para verificação de arquivo falhou")
        logging.error(response.text)                      


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logging.error(f"Erro no fluig {e}")
        raise e





