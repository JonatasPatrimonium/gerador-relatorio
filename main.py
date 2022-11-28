import math

import requests
import pandas as pd
from bs4 import BeautifulSoup

#   Constantes: URL, USER-AGENT, NOME DO ARQUIVO DE ENTRADA, NOME DO ARQUIVO DE SAIDA E NOME DO ARQUIVO DE COOKIES
URL_RESP = "https://app.acessorias.com/respdptos.php?geraR&fieldFilters=Dpt_8,Dpt_2,Dpt_1,Dpt_20,Dpt_3&modo=VNT"
PATH_EXCEL = "entrada.xlsx"
PATH_XLS = 'ResponsaveisDptos.xls'
COOKIE_FILE = "cookie.txt"
USERAGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.5304.63 Safari/537.36"
OUTFILE = 'relatorio.xlsx'


def write_file(data, file):
    try:
        with open(file, 'wt') as outfile:
            outfile.write(data)

        return True
    except Exception as e:

        return False

def read_file(file):
    try:
        with open(file, 'rt') as outfile:
            data = outfile.read()

        return data

    except Exception as e:
        return False


def get_html(url, cookie):
    header = {"User-Agent": USERAGENT, "Cookie": cookie}

    print("Realizando Requisição!\n[ Por Favor Aguarde... ]")

    try:
        response = requests.post(url, headers=header)
        if response.status_code == 200:
            if response.text.find("expirada") != -1:
                print('Sessão expirada!\n[ Por Favor Renove os Cookies ]')
                return -1

            print("Requisição Realizada com Sucesso!")
            return response.text

        else:
            print("Falha ao tentar realizar a requisição!\n[ Status: {} ]".format(response.status_code))
            return False
    except Exception as e:
        print('Erro ao tentar realizar a requisição!\n[ ERRO: {} ]'.format(e))
        return False


# função que trabalha o HTML para conversão
def work_html(html):
    lista = []  # Lista de empresas
    titulo_rep = []  # Lista dos titulos de representantes

    # Extrai todas as tags 'tr' do html
    soup = BeautifulSoup(html, 'html.parser')
    tr_list = soup.find_all('tr')

    for tr in tr_list:
        tr_text = str(tr)
        # Extrai todas as tags 'td' contidas em 'tr' referentes aos titulos dos responsaveis
        tds_titulo = tr.find_all('td', {'rowspan': '2'})
        for td_titulo in tds_titulo:
            titulo = str(td_titulo.text)
            titulo = titulo.replace('\xa0', '')
            titulo_rep.append(titulo)

        if tr_text.find("rowspan") == -1 and tr_text.find("<strong>") == -1:

            # Extrai todas as tags 'td' contidas em 'tr' referentes ao conteudo
            tr_soup = BeautifulSoup(tr_text, 'html.parser')
            td_list = tr_soup.find_all('td')
            i_td = 0
            responsaveis = dict()
            nome = ''
            id = ''
            cnpj = ''
            for i, td in enumerate(td_list):
                td_str = str(td)
                dicionario = dict()
                # Armazena o nome dos responsaveis em um dicinario (responsaveis) dentro do objeto empresa
                if td_str.find("colspan") == -1:
                    responsaveis[titulo_rep[i_td]] = td.text
                    i_td += 1
                    if i_td >= 16:
                        i_td = 0

                # Armazena o nome e o id (cnpj) no objeto empresa
                elif td_str.find('align="left"') == -1:
                    nome = td_str.replace('<td colspan="2" style="width:30%;">', '')
                    i_id = nome.find('[')
                    f_id = nome.find(']')
                    id = nome[i_id + 1:f_id]
                    id = int(id)
                    nome = nome[:i_id - 1]
                    cnpj = td.small.text

            dicionario['id'] = id
            dicionario['nome'] = nome
            dicionario['cnpj'] = cnpj
            dicionario['responsaveis'] = responsaveis
            lista.append(dicionario)

    return lista


def read_plan(plan):
    try:
        dataframe = pd.read_excel(plan, engine='openpyxl')
        lista = []
        for i, row in dataframe.iterrows():
            dicionario = dict()
            if not math.isnan(row['ID ']):
                dicionario['id'] = int(row['ID '])
                dicionario['razao'] = row['Razão social ']
                dicionario['cnpj'] = row['CNPJ']
                dicionario['regime'] = row['Regime']
                dicionario['cidade'] = row['Cidade']
                dicionario['uf'] = row['UF ']
                dicionario['cadastro'] = row['Cadastro']
                dicionario['cli'] = row['Cli. até']
                dicionario['ativa'] = row['Ativa?']
                dicionario['grupo'] = row['Grupo de Empresas']
                dicionario['tags'] = row['Tags']

                lista.append(dicionario)

        return lista

    except Exception as e:
        return False


def work_data(list_json, list_plan):
    lista = []
    for i, data_plan in enumerate(list_plan):
        dicionario = dict()
        for i, data_json in enumerate(list_json):
            if data_json['id'] == data_plan['id']:
                dicionario['ID'] = data_plan['id']
                dicionario['Ativa?'] = data_plan['ativa']
                dicionario['Cadastro'] = data_plan['cadastro']
                dicionario['CNPJ'] = data_plan['cnpj']
                dicionario['Razão Social'] = data_plan['razao']
                dicionario['Regime'] = data_plan['regime']
                dicionario['Cidade'] = data_plan['cidade']
                dicionario['Cli. até'] = data_plan['cli']
                dicionario['Grupo de Empresas'] = data_plan['grupo']
                dicionario['Tags'] = data_plan['tags']
                dicionario['Customer Success'] = data_json['responsaveis']['1_1 - Customer Success']
                dicionario['Fiscal'] = data_json['responsaveis']['1_2 - Fiscal']
                dicionario['Contábil'] = data_json['responsaveis']['1_3 - Cont\u00e1bil']
                dicionario['Pessoal - Folha'] = data_json['responsaveis']['Pessoal - Folha']
                dicionario['Pessoal - Impostos'] = data_json['responsaveis']['Pessoal - Impostos']

        lista.append(dicionario)

    return lista


def export_excel(path, lista):
    try:
        dataframe = pd.DataFrame(lista)
        dataframe = dataframe.sort_values('ID')
        dataframe.to_excel(path, index=False)
        return True
    except:
        return False

if __name__ == '__main__':
    ops_sim = ['sim', 'SIM', 'Sim', 's', 'S', 'Yes', 'yes', 'y', 'Y', 'SIm', 'siM', 'sIM', 'sIm']
    print('Você deseja utilizar um arquivo XLS ou HTML contendo o relatório dos responsáveis? ')
    op = input('SIM (S) ou NÃO (N): ')

    if op in ops_sim:
        infile = input('INSIRA O NOME DO ARQUIVO DE ENTRADA [ RESPONSAVEIS ]:\n[ Padrão: {} ] '.format(PATH_XLS)) or PATH_XLS
        if infile:
            print('Realizando a leitura do Arquivo de Entrada!\n[ {} ]'.format(infile))
            html = read_file(infile)
        else:
            html = False

    else:
        print('Realizando a leitura dos Cookies!\n[ {} ]'.format(COOKIE_FILE))
        cookie = read_file(COOKIE_FILE)
        if cookie:
            html = get_html(URL_RESP, cookie)
            while html == -1:
                print('Deseja reanovar os cookies agora?')
                op = input('SIM (S) ou NÃO (N): ')
                if op in ops_sim:
                    cookie = input('Insira os novos cookies aqui: ')
                    write_file(cookie, COOKIE_FILE)
                    cookie = read_file(COOKIE_FILE)
                    print('Novos cookies inseridos!')
                    html = get_html(URL_RESP, cookie)

                else:
                    html = False
        else:
            print('Erro ao tentar ler o arquivo de cookies!')
            html = False

    if html:

        intfile_geral = input('INSIRA O NOME DO ARQUIVO DE ENTRADA [ RELATORIO GERAL ]:\n[ Padrão {} ] '.format(PATH_EXCEL)) or PATH_EXCEL
        lista_plan = read_plan(intfile_geral)

        if lista_plan:
            print("[ Extraindo Dados... ]")
            lista_resp = work_html(html)

            lista_final = work_data(lista_resp, lista_plan)

            res = False
            while not res:
                outfile = input('Insira o nome do arquivo final [ Relatrio ]:\n[ Padrão: {} ] '.format(OUTFILE)) or OUTFILE
                print("Salvando Arquivo de Saída!\n[ {} ]".format(outfile))
                res = export_excel(outfile, lista_final)

            print("FIM!\n[ Operação realizada com sucesso ]")

        else:
            print('Erro ao tentar ler o arquivo de entrada!\n[ Exporte o relatorio completo e salve como {} ]'.format(PATH_EXCEL))

    else:
        print("Erro ao tentar ler o arquivo informado: {}".format(infile))