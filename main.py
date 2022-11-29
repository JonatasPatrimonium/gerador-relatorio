import math
import os

import requests
import pandas as pd
from bs4 import BeautifulSoup

URL_RESP = "https://app.acessorias.com/respdptos.php?geraR&fieldFilters=Dpt_8,Dpt_2,Dpt_1,Dpt_20,Dpt_3&modo=VNT"
URL_GERAL = 'https://app.acessorias.com/relempresas.php?tc=E&modo=xlsCpl'
PATH_EXCEL = "entrada.xlsx"
PATH_XLS = 'ResponsaveisDptos.xls'
COOKIE_FILE = "cookie.txt"
USERAGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.5304.63 Safari/537.36"
OUTFILE = 'relatorio'


def write_file(data, file, op):
    try:
        with open(file, op) as outfile:
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

    try:
        response = requests.post(url, headers=header)
        if response.status_code == 200:
            if response.text.find("expirada") != -1:
                return False

            return response

        else:
            return False

    except:
        return False


# função que trabalha o HTML para conversão
def work_html(html):
    lista = []  # Lista de empresas
    titulo_rep = []  # Lista dos titulos de representantes

    # Extrai todas as tags 'tr' do html
    try:
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
    except:
        return False


def read_plan(plan):
    try:
        dataframe = pd.read_excel(plan, engine='openpyxl', skiprows=[0, 1])
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

    except:
        return False


def work_data(list_json, list_plan):
    lista = []
    try:
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
    except:
        return False


def export_excel(path, lista):
    try:
        dataframe = pd.DataFrame(lista)
        dataframe = dataframe.sort_values('ID')
        dataframe.to_excel(path, index=False)
        return True
    except:
        return False


if __name__ == '__main__':

    lista_geral = False
    data_resp = False
    ops_sim = ['sim', 'SIM', 'Sim', 's', 'S', 'Yes', 'yes', 'y', 'Y', 'SIm', 'siM', 'sIM', 'sIm']

    try:
        print('[ GERADOR DE RELATÓRIO ]\n\n')
        print('Usar arquivos locais? [ Padrão: Não ]')
        op = input('SIM (S) ou NÃO (N): ')
        if op in ops_sim:
            while True:
                file_resp = input('Relatorio de Responsáveis [ Padrão: {} ]\nInsira o nome do arquivo: '.format(PATH_XLS)) or PATH_XLS
                data_resp = read_file(file_resp)
                if data_resp:
                    file_geral = input('Relatorio Completo [ Padrão: {} ]\nInsira o nome do arquivo: '.format(PATH_EXCEL)) or PATH_EXCEL
                    lista_geral = read_plan(file_geral)
                    if lista_geral:
                        break
                    else:
                        print('Erro ao tentar ler o arquivo [ {} ]\nPor favor tente novamente'.format(file_geral))
                else:
                    print('Erro ao tentar ler o arquivo [ {} ]\nPor favor tente novamente'.format(file_resp))

        else:
            cookie = read_file(COOKIE_FILE)
            if not cookie:
                write_file('', 'cookie.txt', 'wt')

            while True:
                print('[ Buscando dados do Relatório de Responsáveis... ]')
                response = get_html(URL_RESP, cookie)

                if not response:
                    print('[ Sessão Expirada! ]')
                    cookie = input('Insira aqui os novos cookies: ')
                    write_file(cookie, COOKIE_FILE, 'wt')
                    continue

                data_resp = response.text

                if data_resp:
                    print('[ Buscando dados do Relatório Completo... ]')
                    response = get_html(URL_GERAL, cookie)
                    if response:
                        write_file(response.content, 'temp.xlsx', 'wb')
                        lista_geral = read_plan('temp.xlsx')
                        if lista_geral:
                            if os.path.exists('temp.xlsx'):
                                os.remove('temp.xlsx')
                            break

                    print('Ocorreu um erro ao tentar buscar o Relatório Completo!\nTentando novamente...')

                else:
                    print('Ocorreu um erro ao tentar buscar o Relatório de Responsáveis!\nTentando novamente...')

        if data_resp and lista_geral:
            print('[ Extraindo dados... ]')
            lista_resp = work_html(data_resp)
            if lista_resp:
                lista_final = work_data(lista_resp, lista_geral)
                if lista_final:
                    while True:
                        outfile = input('Arquivo Final [ Padrão: {}.xlsx ]\nInsira o nome do arquivo a ser gerado (sem a extensão): '.format(OUTFILE)) or OUTFILE
                        status = export_excel('{}.xlsx'.format(outfile), lista_final)
                        if status:
                            print('Aquivo salvo com sucesso!\n[ {}.xlsx ]'.format(outfile))
                            break

                        print('Erro ao tentar salvar o arquivo!\nTente novamente')

    except KeyboardInterrupt:
        print('\n\nPrograma finalizado pelo usuario!')

    print('[ FIM! ]')
