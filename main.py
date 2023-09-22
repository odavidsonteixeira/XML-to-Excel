import xmltodict
import os
import pandas as pd


# import json    Biblioteca que transforma um dicionário em json


def pegar_infos(nome_arquivo, valores):
    # print(f'Pegou as informações de {nome_arquivo}')
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)

        if "NFe" in dic_arquivo:
            infos_nfe = dic_arquivo['NFe']['infNFe']
        else:
            infos_nfe = dic_arquivo['nfeProc']['NFe']['infNFe']

        numero_nota = infos_nfe['@Id']
        empresa_emissora = infos_nfe['emit']['xNome']
        nome_cliente = infos_nfe['dest']['xNome']
        endereco = infos_nfe['dest']['enderDest']
        if 'vol' in infos_nfe['transp']:
            peso_bruto = infos_nfe['transp']['vol']['pesoB']
        else:
            peso_bruto = "Não informado!"

        valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso_bruto])


lista_arquivos = os.listdir("nfs")

colunas = ["Número da nota", "Empresa emissora", "Nome do cliente", "Endereço", "Peso bruto"]
valores = []

for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("NotasFiscais.xlsx", index=False)

