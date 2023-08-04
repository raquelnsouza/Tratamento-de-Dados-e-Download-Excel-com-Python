# IMPORTANDO AS BIBLIOTECAS

import pandas as pd
from io import BytesIO
import os
from flask import Flask, send_file
import requests


# INICIANDO O FLASK
# pois era necessário rodar o script em um servidor web
app = Flask(__name__)

@app.route('/ROTA')
def download_excel():

    # Extraindo os dados da fonte
    link = 'https://www.site.com/data.xlsx'
    response = requests.get(link)
    file = BytesIO(response.content)

    # Lendo o arquivo
    df1 = pd.read_excel(file, sheet_name='aba-1', engine='openpyxl', thousands='.', decimal=',')
    df2 = pd.read_excel(file, sheet_name='aba-2', engine='openpyxl', thousands='.', decimal=',')
    df3 = pd.read_excel(file, sheet_name='aba-3', engine='openpyxl', thousands='.', decimal=',')
    df4 = pd.read_excel(file, sheet_name='aba-4', engine='openpyxl', thousands='.', decimal=',')

    # Convertendo as colunas para string
        # houveram alguns problemas com o tipo de dados da base de dados, então para resolver foi necessário transformar todas as abas em string
    df1 = df1.astype(str)
    df2 = df2.astype(str)
    df3 = df3.astype(str)
    df4 = df4.astype(str)

    # Mesclando as tabelas
    df_merged1 = pd.merge(df1, df2, left_on='_index', right_on='_parent_index', how='outer', suffixes=('_df1', '_df2'))
    df_merged2 = pd.merge(df_merged1, df3, left_on='_parent_index', right_on='_parent_index', how='outer',suffixes=('_merged1', '_df3'))
    df_merged3 = pd.merge(df_merged2, df4, left_on='_parent_index', right_on='_parent_index', how='outer',suffixes=('_merged2', '_df4'))

    # Função para garantir que apenas valores distintos serão concatenados com ","
        # Essa função foi criada por uma peculiaridade da base de dados, 
        # pois haviam abas com várias linhas para o mesmo ID e para esta solução 
        # era necessário que todos os dados do ID vinhessem na mesma linha separados por vírgula.

    def juntar_linhas(x):
        values = set()
        resultado = []
        for value in x.dropna():
            if value not in values:
                resultado.append(value)
                values.add(value)
        return ','.join(resultado)

    # Agrupando os dados que têm linhas extras pelo "_parent_index" e concatenando com ","
    df_merged1 = df_merged2.groupby('_parent_index').agg(juntar_linhas).reset_index()
    df_merged2 = df_merged3.groupby('_parent_index').agg(juntar_linhas).reset_index()

    # Salva os dados em um arquivo Excel
    output_path = os.path.abspath('NOME DO ARQUIVO.xlsx')
    df_merged2.to_excel(output_path, index=False)

    # Envia o arquivo Excel como resposta HTTP
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run()
