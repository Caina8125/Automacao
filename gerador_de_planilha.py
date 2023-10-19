import pandas as pd
from tkinter import filedialog

planilha_fhasso = filedialog.askopenfilename()
df_plan_fhasso = pd.read_excel(planilha_fhasso)

planilha_gdf = filedialog.askopenfilename()
df_plan_gdf = pd.read_excel(planilha_gdf)

for index, linha in df_plan_fhasso.iterrows():
    lista = []
    numero_senha_fhasso = str(linha['Autorização']).replace('.0', '')
    numero_op_fhasso = str(linha['Nro. Guia']).replace('.0', '')
    fatura_inicial = ...
    realizado = linha['Realizado']
    procedimento = str(linha['Procedimento']).replace('.0', '')
    valor_original = linha['Valor Original']
    valor_glosa_fhasso = str(linha['Valor Glosa']).replace('-', '')

    for i, l in df_plan_gdf.iterrows():
        numero_senha_plan_gdf = str(l['Autorização Origem']).replace('.0', '')
        lote_prestador = ...
        data_de_atendimento_gdf = l['Data de Atendimento']
        codigo_gdf = str(l['Código']).replace('.0', '')
        valor_apresentado = l['Vl Apresentado (R$)']
        valor_glosa_gdf = l['Vl de Glosa (R$)']

        comparacao_senha = numero_senha_fhasso == numero_senha_plan_gdf
        comparacao_numero = numero_op_fhasso == numero_senha_plan_gdf
        comparacao_data = realizado == data_de_atendimento_gdf
        comparacao_codigo = procedimento == codigo_gdf
        comparacao_valor_original = valor_original == valor_apresentado
        comparacao_valor_glosa = valor_glosa_fhasso == valor_glosa_gdf

        if (comparacao_senha or comparacao_numero) and comparacao_data and comparacao_codigo and comparacao_valor_original and comparacao_valor_glosa:
            lista_linha = [l['Autorização'], numero_senha_plan_gdf]



