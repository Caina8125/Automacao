import pandas as pd
from tkinter import filedialog

def gerar_planilha():
    planilha_fhasso = filedialog.askopenfilename()
    planilha_gdf = filedialog.askopenfilename()
    df_plan_fhasso = pd.read_excel(planilha_fhasso)
    df_plan_fhasso = df_plan_fhasso.loc[df_plan_fhasso["Recupera ?"] == "Sim"]
    df_plan_fhasso = df_plan_fhasso.sort_values(by=['Fatura Inicial'])
    df_plan_gdf = pd.read_excel(planilha_gdf)
    lista = []

    for index, linha in df_plan_fhasso.iterrows():
        numero_senha_fhasso = str(linha['Autorização']).replace('.0', '')
        numero_op_fhasso = str(linha['Nro. Guia']).replace('.0', '')
        fatura_inicial = str(linha['Fatura Inicial']).replace('.0', '')
        fatura_recurso = str(linha['Fatura Recurso']).replace('.0', '')
        realizado = linha['Realizado']
        procedimento = str(linha['Procedimento']).replace('.0', '')
        valor_original = linha['Valor Original']
        valor_glosa_fhasso = str(linha['Valor Glosa']).replace('-', '')
        controle = str(linha['Controle']).replace('.0', '')

        for i, l in df_plan_gdf.iterrows():
            lista_linha = []
            numero_senha_plan_gdf = str(l['Autorização Origem']).replace('.0', '')
            lote_prestador = str(l['Lote Prestador']).replace('.0', '')
            data_de_atendimento_gdf = l['Data de Realização']
            codigo_gdf = str(l['Código']).replace('.0', '')
            valor_apresentado = l['Vl Apresentado (R$)']
            valor_glosa_gdf = str(l['Vl de Glosa (R$)'])
            autorizacao_nova = str(l['Autorização']).replace('.0', '')

            comparacao_senha = numero_senha_fhasso == numero_senha_plan_gdf
            comparacao_numero = numero_op_fhasso == numero_senha_plan_gdf
            comparacao_fatura = fatura_inicial == lote_prestador
            comparacao_data = realizado == data_de_atendimento_gdf
            comparacao_codigo = procedimento == codigo_gdf
            comparacao_valor_original = valor_original == valor_apresentado
            comparacao_valor_glosa = valor_glosa_fhasso == valor_glosa_gdf

            if (comparacao_senha or comparacao_numero) and comparacao_fatura and comparacao_data and comparacao_codigo and comparacao_valor_original and comparacao_valor_glosa:
                lista_linha = [int(controle), int(autorizacao_nova), int(numero_senha_fhasso), int(numero_op_fhasso), int(fatura_inicial), int(fatura_recurso)]
                break

        lista.append(lista_linha)

    cabecalho = ['Controle', 'Autorização Nova', 'Autorização Original', 'Nro. Guia', 'Fatura Inicial', 'Fatura Recurso']

    df_nova_planilha = pd.DataFrame(lista)
    df_nova_planilha.columns = cabecalho
    df_nova_planilha.to_excel(r'C:\Users\lucas.paz\Documents\Teste\GDF.xlsx', index=False)