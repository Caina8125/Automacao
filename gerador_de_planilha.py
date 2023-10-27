import pandas as pd
from tkinter import filedialog
import tkinter.messagebox

def verificar_frame_vazio(df_filtrado, df, coluna, numero):
    if df_filtrado.empty == True:
        fatura_df = df.loc[(df[coluna] == numero)]
        return fatura_df
        
    else:
        return df_filtrado

def gerar_planilha():
    try:
        tkinter.messagebox.showinfo('Planilha Recurso', 'Selecione a planilha do relatório por data de pagamento.')
        planilha_fhasso = filedialog.askopenfilename()
        tkinter.messagebox.showinfo('Planilha Detalhado Unificado', 'Selecione a planilha do relatório Detalhado Normal/Especial (Unificado).')
        planilha_unificada = filedialog.askopenfilename()
        tkinter.messagebox.showinfo('Planilha Demonstrativo', 'Selecione a planilha do demonstrativo do GDF')
        planilha_gdf = filedialog.askopenfilename()
        df_plan_fhasso = pd.read_excel(planilha_fhasso)
        df_plan_fhasso = df_plan_fhasso.loc[df_plan_fhasso["Recupera ?"] == "Sim"]
        df_plan_fhasso = df_plan_fhasso.sort_values(by=['Fatura Inicial'])
        df_plan_fhasso['Controle Recurso'] = ''
        df_plan_unificada = pd.read_excel(planilha_unificada)
        df_plan_unificada['INSERIDO'] = ''
        df_plan_gdf = pd.read_excel(planilha_gdf)
        lista = []
        count_nao_encontradas = 0

        for index, linha in df_plan_fhasso.iterrows():
            numero_senha_fhasso = str(linha['Autorização']).replace('.0', '')
            numero_op_fhasso = str(linha['Nro. Guia']).replace('.0', '')
            numero_amhptiss = str(linha['Amhptiss']).replace('.0', '')
            fatura_recurso = str(linha['Fatura Recurso']).replace('.0', '')
            fatura_inicial = str(linha['Fatura Inicial']).replace('.0', '')
            realizado = linha['Realizado']
            procedimento = str(linha['Procedimento']).replace('.0', '')
            valor_original = str(linha['Valor Original'])
            valor_glosa = str(linha['Valor Glosa']).replace('-', '')
            fatura_df_filtrado = df_plan_unificada.loc[(df_plan_unificada['PROCESSOID'] == int(fatura_recurso))]
            fatura_df = verificar_frame_vazio(fatura_df_filtrado, df_plan_unificada, 'PROCESSOID', fatura_recurso)
            fatura_df_gdf_filtrado = df_plan_gdf.loc[(df_plan_gdf['Lote Prestador'] == int(fatura_inicial))]
            fatura_df_gdf = verificar_frame_vazio(fatura_df_gdf_filtrado, df_plan_gdf, 'Lote Prestador', fatura_inicial)

            if fatura_df.empty == True or fatura_df_gdf.empty == True:
                count_nao_encontradas += 1
                continue

            for i, l in fatura_df.iterrows():
                primeira_validacao = False
                senha_plan_uni = str(l['AUTORIZACAO']).replace('.0', '')
                numero_op_plan_uni = str(l['GUIAATENDIMENTO']).replace('.0', '')
                hdi_registro = str(l['HDIREGISTRO']).replace('.0', '')
                fatura_recurso_plan_uni = str(l['PROCESSOID']).replace('.0', '')
                realizado_plan_uni = l['DATAREALIZADO']
                procedimento_plan_uni = str(l['CODIGOID']).replace('.0', '')
                controle_plan_uni = str(l['ATENDIMENTOID']).replace('.0', '')

                validacao_senha_oper = numero_senha_fhasso == senha_plan_uni or numero_op_fhasso == numero_op_plan_uni
                valicacao_amhptiss = numero_amhptiss == hdi_registro
                validacao_fatura = fatura_recurso == fatura_recurso_plan_uni
                validacao_data = realizado == realizado_plan_uni
                validacao_procedimento = procedimento == procedimento_plan_uni

                if validacao_senha_oper and validacao_fatura and valicacao_amhptiss and validacao_data and validacao_procedimento:
                    df_plan_fhasso.loc[index, 'Controle Recurso'] = controle_plan_uni
                    fatura_df.loc[i, 'INSERIDO'] = 'Sim'
                    primeira_validacao = True
                    break

                else:
                    continue

            if primeira_validacao == False:
                count_nao_encontradas += 1
                continue

            for ind, lin in fatura_df_gdf.iterrows():
                segunda_comparacao = False
                numero_senha_plan_gdf = str(lin['Autorização Origem']).replace('.0', '')
                data_de_atendimento_gdf = lin['Data de Realização']
                codigo_gdf = str(lin['Código']).replace('.0', '')
                autorizacao_nova = str(lin['Autorização']).replace('.0', '')
                vl_apresentado = str(lin['Vl Apresentado (R$)']).replace('R$ ', '')
                vl_de_glosa = str(lin['Vl de Glosa (R$)']).replace('R$ ', '')

                comparacao_senha = numero_senha_fhasso == numero_senha_plan_gdf
                comparacao_numero = numero_op_fhasso == numero_senha_plan_gdf
                comparacao_data = realizado == data_de_atendimento_gdf
                comparacao_codigo = procedimento == codigo_gdf
                comparacao_valor_original = valor_original == vl_apresentado
                comparacao_valor_glosa = valor_glosa == vl_de_glosa

                if (comparacao_senha or comparacao_numero) and comparacao_data and comparacao_codigo and comparacao_valor_original and comparacao_valor_glosa:
                    controle = df_plan_fhasso['Controle Recurso'][index]
                    lista_linha = [controle, autorizacao_nova, numero_senha_fhasso, numero_op_fhasso, fatura_inicial, fatura_recurso]
                    segunda_comparacao = True
                    break
            
            if segunda_comparacao == False:
                count_nao_encontradas += 1
                continue

            lista.append(lista_linha)

        cabecalho = ['Controle', 'Autorização Nova', 'Autorização Original', 'Nro. Guia', 'Fatura Inicial', 'Fatura Recurso']

        df_nova_planilha = pd.DataFrame(lista)
        df_nova_planilha.columns = cabecalho
        df_nova_planilha.to_excel(r'C:\Users\lucas.paz\Documents\Teste\GDF.xlsx', index=False)
        tkinter.messagebox.showinfo("Gerador de Planilha", f"Planilha gerada! \n Total de linhas não encontradas {count_nao_encontradas}")
    
    except Exception as e:
        tkinter.messagebox.showerror("Gerador de Planilha", f"Ocorreu uma exceção não tratada \n {e.__class__.__name__} - {e}")