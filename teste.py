import pandas as pd

planilha = r"C:\Users\lucas.paz\Desktop\GDF - Sincronização\copy\Extrato de glosa - GDF.xlsx"
df = pd.read_excel(planilha)
autorizacao = df['Autorização']
autorizacao_origem = df['Autorização Origem']
fatura_df = df.loc[(df['Autorização'] == 3037832)]

if fatura_df.empty == True:
    fatura_df = df.loc[(df['Autorização'] == '3037832')]

print(autorizacao)
print('')
print(autorizacao_origem)