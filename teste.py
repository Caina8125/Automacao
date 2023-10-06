import pandas as pd
from tkinter import filedialog

planilha = filedialog.askopenfilename()

df = pd.read_excel(planilha, header=5)
df = df.iloc[:-1]
df = df.dropna()
df['Concluído'] = ''

for index, linha in df.iterrows():
    df.loc[index, 'Concluído'] = 'Sim'
    print(df)