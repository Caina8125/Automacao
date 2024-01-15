import PyPDF2
from tabula.io import read_pdf
from os import listdir
from tkinter.filedialog import askdirectory

def extrair_texto_pdf_pypdf2(caminho_do_pdf):
    lista_de_arquivos = [f'{caminho_do_pdf}\\{arquivo}' for arquivo in listdir(caminho_do_pdf)]
    for arquivo in lista_de_arquivos:
        with open(arquivo, 'rb') as arquivo_pdf:
            leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
            for pagina_numero in range(len(leitor_pdf.pages)):
                pagina = leitor_pdf.pages[pagina_numero]
                texto_sem_quebra = ' '.join(pagina.extract_text().split('\n'))
                vetor_retira_espaços = texto_sem_quebra.split(' ')
                if pagina_numero == 0:
                    protocolo = vetor_retira_espaços[0] 
                    numero_fatura_pagina1 = vetor_retira_espaços[37].split('_')[0].replace('0000000000000','')
                    print(protocolo, numero_fatura_pagina1)
                else:
                    numero_fatura_pagina2 = vetor_retira_espaços[329].replace('Informações','')
                    print(numero_fatura_pagina2)
            if numero_fatura_pagina1 == numero_fatura_pagina2:
                print('Verdadeiro')

carta_remessa = read_pdf(r'C:\Users\lucas.paz\Documents\LUCAS PAULO ROBÔ\44977\44977.pdf', pages="1")
caminho_do_pdf = askdirectory()
extrair_texto_pdf_pypdf2(caminho_do_pdf)