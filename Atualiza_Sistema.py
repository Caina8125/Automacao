import os
import time
import shutil
import tkinter.messagebox
from pathlib import PurePath

pathDev = r"C:\Users\lucas.paz\Desktop\Projetos\Automacao\Automacao"
pathAtualiza = r"\\10.0.0.239\atualiza\Automacao"



def apagarAtualiza():
    try:
        shutil.rmtree(pathAtualiza)
        print('Arquivos apagados da pasta atualiza')
        time.sleep(2)
    except:
        pass

def executarAtualizacao():
    time.sleep(2)
    print("Subindo atualização, aguarde...")
    shutil.copytree(pathDev, pathAtualiza)
    print("Cópia executada")

def Script():
    apagarAtualiza()
    executarAtualizacao()
    tkinter.messagebox.showinfo('Atualização','A atualização do seu sistema agora está na pasta atualiza')

Script()