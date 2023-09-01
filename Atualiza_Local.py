import os
import time
import shutil
import tkinter.messagebox
from pathlib import Path

exeLocal = r"C:\Automacao\output\Automacao\Automacao.exe"
exeAtualiza = r"\\10.0.0.239\Atualiza\Automacao\output\Automacao\Automacao.exe"

pathAtualiza = r"\\10.0.0.239\atualiza\Automacao\output\Automacao"
pathLocal = r"C:\Automacao\output\Automacao"



def dataLocal():
    global trasnfLocal

    data = os.path.getmtime(exeLocal)
    dataFormat = time.ctime(data)
    objLocal = time.strptime(dataFormat)
    trasnfLocal = time.strftime("%d-%m-%y %H:%M:%S", objLocal) 
    print("PathLocal:",trasnfLocal)

def dataAtualiza():
    global trasnfAtualiza

    data = os.path.getmtime(exeAtualiza)
    dataFormat = time.ctime(data)
    objAtualiza = time.strptime(dataFormat)
    trasnfAtualiza = time.strftime("%d-%m-%y %H:%M:%S", objAtualiza)
    print("PathAtualiza:",trasnfAtualiza)

def executarAtualizacao():
    time.sleep(2)
    print("Subindo atualização, aguarde...")
    shutil.copytree(pathAtualiza, pathLocal)
    print("Atualização executada")

def apagarAtualiza():
    try:
        shutil.rmtree(pathLocal)
        print('Automação antiga apagada do seu disco')
        time.sleep(2)
    except:
        pass

def Script():

    dataLocal()
    dataAtualiza()

    if(trasnfAtualiza == trasnfLocal):
        print("Mesma Hora")

    else:
        tkinter.messagebox.showinfo('Atualização','Uma atualização foi feita, o sistema será reiniciado')
        print("executar atualização")
        apagarAtualiza()
        executarAtualizacao()
        dataLocal()
        dataAtualiza()