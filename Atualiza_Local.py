import os
import time
import shutil
import tkinter.messagebox
from pathlib import Path

exeLocal = r"C:\Instaladores\setup.exe"
exeAtualiza = r"\\10.0.0.239\atualiza\Automacao\setup.exe"

pathAtualiza = r"\\10.0.0.239\atualiza\Automacao"
pathLocal = r"C:\Automação"



def dataLocal():
    global trasnfLocal

    data = os.path.getmtime(exeLocal)
    dataFormat = time.ctime(data)
    objLocal = time.strptime(dataFormat)
    trasnfLocal = time.strftime("%d-%m-%y %H:%M:%S", objLocal) 
    # print("PathLocal:",trasnfLocal)
    return trasnfLocal

def dataAtualiza():
    global trasnfAtualiza

    data = os.path.getmtime(exeAtualiza)
    dataFormat = time.ctime(data)
    objAtualiza = time.strptime(dataFormat)
    trasnfAtualiza = time.strftime("%d-%m-%y %H:%M:%S", objAtualiza)
    # print("PathAtualiza:",trasnfAtualiza)
    return trasnfAtualiza

def executarAtualizacao():
    time.sleep(2)
    print("Subindo atualização, aguarde...")
    shutil.copytree(pathAtualiza, pathLocal)
    print("Atualização executada")

def apagarAtualiza():
    try:
        time.sleep(3)
        shutil.rmtree(pathLocal)
        print('Automação antiga apagada do seu disco')
        time.sleep(2)
    except Exception as e:
        print(f"{e.__class__.__name__}: {e}")

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