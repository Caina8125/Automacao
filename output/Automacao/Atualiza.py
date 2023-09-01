import os
import time
import shutil
import tkinter.messagebox

pathLocal = r"C:\Automacao\output\Automacao\Automacao.exe"
pathAtualiza = r"\\10.0.0.239\Atualiza\Automacao\output\Automacao\Automacao.exe"

def dataLocal():
    global trasnfLocal

    data = os.path.getmtime(pathLocal)
    dataFormat = time.ctime(data)
    objLocal = time.strptime(dataFormat)
    trasnfLocal = time.strftime("%d-%m-%y %H:%M:%S", objLocal) 
    print("PathLocal:",trasnfLocal)

def dataAtualiza():
    global trasnfAtualiza

    data = os.path.getmtime(pathAtualiza)
    dataFormat = time.ctime(data)
    objAtualiza = time.strptime(dataFormat)
    trasnfAtualiza = time.strftime("%d-%m-%y %H:%M:%S", objAtualiza)
    print("PathAtualiza:",trasnfAtualiza)


def executarAtualizacao():
    shutil.copy2(pathAtualiza, pathLocal)
    print("Cópia executada")

def Script():

    dataLocal()
    dataAtualiza()
    if(trasnfAtualiza == trasnfLocal):
        print("Mesma Hora")
    else:
        tkinter.messagebox.showinfo('Atualização','Uma atualização foi feita, o sistema será reiniciado')
        print("executar atualização")

        executarAtualizacao()
        dataLocal()
        dataAtualiza()