from tkinter import *
import tkinter as tk 
from tkinter import ttk
import threading
from PIL import Image, ImageTk
from itertools import count, cycle
import ctypes

from Buscar_fatura import iniciar
from Atualiza import *


class Application:
    def __init__(self, master=None):
        self.fontePadrao = ("Arial", "10")
        self.primeiroContainer = Frame(master, background="white")
        self.primeiroContainer.pack()

        self.segundoContainer = Frame(master, background="white")
        self.segundoContainer["padx"] = 20
        self.segundoContainer.pack()

        self.terceiroContainer = Frame(master, background="white")
        self.terceiroContainer["padx"] = 20
        self.terceiroContainer["pady"] = 10
        self.terceiroContainer.pack()

        self.quartoContainer = Frame(master, background="white")
        self.quartoContainer["padx"] = 20
        self.quartoContainer.pack()

        self.quintoContainer = Frame(master, background="white")
        self.quintoContainer["padx"] = 20
        self.quintoContainer["padx"] = 10
        self.quintoContainer.pack()

        self.sextoContainer = Frame(master, background="white")
        self.sextoContainer["padx"] = 100
        self.sextoContainer["pady"] = 28
        self.sextoContainer.pack()

        self.cabecalho = Label(self.primeiroContainer, bg="#274360")
        self.cabecalho["padx"] = 340
        self.cabecalho["pady"] = 12
        self.cabecalho.pack()

        foto = tk.PhotoImage(file="logo.png")
        
        self.img = tk.Label(self.segundoContainer,image=foto, border=0)
        self.img.foto = foto
        self.img["pady"] = 5
        self.img.pack()

        self.nomeLabel = Label(self.terceiroContainer, text="Selecione a automação",font=self.fontePadrao, background="white")
        self.nomeLabel.pack(side=LEFT)

        self.comboBox = ttk.Combobox(self.quartoContainer, values=["Glosa - Auditoria Geap", "Faturamento - Anexar Honorario Geap", "Financeiro - Buscar Faturas"], width=50)
        self.comboBox["background"] = 'white'
        self.comboBox.pack(side=LEFT)

        self.buttonIniciar = Button(self.sextoContainer, bg="#274360",foreground="white",width=10, command=lambda: threading.Thread(target=self.chamarAutomacao).start())
        self.buttonIniciar["text"] = "Iniciar"
        self.buttonIniciar.pack(side=LEFT)

    def gif(self):
        
        self.info = Label(self.quintoContainer, text= "Trabalhando...",font=('Arial,10,bold'), background="white")
        self.info.pack()

        self.lbl = ImageLabel(self.quintoContainer,background="white")
        self.lbl.pack(side=LEFT)
        self.lbl.load('loader2.gif')

    def ocultar(self):
        self.buttonIniciar.pack_forget()

    def desocultar(self):
        self.lbl.pack_forget()
        self.buttonIniciar.pack(side=LEFT)

    def botaoHistorico(self):
        self.info.pack_forget()

        self.buttonHistorico = Button(self.sextoContainer, bg="#274360",foreground="white",width=10)
        self.buttonHistorico["text"] = "Histórico"
        self.buttonHistorico.pack(side=LEFT)

    def chamarAutomacao(self):
        self.ocultar()
        try:
            self.info.pack_forget()
            self.buttonHistorico.pack_forget()
        except:
            pass

        self.gif()

        automacao = self.comboBox.get()
        if automacao == "Financeiro - Buscar Faturas":
            iniciar()

        self.desocultar()

        self.botaoHistorico()
     
#---------------------------------------------------------------------------------------------------------
#Classe do Gif

class ImageLabel(tk.Label):
   
    def load(self, im):
        if isinstance(im, str):
            im = Image.open(im)
        frames = []

        try:
            for i in count(1):
                frames.append(ImageTk.PhotoImage(im.copy()))
                im.seek(i)
        except EOFError:
            pass
        self.frames = cycle(frames)

        try:
            self.delay = im.info['duration']
        except:
            self.delay = 100

        if len(frames) == 1:
            self.config(image=next(self.frames))
        else:
            self.next_frame()

    def unload(self):
        self.config(image=None)
        self.frames = None

    def next_frame(self):
        if self.frames:
            self.config(image=next(self.frames))
            self.after(self.delay, self.next_frame)


#-------------------------------------------------------------------------------------------
    

Script()

root = tk.Tk()
Application(root)
root.title('AMHP - Automações')
root.geometry("500x300")
root.configure(background="white")
root.resizable(width=False, height=False)

ctypes.windll.kernel32.FreeConsole()
root.mainloop()