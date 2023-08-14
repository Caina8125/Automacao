from tkinter import *
import tkinter as tk 
from tkinter import ttk
from Buscar_fatura import *
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

        # foto = tk.PhotoImage(file="loader.gif")
        # self.img = tk.Label(self.quintoContainer,image=foto, border=0, background="white")
        # self.img.foto = foto
        # self.img["pady"] = 5
        # self.img.pack()

        # self.dadosAutomacao = Label(self.quintoContainer, background="white")
        # self.dadosAutomacao["text"] = self.dadosTela
        # self.dadosAutomacao.pack()

        self.buttonIniciar = Button(self.sextoContainer, bg="#274360",foreground="white",width=10)
        self.buttonIniciar["text"] = "Iniciar"
        self.buttonIniciar["command"] = self.chamarAutomacao
        self.buttonIniciar.pack(side=LEFT)

        # self.buttonCancelar = Button(self.sextoContainer, background="red",foreground="white",width=10)
        # self.buttonCancelar["text"] = "Parar"
        # self.buttonCancelar["command"] = self.pararAutomacao
        # self.buttonCancelar.pack()



    def chamarAutomacao(self):
        automacao = self.comboBox.get()
        
        if automacao == "Financeiro - Buscar Faturas":
            iniciar()

    # def pararAutomacao(self):

    #     sys.exit(self.chamarAutomacao)

    # def dadosTela(self):
    #     getResposta()
                

Script()

root = tk.Tk()
Application(root)
root.title('AMHP - Automações')
root.geometry("500x300")
root.configure(background="white")
root.resizable(width=0, height=0)

# ctypes.windll.kernel32.FreeConsole()
root.mainloop()
