from tkinter import *
import tkinter as tk 
from tkinter import ttk
import threading
from PIL import Image, ImageTk
from itertools import count, cycle
from Recursar_Duplicado import *
from Buscar_fatura import iniciar
from Atualiza_Local import *
from Anexar_Honorario import anexar_guias
from VerificarSituacao_BRB import verificacao_brb
from recursar_brb import recursar_brb
from recursar_cassi import recursar_cassi
from recursar_fascal import recursar_fascal
from recursar_evida import recursar_evida
from Recursar_SemDuplicado import recursar_sem_duplicado
from Recursar_Caixa import recursar_caixa
from Recurso_Benner import recursar_benner
from Recursar_SIS import recursar_sis
from GEAP_Conferencia import conferencia
from VerificarSituacao_Fascal import verificacao_fascal
from VerificarSituacao_Gama import verificar_gama
from Demonstrativo_Amil import demonstrativo_amil
from Demonstrativo_Brb import demonstrativo_brb
from Demonstrativo_Caixa import demonstrativo_caixa
from Demonstrativo_Cassi import demonstrativo_cassi
from Demonstrativo_Camara import demonstrativo_camara
from Demonstrativo_Camed import demonstrativo_camed
from Demonstrativo_Casembrapa import demonstrativo_casembrapa
from Demonstrativo_Cassi import demonstrativo_cassi
from Demonstrativo_Codevasf import demonstrativo_codevasf
from Demonstrativo_Evida import demonstrativo_evida
from Demonstrativo_Fapes import demonstrativo_fapes
from Demonstrativo_Fascal import demonstrativo_fascal
from Demonstrativo_Gama import demonstrativo_gama
from Demonstrativo_Life_Empresarial import demonstrativo_life
from Demonstrativo_Mpu import demonstrativo_mpu
from Demonstrativo_Pmdf import demonstrativo_pmdf
from Demonstrativo_Postal import demonstrativo_postal
from Demonstrativo_Real_Grandeza import demonstrativo_real
from Demonstrativo_Serpro import demonstrativo_serpro
from Demonstrativo_Sis import demonstrativo_sis
from Demonstrativo_Stf import demonstrativo_stf
from Demonstrativo_Tjdft import demonstrativo_tjdft
from Demonstrativo_Unafisco import demonstrativo_unafisco
from Gerar_relatorios_Brindes import Gerar_Relat_Normal
from Nota_Fiscal import subirNF
from gerador_de_planilha import gerar_planilha
from leitor_de_pdf_gama import pdf_reader
from Enviar_Pdf_Brb import enviar_pdf
from bacen_conferencia import conferir_bacen
from bacen_envio_xml import fazer_envio_xml
from datetime import datetime
import os

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
        self.quartoContainer["pady"] = 10
        self.quartoContainer.pack()
        

        self.quintoContainer = Frame(master, background="white")
        self.quintoContainer["padx"] = 20
        self.quintoContainer["pady"] = 5
        self.quintoContainer.pack()

        self.sextoContainer = Frame(master, background="white")
        self.sextoContainer["padx"] = 100
        self.sextoContainer["pady"] = 10
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

        self.comboBox = ttk.Combobox(self.quartoContainer, values=["Faturamento - Anexar Honorario Geap",
                                                                   "Faturamento - Conferência GEAP",
                                                                   "Faturamento - Conferência Bacen",
                                                                   "Faturamento - Enviar PDF Bacen",
                                                                   "Faturamento - Enviar PDF BRB",
                                                                   "Faturamento - Enviar XML Bacen",
                                                                   "Faturamento - Leitor de PDF GAMA",
                                                                   "Faturamento - Verificar Situação BRB",
                                                                   "Faturamento - Verificar Situação Fascal",
                                                                   "Faturamento - Verificar Situação Gama",
                                                                   "Financeiro - Buscar Faturas GEAP", 
                                                                   "Financeiro - Demonstrativos Amil", 
                                                                   "Financeiro - Demonstrativos BRB", 
                                                                   "Financeiro - Demonstrativos Câmara dos Deputados", 
                                                                   "Financeiro - Demonstrativos Camed", 
                                                                   "Financeiro - Demonstrativos Casembrapa", 
                                                                   "Financeiro - Demonstrativos Cassi", 
                                                                   "Financeiro - Demonstrativos Codevasf", 
                                                                   "Financeiro - Demonstrativos E-Vida", 
                                                                   "Financeiro - Demonstrativos Fapes", 
                                                                   "Financeiro - Demonstrativos Fascal", 
                                                                   "Financeiro - Demonstrativos Gama", 
                                                                   "Financeiro - Demonstrativos Life Empresarial", 
                                                                   "Financeiro - Demonstrativos MPU", 
                                                                   "Financeiro - Demonstrativos PMDF", 
                                                                   "Financeiro - Demonstrativos Postal", 
                                                                   "Financeiro - Demonstrativos Real Grandeza", 
                                                                   "Financeiro - Demonstrativos Saúde Caixa", 
                                                                   "Financeiro - Demonstrativos Serpro", 
                                                                   "Financeiro - Demonstrativos SIS", 
                                                                   "Financeiro - Demonstrativos STF", 
                                                                   "Financeiro - Demonstrativos TJDFT", 
                                                                   "Financeiro - Demonstrativos Unafisco", 
                                                                   "Glosa - Gerador de Planilha GDF",
                                                                   "Glosa - Recursar Benner(Câmara, CAMED, FAPES, Postal)",
                                                                   "Glosa - Recursar BRB",
                                                                   "Glosa - Recursar Cassi",
                                                                   "Glosa - Recursar E-VIDA",
                                                                   "Glosa - Recursar Fascal",
                                                                   "Glosa - Recursar GEAP Duplicado",
                                                                   "Glosa - Recursar GEAP Sem Duplicado",
                                                                   "Glosa - Recursar Saúde Caixa",
                                                                   "Glosa - Recursar Serpro",
                                                                   "Glosa - Recursar SIS",
                                                                   "Relatório - Brindes",
                                                                   "Tesouraria - Nota Fiscal"
                                                                    ], width=50)
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

    def ocultar_data(self):
        try:
            self.inserir_data_inicial.pack_forget()
            self.inserir_data_final.pack_forget()
            self.botao_ok.pack_forget()
            self.voltar.pack_forget()
        except:
            pass

    def desocultar(self):
        self.lbl.pack_forget()

    def botao_iniciar(self):
        self.buttonIniciar = Button(self.sextoContainer, bg="#274360",foreground="white",width=10, command=lambda: threading.Thread(target=self.chamarAutomacao).start())
        self.buttonIniciar["text"] = "Iniciar"
        self.buttonIniciar.pack(side=LEFT)

    def obter_datas(self):           
        global data_inicial, data_final, validacao
        data_inicial = self.inserir_data_inicial.get()
        data_final = self.inserir_data_final.get()
        validacao = data_valida(data_inicial, data_final)
        self.ocultar_data()

        if validacao:
            self.gif()
            demonstrativo_cassi(data_inicial, data_final)
            self.reiniciar()

        else:
            tkinter.messagebox.showerror( 'Data inválida!' , 'Digíte uma data válida')
            self.inserir_data()

    def inserir_data(self):
        self.inserir_data_inicial = tk.Entry(self.quintoContainer)
        self.inserir_data_inicial.insert(0, "Digite a data inicial")
        

        self.inserir_data_final = tk.Entry(self.quintoContainer)
        self.inserir_data_final.insert(0, "Digite a data final")

        self.botao_ok = Button(self.sextoContainer, bg="#274360",foreground="white", text="OK", command=lambda: threading.Thread(target=self.obter_datas).start())

        self.voltar = Button(self.sextoContainer, bg="#274360", foreground="white", text="Voltar", command=lambda: threading.Thread(target=self.voltar_inicio).start())

        self.inserir_data_inicial.pack(side=LEFT)
        self.inserir_data_inicial.bind("<FocusIn>", self.limpar_placeholder1)
        self.inserir_data_final.pack(side=LEFT)
        self.inserir_data_final.bind("<FocusIn>", self.limpar_placeholder2)
        self.botao_ok.pack(side=LEFT, padx=10)
        self.voltar.pack(side=RIGHT)

    def limpar_placeholder1(self, event):
        if self.inserir_data_inicial.get() == "Digite a data inicial":
            self.inserir_data_inicial.delete(0, "end")
            self.inserir_data_inicial.config(fg="black")

    def limpar_placeholder2(self, event):
        if self.inserir_data_final.get() == "Digite a data final":
            self.inserir_data_final.delete(0, "end")
            self.inserir_data_final.config(fg="black")

    # def botaoHistorico(self):
    #     self.info.pack_forget()

    #     self.buttonHistorico = Button(self.sextoContainer, bg="#274360",foreground="white",width=10)
    #     self.buttonHistorico["text"] = "Histórico"
    #     self.buttonHistorico.pack(side=LEFT)

    def chamarAutomacao(self):
        self.ocultar()
        try:
            self.info.pack_forget()
            # self.buttonHistorico.pack_forget()
        except:
            pass
        
        automacao = self.comboBox.get()

        match automacao:
            case "Faturamento - Anexar Honorario Geap":
                self.gif()
                anexar_guias()
                self.reiniciar()

            case "Faturamento - Conferência GEAP":
                self.gif()
                conferencia()
                self.reiniciar()

            case "Faturamento - Conferência Bacen":
                self.gif()
                conferir_bacen()
                self.reiniciar()

            case "Faturamento - Enviar PDF Bacen":
                self.reiniciar()

            case "Faturamento - Enviar PDF BRB":
                self.gif()
                enviar_pdf()
                self.reiniciar()

            case "Faturamento - Enviar XML Bacen":
                self.gif()
                fazer_envio_xml()
                self.reiniciar()

            case "Faturamento - Leitor de PDF GAMA":
                self.gif()
                pdf_reader()
                self.reiniciar()

            case "Faturamento - Verificar Situação BRB":
                self.gif()
                verificacao_brb()
                self.reiniciar()

            case "Faturamento - Verificar Situação Fascal":
                self.gif()
                verificacao_fascal()
                self.reiniciar()

            case "Faturamento - Verificar Situação Gama":
                self.gif()
                verificar_gama()
                self.reiniciar()

            case "Financeiro - Buscar Faturas GEAP":
                self.gif()
                iniciar()
                self.reiniciar()

            case "Financeiro - Demonstrativos Amil":
                self.gif()
                demonstrativo_amil()
                self.reiniciar()

            case "Financeiro - Demonstrativos BRB":
                self.gif()
                demonstrativo_brb()
                self.reiniciar()

            case "Financeiro - Demonstrativos Câmara dos Deputados":
                self.gif()
                demonstrativo_camara()
                self.reiniciar()

            case "Financeiro - Demonstrativos Camed":
                self.gif()
                demonstrativo_camed()
                self.reiniciar()

            case "Financeiro - Demonstrativos Casembrapa":
                self.gif()
                demonstrativo_casembrapa()
                self.reiniciar()

            case "Financeiro - Demonstrativos Cassi":
                self.inserir_data()

            case "Financeiro - Demonstrativos Codevasf":
                self.gif()
                demonstrativo_codevasf()
                self.reiniciar()

            case "Financeiro - Demonstrativos E-Vida":
                self.gif()
                demonstrativo_evida()
                self.reiniciar()

            case "Financeiro - Demonstrativos Fapes":
                self.gif()
                demonstrativo_fapes()
                self.reiniciar()

            case "Financeiro - Demonstrativos Fascal":
                self.gif()
                demonstrativo_fascal()
                self.reiniciar()

            case "Financeiro - Demonstrativos Gama":
                self.gif()
                demonstrativo_gama()
                self.reiniciar()

            case "Financeiro - Demonstrativos Life Empresarial":
                self.gif()
                demonstrativo_life()
                self.reiniciar()

            case "Financeiro - Demonstrativos MPU":
                self.gif()
                demonstrativo_mpu()
                self.reiniciar()

            case "Financeiro - Demonstrativos PMDF":
                self.gif()
                demonstrativo_pmdf()
                self.reiniciar()

            case "Financeiro - Demonstrativos Postal":
                self.gif()
                demonstrativo_postal()
                self.reiniciar()

            case "Financeiro - Demonstrativos Real Grandeza":
                self.gif()
                demonstrativo_real()
                self.reiniciar()

            case "Financeiro - Demonstrativos Saúde Caixa":
                self.gif()
                demonstrativo_caixa()
                self.reiniciar()

            case "Financeiro - Demonstrativos Serpro":
                self.gif()
                demonstrativo_serpro()
                self.reiniciar()

            case "Financeiro - Demonstrativos SIS":
                self.gif()
                demonstrativo_sis()
                self.reiniciar()

            case "Financeiro - Demonstrativos STF":
                self.gif()
                demonstrativo_stf()
                self.reiniciar()

            case "Financeiro - Demonstrativos TJDFT":
                self.gif()
                demonstrativo_tjdft()
                self.reiniciar()

            case "Financeiro - Demonstrativos Unafisco":
                self.gif()
                demonstrativo_unafisco()
                self.reiniciar()

            case "Glosa - Gerador de Planilha GDF":
                self.gif()
                gerar_planilha()
                self.reiniciar()

            case "Glosa - Recursar Benner(Câmara, CAMED, FAPES, Postal)":
                self.gif()
                recursar_benner()
                self.reiniciar()

            case "Glosa - Recursar BRB":
                self.gif()
                recursar_brb()
                self.reiniciar()

            case "Glosa - Recursar Cassi":
                self.gif()
                recursar_cassi()
                self.reiniciar()

            case "Glosa - Recursar GEAP Duplicado":
                self.gif()
                recursar_duplicado()
                self.reiniciar()
            
            case "Glosa - Recursar E-VIDA":
                self.gif()
                recursar_evida()
                self.reiniciar()

            case "Glosa - Recursar Fascal":
                self.gif()
                recursar_fascal()
                self.reiniciar()

            case "Glosa - Recursar GEAP Sem Duplicado":
                self.gif()
                recursar_sem_duplicado()
                self.reiniciar()

            case "Glosa - Recursar Saúde Caixa":
                self.gif()
                recursar_caixa()
                self.reiniciar()

            case "Glosa - Recursar SIS":
                self.gif()
                recursar_sis()
                self.reiniciar()
            
            case "Relatório - Brindes":
                self.gif()
                Gerar_Relat_Normal()
                self.reiniciar()

            case "Tesouraria - Nota Fiscal":
                self.gif()
                subirNF()
                self.reiniciar()
            case _:
                self.botao_iniciar()

    def voltar_inicio(self, master=None):
        self.ocultar_data()
        self.voltar.pack_forget()
        self.quintoContainer.pack_forget()
        self.sextoContainer.pack_forget()

        self.quintoContainer = Frame(master, background="white")
        self.quintoContainer["padx"] = 20
        self.quintoContainer["pady"] = 5
        self.quintoContainer.pack()

        self.sextoContainer = Frame(master, background="white")
        self.sextoContainer["padx"] = 100
        self.sextoContainer["pady"] = 10
        self.sextoContainer.pack()
        self.botao_iniciar()

    def reiniciar(self, master=None):
        self.desocultar()

        self.info.pack_forget()

        ImageLabel().unload()

        self.quintoContainer.pack_forget()

        self.sextoContainer.pack_forget()

        self.quintoContainer = Frame(master, background="white")
        self.quintoContainer["padx"] = 20
        self.quintoContainer["pady"] = 5
        self.quintoContainer.pack()

        self.sextoContainer = Frame(master, background="white")
        self.sextoContainer["padx"] = 100
        self.sextoContainer["pady"] = 10
        self.sextoContainer.pack()

        self.botao_iniciar()
        # self.botaoHistorico()
     
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
            self.delay = 80

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
def data_valida(date_string1, date_string2, date_format="%d/%m/%Y"):
    try:
        datetime.strptime(date_string1, date_format)
        datetime.strptime(date_string2, date_format)
        return True
    except ValueError:
        return False

local = dataLocal()
atualiza = dataAtualiza()

if(local == atualiza):
    print("Software atualizado")
    root = tk.Tk()
    Application(root)
    root.iconbitmap('Robo.ico')
    root.title('AMHP - Automações')
    root.geometry("530x330")
    root.configure(background="white")
    root.resizable(width=False, height=False)
    root.eval('tk::PlaceWindow . center')
    # ctypes.windll.kernel32.FreeConsole()
    root.mainloop()

else:
    try:
        tkinter.messagebox.showwarning('AMHP - Automações', 'Uma atualização será feita!')
        os.startfile(r"C:\Instaladores\INSTALA-AUTOMACAO.bat")
    except Exception as e:
        print(f"{e.__class__.__name__}: {e}")