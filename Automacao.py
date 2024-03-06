from tkinter import *
import tkinter as tk 
from tkinter import ttk
import threading
from PIL import Image, ImageTk
from itertools import count, cycle
from filtro_matricula import filtrar_matricula
from user_authentication import UserLogin
from Recursar_Duplicado import *
from Buscar_fatura import iniciar
from Benner.enviar_pdf_benner import enviar_pdf_benner
from Benner.enviar_xml_benner import enviar_xml_benner
from Atualiza_Local import *
from Anexar_Guia_Geap import anexar_guias
from Enviar_Xml_Caixa import Enviar_caixa
from VerificarSituacao_BRB import verificacao_brb
from recursar_brb import recursar_brb
from recursar_cassi import recursar_cassi
from recursar_fascal import recursar_fascal
from recursar_evida import recursar_evida
from Recursar_SemDuplicado import recursar_sem_duplicado
from Recursar_Caixa import recursar_caixa
from Recurso_Benner import recursar_benner
from Recursar_SIS import recursar_sis
from Recursar_Stf import recursar_stf
from recursar_real_grandeza import recursar_real
from recursar_stm import recursar_stm
from recursar_tjdft import recursar_tjdft
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
from Nota_Fiscal_2 import subirNF2
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

        self.comboBox = ttk.Combobox(self.quartoContainer, values=["Faturamento - Anexar Guia Geap",
                                                                   "Faturamento - Conferência GEAP",
                                                                   "Faturamento - Conferência Bacen",
                                                                   "Faturamento - Enviar PDF Bacen",
                                                                   "Faturamento - Enviar PDF Benner",
                                                                   "Faturamento - Enviar PDF BRB",
                                                                   "Faturamento - Enviar XML Bacen",
                                                                   "Faturamento - Enviar XML Benner",
                                                                   "Faturamento - Enviar XML Caixa",
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
                                                                   "Glosa - Filtro Matrículas",
                                                                   "Glosa - Recursar Benner(Câmara, CAMED, FAPES, Postal)",
                                                                   "Glosa - Recursar BRB",
                                                                   "Glosa - Recursar Cassi",
                                                                   "Glosa - Recursar E-VIDA",
                                                                   "Glosa - Recursar Fascal",
                                                                   "Glosa - Recursar GEAP Duplicado",
                                                                   "Glosa - Recursar GEAP Sem Duplicado",
                                                                   "Glosa - Recursar Real Grandeza",
                                                                   "Glosa - Recursar Saúde Caixa",
                                                                   "Glosa - Recursar SIS",
                                                                   "Glosa - Recursar STF",
                                                                   "Glosa - Recursar STM",
                                                                   "Glosa - Recursar TJDFT",
                                                                   "Relatório - Brindes",
                                                                   "Tesouraria - Nota Fiscal"
                                                                    ], width=50)
        self.comboBox["background"] = 'white'
        self.comboBox.pack(side=LEFT)

        self.botao_iniciar()

    def gif(self):
        
        self.info = Label(self.quintoContainer, text= "Trabalhando...",font=('Arial,10,bold'), background="white")
        self.info.pack()

        self.lbl = ImageLabel(self.quintoContainer,background="white")
        self.lbl.pack(side=LEFT)
        self.lbl.load('loader2.gif')

    def ocultar(self):
        self.buttonIniciar.grid_forget()
        self.texto.pack_forget()

    def ocultar_data(self):
        try:
            self.inserir_data_inicial.pack_forget()
            self.inserir_data_final.pack_forget()
            self.botao_ok.pack_forget()
            self.voltar.pack_forget()
        except:
            pass

    def ocultar_login(self):
        try:
            self.label_user.grid_forget()
            self.insert_user.grid_forget()
            self.label_password.grid_forget()
            self.insert_password.grid_forget()
            self.botao_ok.pack_forget()
            self.voltar.pack_forget()
        except:
            pass

    def desocultar(self):
        self.lbl.pack_forget()    

    def botao_iniciar(self):
        self.buttonIniciar = Button(self.quintoContainer, bg="#274360",foreground="white",width=10, command=lambda: threading.Thread(target=self.chamarAutomacao).start())
        self.buttonIniciar["text"] = "Iniciar"
        self.buttonIniciar.grid(row=1, column=0, padx=10, pady=20)
        self.texto = Label(self.sextoContainer, 
                           text="As automações no navegador necessitam de\nautenticação no proxy. Caso deseje utilizar\numa delas, digite seu usuário e senha da rede.",
                           font=self.fontePadrao, background="white")
        self.texto.pack()

    def obter_datas(self):           
        global data_inicial, data_final, validacao
        data_inicial = self.inserir_data_inicial.get()
        data_final = self.inserir_data_final.get()
        validacao = data_valida(data_inicial, data_final)
        self.ocultar_data()

        if validacao:
            self.inserir_login(demonstrativo_cassi)

        else:
            tkinter.messagebox.showerror( 'Data inválida!' , 'Digíte uma data válida')
            self.inserir_data()

    def run_funtions(self, funcao1, user, password):
        if funcao1.__name__ == 'demonstrativo_cassi':
            funcao1(data_inicial, data_final, user, password)
        else:
            funcao1(user, password)
        self.reiniciar()

    # def run_funcion_cassi()

    def exec_automacao(self, funcao):
        user = self.insert_user.get()
        password = self.insert_password.get()
        user_login = UserLogin(user, password)
        self.ocultar_login()
        self.gif()
        threading.Thread(target=lambda: self.run_funtions(funcao, user_login.user, user_login.password)).start()

    def inserir_data(self):
        self.inserir_data_inicial = tk.Entry(self.quintoContainer)
        self.inserir_data_inicial.insert(0, "Digite a data inicial")
        

        self.inserir_data_final = tk.Entry(self.quintoContainer)
        self.inserir_data_final.insert(0, "Digite a data final")

        self.botao_ok = Button(self.sextoContainer, bg="#274360",foreground="white", text="OK", command=lambda: threading.Thread(target=self.obter_datas).start())

        self.voltar = Button(self.sextoContainer, bg="#274360", foreground="white", text="Voltar", command=lambda: threading.Thread(target=self.voltar_inicio(self.ocultar_data)).start())

        self.inserir_data_inicial.pack(side=LEFT)
        self.inserir_data_inicial.bind("<FocusIn>", self.limpar_placeholder_data1)
        self.inserir_data_final.pack(side=LEFT)
        self.inserir_data_final.bind("<FocusIn>", self.limpar_placeholder_data2)
        self.botao_ok.pack(side=LEFT, padx=10)
        self.voltar.pack(side=RIGHT)

    def inserir_login(self, event):
        self.label_user = tk.Label(self.quintoContainer, text="Usuário:")
        self.label_user.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        self.insert_user = tk.Entry(self.quintoContainer)
        self.insert_user.grid(row=0, column=1, padx=10, pady=10)
        
        self.label_password = tk.Label(self.quintoContainer, text="Senha:")
        self.label_password.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        self.insert_password = tk.Entry(self.quintoContainer, show='*')
        self.insert_password.grid(row=1, column=1, padx=10, pady=10)

        self.botao_ok = Button(self.sextoContainer, bg="#274360",foreground="white", text="OK", command=lambda: threading.Thread(target=self.exec_automacao(event)).start())

        self.voltar = Button(self.sextoContainer, bg="#274360", foreground="white", text="Voltar", command=lambda: threading.Thread(target=self.voltar_inicio(self.ocultar_login)).start())

        # self.insert_user.pack()
        # self.insert_user.bind("<FocusIn>", self.limpar_placeholder_user)
        # self.insert_password.pack()
        # self.insert_password.bind("<FocusIn>", self.limpar_placeholder_password)
        self.botao_ok.pack(side=LEFT, padx=10)
        self.voltar.pack(side=RIGHT)

    def limpar_placeholder_data1(self, event):
        if self.inserir_data_inicial.get() == "Digite a data inicial":
            self.inserir_data_inicial.delete(0, "end")
            self.inserir_data_inicial.config(fg="black")

    def limpar_placeholder_data2(self, event):
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
            case "Faturamento - Anexar Guia Geap":
                self.inserir_login(anexar_guias)

            case "Faturamento - Conferência GEAP":
                self.inserir_login(conferencia)

            case "Faturamento - Conferência Bacen":
                self.inserir_login(conferir_bacen)

            case "Faturamento - Enviar PDF Bacen":
                self.reiniciar()

            case "Faturamento - Enviar PDF Benner":
                self.inserir_login(enviar_pdf_benner)

            case "Faturamento - Enviar PDF BRB":
                self.inserir_login(enviar_pdf)

            case "Faturamento - Enviar XML Bacen":
                self.inserir_login(fazer_envio_xml)

            case "Faturamento - Enviar XML Benner":
                self.inserir_login(enviar_xml_benner)

            case "Faturamento - Enviar XML Caixa":
                self.inserir_login(Enviar_caixa)

            case "Faturamento - Leitor de PDF GAMA":
                self.gif()
                pdf_reader()
                self.reiniciar()

            case "Faturamento - Verificar Situação BRB":
                self.inserir_login(verificacao_brb)

            case "Faturamento - Verificar Situação Fascal":
                self.inserir_login(verificacao_fascal)

            case "Faturamento - Verificar Situação Gama":
                self.inserir_login(verificar_gama)

            case "Financeiro - Buscar Faturas GEAP":
                self.inserir_login(iniciar)

            case "Financeiro - Demonstrativos Amil":
                self.inserir_login(demonstrativo_amil)

            case "Financeiro - Demonstrativos BRB":
                self.inserir_login(demonstrativo_brb)

            case "Financeiro - Demonstrativos Câmara dos Deputados":
                self.inserir_login(demonstrativo_camara)

            case "Financeiro - Demonstrativos Camed":
                self.inserir_login(demonstrativo_camed)

            case "Financeiro - Demonstrativos Casembrapa":
                self.inserir_login(demonstrativo_casembrapa)

            case "Financeiro - Demonstrativos Cassi":
                self.inserir_data()

            case "Financeiro - Demonstrativos Codevasf":
                self.inserir_login(demonstrativo_codevasf)

            case "Financeiro - Demonstrativos E-Vida":
                self.inserir_login(demonstrativo_evida)

            case "Financeiro - Demonstrativos Fapes":
                self.inserir_login(demonstrativo_fapes)

            case "Financeiro - Demonstrativos Fascal":
                self.inserir_login(demonstrativo_fascal)

            case "Financeiro - Demonstrativos Gama":
                self.inserir_login(demonstrativo_gama)

            case "Financeiro - Demonstrativos Life Empresarial":
                self.inserir_login(demonstrativo_life)

            case "Financeiro - Demonstrativos MPU":
                self.inserir_login(demonstrativo_mpu)

            case "Financeiro - Demonstrativos PMDF":
                self.inserir_login(demonstrativo_pmdf)

            case "Financeiro - Demonstrativos Postal":
                self.inserir_login(demonstrativo_postal)

            case "Financeiro - Demonstrativos Real Grandeza":
                self.inserir_login(demonstrativo_real)

            case "Financeiro - Demonstrativos Saúde Caixa":
                self.inserir_login(demonstrativo_caixa)

            case "Financeiro - Demonstrativos Serpro":
                self.inserir_login(demonstrativo_serpro)

            case "Financeiro - Demonstrativos SIS":
                self.inserir_login(demonstrativo_sis)

            case "Financeiro - Demonstrativos STF":
                self.inserir_login(demonstrativo_stf)

            case "Financeiro - Demonstrativos TJDFT":
                self.inserir_login(demonstrativo_tjdft)

            case "Financeiro - Demonstrativos Unafisco":
                self.inserir_login(demonstrativo_unafisco)

            case "Glosa - Gerador de Planilha GDF":
                self.gif()
                gerar_planilha()
                self.reiniciar()

            case "Glosa - Filtro Matrículas":
                self.gif()
                filtrar_matricula()
                self.reiniciar()

            case "Glosa - Recursar Benner(Câmara, CAMED, FAPES, Postal)":
                self.inserir_login(recursar_benner)

            case "Glosa - Recursar BRB":
                self.inserir_login(recursar_brb)

            case "Glosa - Recursar Cassi":
                self.inserir_login(recursar_cassi)

            case "Glosa - Recursar GEAP Duplicado":
                self.inserir_login(recursar_duplicado)

            case "Glosa - Recursar GEAP Sem Duplicado":
                self.inserir_login(recursar_sem_duplicado)
            
            case "Glosa - Recursar E-VIDA":
                self.inserir_login(recursar_evida)

            case "Glosa - Recursar Fascal":
                self.inserir_login(recursar_fascal)

            case "Glosa - Recursar Real Grandeza":
                self.inserir_login(recursar_real)

            case "Glosa - Recursar Saúde Caixa":
                self.inserir_login(recursar_caixa)

            case "Glosa - Recursar SIS":
                self.inserir_login(recursar_sis)

            case "Glosa - Recursar STF":
                self.inserir_login(recursar_stf)

            case "Glosa - Recursar STM":
                self.inserir_login(recursar_stm)

            case "Glosa - Recursar TJDFT":
                self.inserir_login(recursar_tjdft)

            case "Relatório - Brindes":
                self.gif()
                Gerar_Relat_Normal()
                self.reiniciar()

            case "Tesouraria - Nota Fiscal":
                self.inserir_login(subirNF2)
            case _:
                self.botao_iniciar()

    def voltar_inicio(self, funcao, master=None):
        funcao()
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
    root.geometry("500x400")
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