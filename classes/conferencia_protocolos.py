import os
import PyPDF2

class ConferenciaProtocolos():
    lista_convenios_guias_digitalizadas = ['(433)', '(457)', '(381)', '(319)', '(225)', '(056)', '(023)']

    def remessa_em_arquivo_unico(self, nome_arquivo, diretorio, processo):
        for root, dirs, files in os.walk(diretorio):
            arquivo = self.name_is_in_list(nome_arquivo, files)
            if arquivo:
                path_arquivo = os.path.join(root, arquivo)
                if self.ler_pdfs(path_arquivo, processo):
                    return True
        return False
    
    def remessa_is_diretorio(self, diretorio, remessa, lista_processos):
        for root, dirs, files in os.walk(diretorio):
            print(root)
            if remessa in root.split('/')[-1]:
                for processo in lista_processos:
                    if self.name_is_in_list():
                        ...
    
    def name_is_in_list(self, nome, str_list):
        for name in str_list:
            if nome in name:
                return name
            
        return None
    
    def ler_pdfs(self, arquivo: str, processo) -> bool:
        with open(arquivo, 'rb') as arquivo_pdf:
            leitor_pdf: PyPDF2.PdfReader = PyPDF2.PdfReader(arquivo_pdf)

            if self.is_aceite(leitor_pdf.pages):
                return False

            for pagina_numero in range(len(leitor_pdf.pages)):
                pagina = leitor_pdf.pages[pagina_numero]
                print(' '.join(pagina.extract_text().split('\n')))
                texto_sem_quebra = ' '.join(pagina.extract_text().split('\n'))
                if 'Encaminhamos guias de atendimentos, referentes aos serviços prestados' in texto_sem_quebra:
                    continue
                if processo in texto_sem_quebra:
                    return True
            
            return False
        
    def is_aceite(self, arquivo):
        for pagina_numero in range(len(arquivo)):
            pagina = arquivo[pagina_numero]
            texto = ' '.join(pagina.extract_text().split('\n'))
            if 'Relatório de procedimentos realizados (Carta)' not in texto:
                return False
        return True
    
    def _(self):
        dir = ...
        n_remessa = ...
        nome_convenio = ...
        processo = ...

        if self.remessa_em_arquivo_unico(n_remessa, dir, processo):
            ...
        
        elif self.remessa_is_diretorio():
            ...

ConferenciaProtocolos().remessa_is_diretorio(r'\\10.0.0.239\financeiro - faturamento\Protocolos de Convênios - Aceite\(20) CONAB\2024', '45271')