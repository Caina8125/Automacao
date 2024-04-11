import os

class ConferenciaProtocolos():
    lista_convenios_guias_digitalizadas = ['(433)', '(457)', '(381)', '(319)', '(225)', '(056)', '(023)']

    def procurar_arquivo(nome_arquivo, diretorio):
    # Percorre todos os arquivos e subdiretórios no diretório
        for root, _, files in os.walk(diretorio):
            # Verifica se o arquivo está na lista de arquivos
            if nome_arquivo in files:
                return os.path.join(root, nome_arquivo)  # Retorna o caminho completo do arquivo
        return None  # Retorna None se o arquivo não for encontrado