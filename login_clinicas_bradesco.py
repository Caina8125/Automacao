from pandas import read_excel

def get_login_data() -> list[dict]:
    try:
        logins_list = []
        path = r"\\10.0.0.239\automacao_integralis\ANEXOS BRADESCO\Logins e Senhas.xlsx"
        df_planilha_login = read_excel(path).sort_values(by=['Clinica'])

        for index, linha in df_planilha_login.iterrows():
            logins_list.append(
                {
                    'nome': linha['Clinica'],
                    'user': f'{linha["Login"]}'.replace('.0', ''),
                    'password': f'{linha["Senha"]}'.replace('.0', '')
                }
            )
        
        return logins_list
    
    except:
        pass