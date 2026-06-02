import time
import pandas as pd
from io import StringIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 1. Configurações do Navegador
options = Options()
options.add_argument('--start-maximized')

print("Iniciando o navegador...")
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=options)

url = "https://microdados.evolucional.com.br/microdados/ranking-enem-2024?redacao=sim"
navegador.get(url)

# Dá um tempo inicial para o site carregar completamente
print("Carregando o site...")
time.sleep(5)

todos_dados = []
paginas_para_raspar = 455

for pagina in range(1, paginas_para_raspar + 1):
    print(f"Extraindo página {pagina} de {paginas_para_raspar}...")
    
    try:
        # 2. Espera até a tabela aparecer na tela
        wait = WebDriverWait(navegador, 10)
        tabela = wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        
        # 3. Lê o HTML da tabela e converte em uma planilha temporária (DataFrame)
        html_tabela = tabela.get_attribute('outerHTML')
        df_pagina = pd.read_html(StringIO(html_tabela), decimal=',', thousands='.')[0]
        todos_dados.append(df_pagina)
        
        # 4. Vai para a próxima página (se não estiver na última)
        if pagina < paginas_para_raspar:
            # Procura o botão "Próximo" e clica nele
            botao_proximo = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Próximo')]")))
            # Usando script JS para clicar e evitar problemas com botões sobrepostos
            navegador.execute_script("arguments[0].click();", botao_proximo)
            
            # Pausa para dar tempo do site buscar os novos dados
            time.sleep(2)
            
    except Exception as e:
        print(f"Erro na página {pagina}. O raspador vai parar por aqui. Detalhe: {e}")
        break

# 5. Fecha o navegador
navegador.quit()

# 6. Junta todas as 455 tabelinhas em uma tabela só
if todos_dados:
    print("Consolidando os dados...")
    df_final = pd.concat(todos_dados, ignore_index=True)

    # Limpa colunas vazias ou de layout que o Pandas possa ter pego por engano
    df_final = df_final.dropna(how='all', axis=1)

    # --- LIMPEZA DO NOME DA ESCOLA ---
    # Substitui os últimos 8 números (código INEP) por nada ('') e tira espaços extras
    if 'Escola' in df_final.columns:
        df_final['Escola'] = df_final['Escola'].astype(str).str.replace(r'\d{8}$', '', regex=True).str.strip()
    
    # --- SEPARAR CIDADE E ESTADO ---
    if 'Cidade' in df_final.columns:
        estados_br = [
            "Acre", "Alagoas", "Amapá", "Amazonas", "Bahia", "Ceará", "Distrito Federal", 
            "Espirito Santo", "Goiás", "Maranhão", "Mato Grosso do Sul", "Mato Grosso", 
            "Minas Gerais", "Pará", "Paraíba", "Paraná", "Pernambuco", "Piauí", 
            "Rio de Janeiro", "Rio Grande do Norte", "Rio Grande do Sul", "Rondônia", 
            "Roraima", "Santa Catarina", "São Paulo", "Sergipe", "Tocantins"
        ]
        
        # Monta um padrão de busca com todos os estados
        regex_estados = f"(.*?)({'|'.join(estados_br)})$"
        
        # Extrai a cidade e o estado
        extraido = df_final['Cidade'].astype(str).str.extract(regex_estados)
        
        # Cria a coluna Estado
        df_final['Estado'] = extraido[1]
        
        # Atualiza a coluna Cidade (se algo falhar na extração, ele mantém o texto original por segurança)
        df_final['Cidade'] = extraido[0].fillna(df_final['Cidade']).str.strip()
        
        # Reorganiza a ordem das colunas para "Estado" ficar logo depois de "Cidade"
        cols = df_final.columns.tolist()
        if 'Estado' in cols:
            cols.remove('Estado')
            pos_cidade = cols.index('Cidade')
            cols.insert(pos_cidade + 1, 'Estado')
            df_final = df_final[cols]
    
    # --- SEPARAR DEPENDÊNCIA E LOCALIZAÇÃO ---
    # A coluna original que o Pandas pegou chama-se "Dependência Administrativa"
    if 'Dependência Administrativa' in df_final.columns:
        # Extrai "Urbana" ou "Rural" para uma nova coluna 'Localização'
        df_final['Localização'] = df_final['Dependência Administrativa'].str.extract(r'(Urbana|Rural)')
        
        # Remove a palavra "Urbana" ou "Rural" da coluna original, deixando só a Dependência
        df_final['Dependência Administrativa'] = df_final['Dependência Administrativa'].str.replace(r'Urbana|Rural', '', regex=True).str.strip()
        
        # Reorganiza a ordem das colunas para "Localização" ficar logo depois
        cols = df_final.columns.tolist()
        if 'Localização' in cols:
            cols.remove('Localização')
            pos_dep = cols.index('Dependência Administrativa')
            cols.insert(pos_dep + 1, 'Localização')
            df_final = df_final[cols]
    # --------------------------------------------------------

    # 7. Salva no Excel
    nome_arquivo = "Ranking_ENEM_2024_Completo.xlsx"
    df_final.to_excel(nome_arquivo, index=False, engine='openpyxl')
    print(f"\n Sucesso! Planilha gerada com {len(df_final)} escolas.")
    print(f"Arquivo salvo como: {nome_arquivo}")
else:
    print("Nenhum dado foi extraído. Verifique sua conexão com a internet ou se o site mudou.")
