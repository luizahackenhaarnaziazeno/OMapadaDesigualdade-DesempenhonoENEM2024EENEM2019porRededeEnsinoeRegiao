
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

ANOS = [
    {
        "ano": 2024,
        "url": "https://microdados.evolucional.com.br/microdados/ranking-enem-2024?redacao=sim",
        "arquivo": "Ranking_ENEM_2024_Completo.xlsx",
    },
    {
        "ano": 2019,
        "url": "https://microdados.evolucional.com.br/microdados/edicoes-anteriores/2019?redacao=sim",
        "arquivo": "Ranking_ENEM_2019_Completo.xlsx",
    },
]

ESTADOS_BR = [
    "Acre", "Alagoas", "Amapá", "Amazonas", "Bahia", "Ceará", "Distrito Federal",
    "Espirito Santo", "Goiás", "Maranhão", "Mato Grosso do Sul", "Mato Grosso",
    "Minas Gerais", "Pará", "Paraíba", "Paraná", "Pernambuco", "Piauí",
    "Rio de Janeiro", "Rio Grande do Norte", "Rio Grande do Sul", "Rondônia",
    "Roraima", "Santa Catarina", "São Paulo", "Sergipe", "Tocantins",
]


def detectar_total_paginas(navegador):
    try:
        itens = navegador.find_elements(By.XPATH, "//ul[contains(@class,'pagination')]//a")
        numeros = [int(i.text.strip()) for i in itens if i.text.strip().isdigit()]
        return max(numeros) if numeros else None
    except Exception:
        return None


def limpar_dados(df):
    df = df.dropna(how="all", axis=1)

    if "Escola" in df.columns:
        df["Escola"] = df["Escola"].astype(str).str.replace(r"\d{8}$", "", regex=True).str.strip()

    if "Cidade" in df.columns:
        regex_estados = f"(.*?)({'|'.join(ESTADOS_BR)})$"
        extraido = df["Cidade"].astype(str).str.extract(regex_estados)
        df["Estado"] = extraido[1]
        df["Cidade"] = extraido[0].fillna(df["Cidade"]).str.strip()
        cols = df.columns.tolist()
        cols.remove("Estado")
        cols.insert(cols.index("Cidade") + 1, "Estado")
        df = df[cols]

    if "Dependência Administrativa" in df.columns:
        df["Localização"] = df["Dependência Administrativa"].str.extract(r"(Urbana|Rural)")
        df["Dependência Administrativa"] = (
            df["Dependência Administrativa"].str.replace(r"Urbana|Rural", "", regex=True).str.strip()
        )
        cols = df.columns.tolist()
        cols.remove("Localização")
        cols.insert(cols.index("Dependência Administrativa") + 1, "Localização")
        df = df[cols]

    return df


def raspar_ano(config):
    ano = config["ano"]
    print(f"\n{'='*50}")
    print(f"Iniciando raspagem do ENEM {ano}...")
    print(f"{'='*50}")

    options = Options()
    options.add_argument("--start-maximized")
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico, options=options)

    navegador.get(config["url"])
    print("Carregando o site...")
    time.sleep(5)

    wait = WebDriverWait(navegador, 15)
    try:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        time.sleep(2)
        total_paginas = detectar_total_paginas(navegador)
        if total_paginas:
            print(f"Total de páginas detectado: {total_paginas}")
        else:
            total_paginas = 455
            print(f"Não foi possível detectar páginas. Usando {total_paginas} como padrão.")
    except Exception as e:
        total_paginas = 455
        print(f"Erro ao detectar páginas: {e}. Usando {total_paginas}.")

    todos_dados = []

    for pagina in range(1, total_paginas + 1):
        print(f"  Extraindo página {pagina}/{total_paginas}...", end="\r")
        try:
            wait = WebDriverWait(navegador, 10)
            tabela = wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            html_tabela = tabela.get_attribute("outerHTML")
            df_pagina = pd.read_html(StringIO(html_tabela), decimal=",", thousands=".")[0]
            todos_dados.append(df_pagina)

            if pagina < total_paginas:
                botao = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Próximo')]"))
                )
                navegador.execute_script("arguments[0].click();", botao)
                time.sleep(2)

        except Exception as e:
            print(f"\nErro na página {pagina}. Parando aqui. Detalhe: {e}")
            break

    navegador.quit()

    if todos_dados:
        print(f"\nConsolidando dados do ENEM {ano}...")
        df_final = pd.concat(todos_dados, ignore_index=True)
        df_final = limpar_dados(df_final)
        df_final.to_excel(config["arquivo"], index=False, engine="openpyxl")
        print(f"✓ ENEM {ano}: {len(df_final)} escolas salvas em '{config['arquivo']}'")
        return True
    else:
        print(f"✗ ENEM {ano}: Nenhum dado extraído.")
        return False


# Roda para os dois anos sequencialmente
for config in ANOS:
    raspar_ano(config)

print("\nConcluído! Ambos os arquivos foram gerados.")
