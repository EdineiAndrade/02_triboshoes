from sheets import save_to_google_sheets
from playwright.sync_api import sync_playwright
import pandas as pd
import time
import re
import ast

# Função para extrair dados de um produto na página de detalhes
def extract_product_data(page, url_product,nome_categiria):
    try:
        page.goto(url_product)
        time.sleep(.1)
        # Extrair informações usando os seletores fornecidos
        produto = page.locator('//*[@id="content"]/div[1]/div[2]/h1').inner_text()
        categoria = nome_categiria.title()
        print(f"Processando categoria: {categoria} | produto:{produto}")
        codigo = url_product.split("-")[-1].replace('.html','')
        codigo = int("".join(re.findall(r'\d+', codigo)))
        preco = page.locator('(//*[@class="list-unstyled"])[6]/li/h2').inner_text().replace('R$','')
        description = page.locator('div#tab-description').inner_text().replace('\n','')
        imagem = page.query_selector_all('//*[@class="thumbnails"]//a')
        lista_imagens = list(map(lambda link: link.get_attribute('href'), imagem))
        lista_imagens = ", ".join(lista_imagens)
        
        # Seleciona todas as LABELS com data-qty="0" dentro das divs de opção
        labels = page.locator('div[id^="input-option"] label[data-qty]').all()

        # Cria o dicionário {tamanho: data-qty}
        tamanhos_dict = {
        f'"{label.inner_text().strip()}"': label.get_attribute("data-qty")
            for label in labels
        }
        tamanhos = ', '.join(tamanhos_dict)
        tamanho_meio = tamanhos.split(",")[int(len(tamanhos.split(","))/2)].replace("'","").replace('"','').replace(' ','')
        df_tamanhos = pd.DataFrame(
            list(tamanhos_dict.items()),
            columns=['Valores do Atributo 1', 'Estoque']
        )
        df_tamanhos["Estoque"] = df_tamanhos["Estoque"].astype(int)  
        df_tamanhos["Valores do Atributo 1"] = df_tamanhos["Valores do Atributo 1"].str.replace('"', '', regex=False)
        return [df_tamanhos,{
            'ID': codigo,
            'Preço promocional': 0,
            "Tipo": "variable",
            "GTIN UPC EAN ISBN": "",
            'Nome': produto,
            "Publicado": 1,
            "Em Destaque": 0,
            "Visibilidade no Catálogo": "visible",
            "Descrição Curta": "",
            "Descrição": description,
            "Data de Preço Promocional Começa em": "",
            "Data de Preço Promocional Termina em": "",
            "Status do Imposto": "taxable",
            "Classe de Imposto": "parent",
            "Em Estoque": 0,
            "Estoque": "",
            "Quantidade Baixa de Estoque": 3,
            "São Permitidas Encomendas": 0,
            "Vendido Individualmente": 0,
            "Peso (kg)": 1,
            "Comprimento (cm)": 32,
            "Largura (cm)": 20,
            "Altura (cm)": 12,
            "Permitir avaliações de clientes?": 1,
            "Nota de Compra": "",
            "Preço Promocional": "",
            "Preço": preco,
            "Categorias": categoria,
            "Tags": "",
            "Classe de Entrega": "",
            "Imagens": lista_imagens,
            "Limite de Downloads": "",
            "Dias para Expirar o Download": "",
            "Ascendente": f"id:{codigo}",
            "Grupo de Produtos": "",
            "Upsells": "",
            "Venda Cruzada": "",
            "URL Externa": "",
            "Texto do Botão": "",
            "Posição": 0,
            "Brands": "",
            "Nome do Atributo 1": "Tamanho",
            "Valores do Atributo 1": tamanhos,
            "Visibilidade do Atributo 1": 0,
            "Atributo Global 1": 1,
            "Atributo Padrão 1": tamanho_meio,
            "Nome do Atributo 2": "",
            "Valores do Atributo 2": "",
            "Visibilidade do Atributo 2": 0,
            "Atributo Global 2": "",
            "Atributo Padrão 2": ""
        }]

    except Exception as e:
        print(f"Erro ao extrair dados do produto: {e}")
        return None

# Função para salvar os dados em um arquivo Excel
def save_to_excel(data, filename='products.xlsx'):    
    #data.to_excel(filename, index=False)
    save_to_google_sheets(data) 

# Função principal para realizar o scraping
def scrape_categories(base_url):
    products_data = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # headless=False para ver o navegador em ação
        page = browser.new_page()
        page.goto(base_url)
        cont = 0
        products_data = []
        # Definição do caminho do arquivo
        categorias = page.query_selector_all('//*[@class="nav navbar-nav"]/li/a')
        urls_categoria = list(map(lambda link: link.get_attribute('href'), categorias))
        for url_categoria in urls_categoria:
            categoria = url_categoria.split('/')[-1]
            nome_categiria = categoria.replace('-',' ')
            page.goto(url_categoria)
            try:
                texto = page.locator('//*[@class="col-sm-6 text-right"]').inner_text()
                numero_paginas = int(re.search(r'\((\d+)', texto).group(1)) if re.search(r'\((\d+)', texto) else 0
            except:
                continue
            for n in range(1, numero_paginas + 1):
                page.goto(f"https://www.triboshoes.com.br/{categoria}?page={n}")
                
                product_links = page.query_selector_all('//*[@class="row no-gutter"]/div/div/div/a')
                product_urls = list(map(lambda link: link.get_attribute('href'), product_links))

                for url_product in product_urls:               
                    product_data = extract_product_data(page, url_product,nome_categiria)
                    df_tamanhos = product_data[0]
                    df_produto = pd.DataFrame([product_data[1]])
                    df_produto["Valores do Atributo 1"] = df_produto["Valores do Atributo 1"].apply(
                        lambda x: ast.literal_eval(x)
                    )
                    # Explodir a coluna "Valores do Atributo 1"   
                    df_explodido = df_produto.explode("Valores do Atributo 1").reset_index(drop=True)
                    df_explodido = df_explodido.drop(columns=["Estoque"])
                    df_explodido["Tipo"] = "variation"  
                    df_explodido[
                    ["Descrição","Peso (kg)","Comprimento (cm)","Largura (cm)",
                     "Altura (cm)","Categorias","Visibilidade do atributo 1","Atributo Padrão 1",
                     "Quantidade Baixa de Estoque","Imagens"]
                                 ] = None
                    df_produto[
                        ["Classe de Imposto","Ascendente","Preço"]
                               ] = None
                    df_explodido["Permitir avaliações de clientes?"] = 0
                    df_explodido["Posição"] = df_explodido.index + 1
                    df_explodido["ID"] = df_explodido["ID"].astype(str) + df_explodido["Posição"].astype(str)
                    df_final = df_explodido.merge(
                        df_tamanhos[["Valores do Atributo 1", "Estoque"]],  # Selecionar colunas desejadas
                        on="Valores do Atributo 1",
                        how="left"
                    )
                    df_final["Em Estoque"] = df_final["Estoque"]   
                    df_final = pd.concat([df_produto, df_final], ignore_index=True)
                    products_data.append(df_final)
                    df_final = pd.concat(products_data, ignore_index=True)
                    df_final = df_final.fillna("")
                    cont = cont + 1
                    if cont >= 1:
                        time.sleep(1)
                        save_to_excel(df_final, 'products.xlsx')
                        time.sleep(1)
                        cont = 0
        browser.close()

        return df_final

# Executar o scraping e salvar os dados
if __name__ == "__main__":
    base_url = 'https://triboshoes.com.br/'  
    data = scrape_categories(base_url)
    save_to_google_sheets(data)