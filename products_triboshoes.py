from sheets import save_to_google_sheets
from playwright.sync_api import sync_playwright
import pandas as pd
import time
import locale
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
        codigo = url_product.split("-")[-1].replace('.html','')
        preco = page.locator('(//*[@class="list-unstyled"])[6]/li/h2').inner_text().replace('R$','')
        description = page.locator('div#tab-description').inner_text().replace('\n','')

        imagem = page.query_selector_all('//*[@class="thumbnails"]/li/a/img')
        lista_imagens = list(map(lambda link: link.get_attribute('src'), imagem))
        lista_imagens = ",".join(lista_imagens)
        
        # Seleciona todas as LABELS com data-qty="0" dentro das divs de opção
        labels = page.locator('div[id^="input-option"] label[data-qty]').all()

        # Cria o dicionário {tamanho: data-qty}
        tamanhos_dict = {
        f'"{label.inner_text().strip()}"': label.get_attribute("data-qty")
            for label in labels
        }
        tamanhos = ', '.join(tamanhos_dict)
        df_tamanho = pd.DataFrame(
        list(tamanhos_dict.items()),
        columns=['Valores do Atributo 1', 'Estoque']
    )
        
        return [df_tamanho,{
            'Categoria': categoria,
            'ID': codigo,
            'SKU': "",
            'EAN': "",
            'NCM': "",
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
            "Classe de Imposto": "",
            "Em Estoque": 1,
            "Estoque": 30,
            "Quantidade Baixa de Estoque": 3,
            "São Permitidas Encomendas": 0,
            "Vendido Individualmente": 0,
            "Peso (kg)": 1,
            "Comprimento (cm)": 32,
            "Largura (cm)": 20,
            "Altura (cm)": 12,
            "Permitir Avaliações de Clientes": 1,
            "Nota de Compra": "",
            "Preço Promocional": "",
            "Preço": preco,
            "Categorias": "",
            "Tags": "",
            "Classe de Entrega": "",
            "Imagens": lista_imagens,
            "Limite de Downloads": "",
            "Dias para Expirar o Download": "",
            "Ascendente": "",
            "Grupo de Produtos": "",
            "Upsells": "",
            "Venda Cruzada": "",
            "URL Externa": "",
            "Texto do Botão": "",
            "Posição": "",
            "Brands": "",
            "Nome do Atributo 1": "Tamanho",
            "Valores do Atributo 1": tamanhos,
            "Visibilidade do Atributo 1": 0,
            "Atributo Global 1": 1,
            "Atributo Padrão 1": "",
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
            texto = page.locator('//*[@class="col-sm-6 text-right"]').inner_text()
            numero_paginas = int(re.search(r'\((\d+)', texto).group(1))

            for n in range(1, numero_paginas + 1):
                page.goto(f"https://www.triboshoes.com.br/{categoria}?page={n}")
                
                product_links = page.query_selector_all('//*[@class="row no-gutter"]/div/div/div/a')
                product_urls = list(map(lambda link: link.get_attribute('href'), product_links))

                for url_product in product_urls:                
                    print(f"Processando categoria: {url_product}")
                    
                    product_data = extract_product_data(page, url_product,nome_categiria)
                    df_tamanhos = product_data[0]
                    df_tamanhos["Valores do Atributo 1"] = df_tamanhos["Valores do Atributo 1"].str.replace('"', '', regex=False)
                    
                    df_produto = pd.DataFrame([product_data[1]])
                    df_produto["Valores do Atributo 1"] = df_produto["Valores do Atributo 1"].apply(
                        lambda x: ast.literal_eval(x)
                    )
                    # Explodir a coluna "Valores do Atributo 1"   
                    df_explodido = df_produto.explode("Valores do Atributo 1").reset_index(drop=True)
                    df_explodido = df_explodido.drop(columns=["Estoque"])
                    df_explodido["Tipo"] = "variation"
                    df_tamanhos["Estoque"] = df_tamanhos["Estoque"].astype(int)    
                    df_final = df_explodido.merge(
                        df_tamanhos[["Valores do Atributo 1", "Estoque"]],  # Selecionar colunas desejadas
                        on="Valores do Atributo 1",
                        how="left"
                    )   
                    df_final = pd.concat([df_produto, df_final], ignore_index=True)

                    products_data.append(df_final)
                    df_final = pd.concat(products_data, ignore_index=True)
                    cont = cont + 1
                    if cont >= 20:
                        time.sleep(1)
                        save_to_excel(df_final, 'products.xlsx')
                        time.sleep(1)
                        cont = 0
        browser.close()

        return df_final

# Executar o scraping e salvar os dados
if __name__ == "__main__":
    base_url = 'https://triboshoes.com.br/'  # Substitua pela URL base do e-commerce
    data = scrape_categories(base_url)
    save_to_google_sheets(data)