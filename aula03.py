"""
Automação Web e busca de informações com Python

Desafio

Trabalhamos em uma importadora e o preço dos nossos produtos é vinculado a cotação de:
- Dólar
- Euro
- Ouro

Precisamos pegar na internet, de forma automática, a cotação desses 3 itens e saber quanto
devemos cobrar pelos nossos produtos, considerando uma margem de contribuição que temos na nossa
base de dados.

Base de Dados: https://drive.google.com/drive/folders/1KmAdo593nD8J9QBaZxPOG1yxHZua4Rtv

Para isso, vamos criar uma automação web:
- Usaremos o selenium
- Importante baixar o webdriver: chromedriver
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd

# Criar o navegador
navegador = webdriver.Chrome("chromedriver.exe")

# Entrar no Google
navegador.get('https://www.google.com/')

# Pegar a cotação do dólar
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    'cotação dólar')

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    Keys.ENTER)

cotacao_dolar = navegador.find_element(By.XPATH,
                                       '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute(
    'data-value')

print(cotacao_dolar)

# Pesquisar e pegar a cotação do euro
navegador.get('https://www.google.com/')

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    'cotação euro')

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    Keys.ENTER)

cotacao_euro = navegador.find_element(By.XPATH,
                                      '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute(
    'data-value')

print(cotacao_euro)

# Pesquisar e Pegar a cotação do ouro do site: https://www.melhorcambio.com/ouro-hoje
navegador.get('https://www.melhorcambio.com/ouro-hoje')

cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(',', '.')

print(cotacao_ouro)

# Fechar o navegador
navegador.quit()

# Importar a base de dados, para atualizar os dados com os valores atuais de cotação, coletados acima
tabela = pd.read_excel('bd/Produtos.xlsx')

# Atualizar as cotações na base de dados

# Cotação do dólar
# tabela.loc[linha, coluna]
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(cotacao_dolar)

# Cotação do euro
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(cotacao_euro)

# Cotação do ouro
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(cotacao_ouro)

# Atualizar o preço de compra e de venda na base de dados
# preço de compra = preço original * cotação
tabela['Preço de Compra'] = tabela["Preço Original"] * tabela['Cotação']

# preço de venda = preço de compra * margem
tabela['Preço de Venda'] = tabela["Preço de Compra"] * tabela['Margem']

print(tabela)

# Exportar a base de dados para atualiza-la
tabela.to_excel('bd/ProdutosNovo.xlsx', index=False)
