import requests
from requests_html import HTMLSession
from bs4 import BeautifulSoup
import xlwings as xl
from tqdm import tqdm

listProducts = []

for i in tqdm(range(0, 32)):
  
  url = f"https://www.soriana.com/soriana/es/c/Vinos-y-Licores/c/G?q=%3Arelevance&page={i}"

  session = HTMLSession()

  htmlContainer = session.get(url).html

  html = BeautifulSoup(htmlContainer.html, "html.parser")

  contenedor = html.select('div.product-item div.details')

  for c in contenedor:
    product = f'{c.select_one("div.product-productNameCont").text}'.replace('\n','').strip()
    price = f'{c.select_one("div.priceContainer").text}'.replace('\n','').strip()
    listProducts.append(tuple([product, price]))

wb = xl.Book()
ws = wb.sheets[0]

ws.range((1,1)).value = ["Product","Price"]
ws.range((2,1)).value = listProducts