import urllib.request
from bs4 import BeautifulSoup
import pandas as pd
import requests
import json

class AppURLopener(urllib.request.FancyURLopener):
    version = "Mozilla/5.0"

opener = AppURLopener()

df = pd.read_excel("teste.xlsx")
dframe = df.to_dict('list')

fiis = dframe["FII"]

valor_atual = {}
lista_valores = []
index=2

for fii in fiis:
    x = requests.get(f'https://brapi.dev/api/quote/{fii}?fundamental=false')
    y = f'https://www.fundsexplorer.com.br/funds/{fii}'
    resposta = json.loads(x.text)
    preco = resposta["results"][0]['regularMarketPrice']
    lista_valores.append(preco)
    page = opener.open(y)
    soup = BeautifulSoup(page, 'html5lib')
    list_item = soup.findAll('span', attrs={'class': 'indicator-value'})
    valor_dividendo = float(list_item[1].text.strip().replace("R$ ",'').replace(",",'.'))

    df.loc[df['FII'] == fii,"Valor Atual"] = [preco]
    df.loc[df['FII'] == fii,"Resultado"] = [f"=((E{index}/C{index})-1)*100"]
    df.loc[df['FII'] == fii,"Financeiro"] = [f"=((E{index}-C{index})*B{index})"]
    df.loc[df['FII'] == fii,"Último Rendimento"] = [valor_dividendo]
    df.loc[df['FII'] == fii,"Total Rendimento"] = [f"=H{index}*B{index}"]
    index+=1

df.to_excel("teste.xlsx", index=False)
