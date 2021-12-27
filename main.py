import pandas as pd
import requests
from bs4 import BeautifulSoup

def generateSKUASPPairs():
    skus = pd.read_excel("products_table.xlsx", usecols="A").transpose()

    asps = pd.read_excel("products_table.xlsx", usecols="B").transpose()

    pairs = {}

    for i in range(len(skus.columns)):
        pairs[skus.iloc[0][i]] = asps.iloc[0][i]

    return pairs

def generateSKUASINPairs():
    skus = pd.read_excel("asin_skus.xlsx", usecols="B").transpose()

    asins = pd.read_excel("asin_skus.xlsx", usecols="A").transpose()

    pairs = {}

    for i in range(len(skus.columns)):
        pairs[skus.iloc[0][i]] = asins.iloc[0][i]

    return pairs


SKUASPPairs = generateSKUASPPairs()
SKUASINPairs = generateSKUASINPairs()

for pair in SKUASINPairs:

    url = "https://camelcamelcamel.com/product/" + SKUASINPairs[pair]

    r = requests.get(url)
    soup = BeautifulSoup(r.content)
