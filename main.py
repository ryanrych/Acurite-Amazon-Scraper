import pandas as pd
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

    driver = webdriver.Chrome('C:\\Users\\ryanj\..PROGRAMS\Python\Acurite-Amazon-Scraper\chromedriver')

    url = "https://camelcamelcamel.com/product/" + SKUASINPairs[pair]

    print("ASIN:", SKUASINPairs[pair])

    driver.get(url)
    # sleep(5)
    # priceTag = driver.find_element_by_xpath("//span[@class=\"stat\"]")
    priceTag = ""
    try:
        priceTag = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.XPATH, "//span[@class=\"stat\"]"))
        )
    except:
        pass
    finally:
        if priceTag == "":
            driver.close()
        else:
            print(priceTag.text)
            driver.close()
