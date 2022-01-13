import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def generateASPASINPairs():
    skus = pd.read_excel("ASIN_ASP_Pairs.xlsx", usecols="A").transpose()
    asins = pd.read_excel("ASIN_ASP_Pairs.xlsx", usecols="B").transpose()
    asps = pd.read_excel("ASIN_ASP_Pairs.xlsx", usecols="C").transpose()

    pairs = {}

    for i in range(len(skus.columns)):
        pairs[skus.iloc[0][i]] = {"ASIN": asins.iloc[0][i], "ASP": asps.iloc[0][i], "Current Lowest Price": -1}

    return pairs

ASPASINPairs = generateASPASINPairs()

for sku in ASPASINPairs:

    driver = webdriver.Chrome('C:\\Users\\ryanj\..PROGRAMS\Python\Acurite-Amazon-Scraper\chromedriver')

    url = "https://camelcamelcamel.com/product/" + ASPASINPairs[sku]["ASIN"]

    driver.get(url)
    priceTag = ""
    try:
        priceTag = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.XPATH, "//span[@class=\"stat\"]"))
        )
    except:
        pass
    finally:
        if priceTag == "":
            driver.quit()
        else:
            ASPASINPairs[sku]["Current Lowest Price"] = priceTag.text[1:]
            driver.quit()

df = pd.DataFrame(ASPASINPairs).transpose()

df.to_excel("aspData.xlsx")

try:
    writer = pd.ExcelWriter("aspData.xlsx", engine="xlsxwriter")

    df.to_excel(writer, sheet_name="Sheet1")
    wb = writer.book
    ws = writer.sheets["Sheet1"]

    lowCostFormat = wb.add_format()
    lowCostFormat.set_bg_color("#cc0000")
    ws.conditional_format("C2:D{}".format(len(df) + 1), {"type": "formula", "criteria": "=$C2>$D2", "format": lowCostFormat})

    writer.close()
except:
    raise Exception("Please close the aspData.xlsx file and run again.")
