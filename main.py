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

data = {"SKU":[],
      "ASIN":[],
      "ASP":[],
      "Current Lowest Price":[]}

for sku in SKUASINPairs:

    driver = webdriver.Chrome('C:\\Users\\ryanj\..PROGRAMS\Python\Acurite-Amazon-Scraper\chromedriver')

    url = "https://camelcamelcamel.com/product/" + SKUASINPairs[sku]

    print("ASIN:", SKUASINPairs[sku])

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

            data["SKU"].append(sku)
            data["ASIN"].append(SKUASINPairs[sku])
            try:
                data["ASP"].append(SKUASPPairs[sku])
            except:
                data["ASP"].append("not found")
            data["Current Lowest Price"].append(priceTag.text)

            driver.close()

    if (input() == "1"):
        break

df = pd.DataFrame(data)

#df["Current Lowest Price"][1] = 2

try:
    writer = pd.ExcelWriter("aspData.xlsx", engine="xlsxwriter")

    df.to_excel(writer, sheet_name="Sheet1")
    wb = writer.book
    ws = writer.sheets["Sheet1"]

    lowCostFormat = wb.add_format()
    lowCostFormat.set_bg_color("red")
    ws.conditional_format("D2:E{}".format(len(df) + 1), {"type": "formula", "criteria": "=$D2>$E2", "format": lowCostFormat})

    writer.close()
except:
    raise Exception("Please close the aspData.xlsx file and run again.")

print(df)