

def generateSKUASPPairs():
    skus = pd.read_excel("products_table.xlsx", usecols="A").transpose()

    asps = pd.read_excel("products_table.xlsx", usecols="B").transpose()

    pairs = {}

    for i in range(len(skus.columns)):
        pairs[skus.iloc[0][i]] = asps.iloc[0][i]

    return pairs


