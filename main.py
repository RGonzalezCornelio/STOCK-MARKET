#MAIN

import openpyxl
import requests
from bs4 import BeautifulSoup
import re
import json
import signal
from colorama import init, Fore


#Le meteremos un handler seguramente porque son mas de 5000 lineas
def handler(signum, frame):
    with open("./counter.txt", 'w') as file:
        file.write(str(x))
    work.save("./STOCK MARKET NAMES.xlsx")
    print(Fore.RED + "❌")
    exit(1)


work = openpyxl.load_workbook("./STOCK MARKET NAMES.xlsx")
companies = work['NAMES']
column = 1 #Columns with the company names
y = int(open('./counter.txt').readline())
linkstarter = "https://finance.yahoo.com/quote/"
headers = {"User-Agent":"Mozilla/5.0"}

for x in range(y, 5):

    print(Fore.RESET + "processing "+ str(x) + "... ", end= '')
    jsonf = {} #Guardaremos la informacion en un json
    cell = companies.cell(row=x, column=column)
    company_name = str(companies.cell(row=x, column=1).value)
    jsonf["Stock Name"] = company_name

    link = linkstarter + company_name + "/key-statistics/"
    jsonf["link"] = link
    #print(company_name + ", " + link)

    try:
        respuesta = requests.get(link, headers=headers)
        if str(respuesta) != "<Response [200]>":
            print(Fore.RED + "❌")
            continue
    except Exception:
        handler("", "")

    #respuesta = requests.get(link, headers=headers)
    #Empezamos a coger datos de la pagina web
    data = BeautifulSoup(respuesta.text, 'html.parser')

    company = str(data.select('title')[0]).split(">")[1].split('(')[0]
    jsonf["Company Name"] = company
    print(company)

    stock_value = str(data.find('div', class_="container yf-1tejb6")).split('>')[3].split('<')[0]
    jsonf["Stock Value"] = stock_value
    print("Stock Value: " + stock_value)

    #statictics_table = str(data.find('table', class_="table yf-kbx2lo"))
    #statictics_table_values = str(data.find_all('td', class_="yf-kbx2lo"))
    #statictics_table_values = statictics_table.find('table', class_="table yf-kbx2lo")
    #print(statictics_table_values)

    market_cap = str(data.find_all('td', class_="yf-kbx2lo")[1]).split('>')[1].split('<')[0]
    enterprise_value = str(data.find_all('td', class_="yf-kbx2lo")[8]).split('>')[1].split('<')[0]
    trailing_PE = str(data.find_all('td', class_="yf-kbx2lo")[15]).split('>')[1].split('<')[0]
    forward_PE = str(data.find_all('td', class_="yf-kbx2lo")[22]).split('>')[1].split('<')[0]
    enterprise_value_EBITDA = str(data.find_all('td', class_="yf-kbx2lo")[57]).split('>')[1].split('<')[0]
    print("market_cap: " + market_cap + "Enterprise value: " + enterprise_value + "Trailing P/E: " + trailing_PE + "forward_PE: " + forward_PE + "enterprise_value_EBITDA: " + enterprise_value_EBITDA)

    jsonf["Market Cap"] = market_cap
    jsonf["Enterprise Value"] = enterprise_value
    jsonf["Trailing P/E"] = trailing_PE
    jsonf["Forward P/E"] = forward_PE
    jsonf["Enterprise Value/EBITDA"] = enterprise_value_EBITDA

    jsonfile = open('./DataJSON/' + str(x) + '.json', 'w')
    json.dump(jsonf, jsonfile)
    print(Fore.GREEN + "✓")

work.save("./STOCK MARKET NAMES.xlsx")