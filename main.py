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

init()
signal.signal(signal.SIGINT, handler)

work = openpyxl.load_workbook("./STOCK MARKET NAMES.xlsx")
companies = work['NAMES']
column = 1 #Columns with the company names
y = int(open('./counter.txt').readline())
linkstarter = "https://finance.yahoo.com/quote/"
headers = {"User-Agent":"Mozilla/5.0"}

for x in range(y, 10):

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
    #print("market_cap: " + market_cap + "Enterprise value: " + enterprise_value + "Trailing P/E: " + trailing_PE + "forward_PE: " + forward_PE + "enterprise_value_EBITDA: " + enterprise_value_EBITDA)

    jsonf["Market Cap"] = market_cap
    jsonf["Enterprise Value"] = enterprise_value
    jsonf["Trailing P/E"] = trailing_PE
    jsonf["Forward P/E"] = forward_PE
    jsonf["Enterprise Value/EBITDA"] = enterprise_value_EBITDA


    profit_margin = str(data.find_all('td', class_="value yf-vaowmx")[2]).split('>')[1].split('<')[0]
    operating_margin = str(data.find_all('td', class_="value yf-vaowmx")[3]).split('>')[1].split('<')[0]
    quarterly_revenue_growth = str(data.find_all('td', class_="value yf-vaowmx")[8]).split('>')[1].split('<')[0]
    quarterly_earnings_growth = str(data.find_all('td', class_="value yf-vaowmx")[13]).split('>')[1].split('<')[0]
    beta = str(data.find_all('td', class_="value yf-vaowmx")[22]).split('>')[1].split('<')[0]
    percentage_by_insiders = str(data.find_all('td', class_="value yf-vaowmx")[34]).split('>')[1].split('<')[0]
    forward_annual_divident_rate = str(data.find_all('td', class_="value yf-vaowmx")[41]).split('>')[1].split('<')[0]
    five_year_divident_rate = str(data.find_all('td', class_="value yf-vaowmx")[45]).split('>')[1].split('<')[0]
    #print("profit_margin: " + profit_margin + " operating_margin: " + operating_margin + " quarterly_revenue_growth: " + quarterly_revenue_growth + " quarterly_earnings_growth: " + quarterly_earnings_growth\
    #       + " beta: " + beta + " percentage_by_insiders: " + percentage_by_insiders + " forward_annual_divident_rate" + forward_annual_divident_rate + " five_year_divident_rate: " + five_year_divident_rate)

    jsonf["Profit Margin"] = profit_margin
    jsonf["Operating Margin"] = operating_margin
    jsonf["Quarterly Revenue Growth"] = quarterly_revenue_growth
    jsonf["Quarterly Earnings Growth"] = quarterly_earnings_growth
    jsonf["Beta (5Y Monthly)"] = beta
    jsonf["Percentage by insiders"] = percentage_by_insiders
    jsonf["Forward Annual Dividend Rate"] = forward_annual_divident_rate
    jsonf["5 Year Average Dividend Yield"] = five_year_divident_rate
    

    jsonfile = open('./DataJSON/' + str(x) + '.json', 'w')
    json.dump(jsonf, jsonfile)
    print(Fore.GREEN + "✓")

work.save("./STOCK MARKET NAMES.xlsx")