import requests
from bs4 import BeautifulSoup
import pandas as pd 
from openpyxl import load_workbook
from io import StringIO




def getWebpage():
    url = "https://finviz.com/screener.ashx?v=121&f=cap_large,fa_div_o5&ft=4"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
    reponse = requests.get(url=url, headers=headers)
    soup  = BeautifulSoup(reponse.content, 'html.parser')
    
    table = soup.find('table', class_='styled-table-new is-rounded is-tabular-nums w-full screener_table')
    
    print(table)
