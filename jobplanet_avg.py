import xlrd
from selenium import webdriver
import openpyxl
import requests
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup
wb = openpyxl.load_workbook('1000companydataset.xlsx',data_only=True)
ws = wb.active
url='https://www.jobplanet.co.kr/search?utf8=&query='
for idx,i in enumerate(ws['A2':'A1011']):
    print(idx,i[0].value,end=' ')

    req = requests.get(url+i[0].value)

    html = req.text

    soup = BeautifulSoup(html, 'html.parser')
    #print(soup)
    try:
        result_card = soup.find('span',{'class':'rate_ty02'})
        print(result_card.text)
        ws.cell(row=i[0].row, column = 13).value = result_card.text
        if idx%10 == 0:
            wb.save('./1000companydataset.xlsx')
            print('saved.')

    except Exception as e:
        print('avg not found')
    time.sleep(2)

wb.save('1000companydataset.xlsx')





