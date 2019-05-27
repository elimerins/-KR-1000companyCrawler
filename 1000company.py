from selenium import webdriver
import openpyxl
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup
import chardet

import requests
wb = openpyxl.load_workbook('1000companydataset.xlsx',data_only=True)
ws = wb.active
# HTTP GET Request
for i in range(100,101):
    req = requests.get('http://m.mk.co.kr/yearbook/index.php?page='+str(i)+'&TM=Y2&MM=T0')

    #참조 : http://pythonstudy.xyz/python/article/403-%ED%8C%8C%EC%9D%B4%EC%8D%AC-Web-Scraping
    req.encoding=None

    # HTML 소스 가져오기
    html = req.text

    soup = BeautifulSoup(html, 'html.parser')
    tds = soup.find("tbody")
    companies=tds.find_all('a')
    ranking=tds.find_all('td',{'class':'center'})
    company=[]
    ranks=[]

    for a in companies:
        b=a.text
        b=b.replace('(주)','')
        company.append(b)

    for rank in ranking:
        ranks.append(rank.text)

    for j in range(10):
        #print(company[j])
        ws.cell(row=i*10+j+2, column=1).value=company[j]
        ws.cell(row=i*10+j+2, column=2).value=int(ranks[j])
    wb.save('1000companydataset.xlsx')
    print('page '+str(i)+' done.')