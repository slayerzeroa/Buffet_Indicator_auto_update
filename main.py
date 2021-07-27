'''
Author : slayerzeroa
Date : 2021-07-27

Buffett indicator crawling project
'''

import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import scipy.stats as stat
import openpyxl
from datetime import date
import os as os

date = date.today() # 오늘 날짜

webpage = requests.get("https://www.currentmarketvaluation.com/models/buffett-indicator.php") #크롤링 주소
soup = BeautifulSoup(webpage.content, "html.parser") # 파서

data = soup.find_all("h4") #h4 tag 가져오기

path = "./" #폴더 위치
file_list = os.listdir(path) #폴더 내 파일 리스트 가져오기

if 'crawling.xlsx' in file_list: #만약 파일리스트에 엑셀 파일 있으면
    wb = openpyxl.load_workbook('crawling.xlsx') #불러와
    sheet = wb.active #그리고 sheet 활성화
else: #없으면
    wb = openpyxl.Workbook() #새로 만듭시다
    wb.save('crawling.xlsx') # 만들어요~
    sheet = wb.active #그리고 sheet 활성화
    sheet["A1"] = 'Date'
    sheet["B1"] = 'Aggregate US Market Value($T)'
    sheet["C1"] = 'Annualized GDP($T)'
    sheet["D1"] = 'Buffett Indicator'

n=1 # n = 1
data_list = [date] #data_list에 date 추가해봐요

for data_h4 in data: #반복!!!
    delete = data_h4.get_text() #h4 태그의 text만 가져오기
    delete = delete.translate(str.maketrans('', '', ' \n\t\r')) # text 내 공백문자 모두 삭제
    if '%' in delete:
        delete = delete.replace('=', ':')  # 분리하기
        delete = delete.replace('$', ':')
        delete = delete.replace('T', ':')
        delete = delete.replace('%', ':')
        delete = delete.split(':') # 분리하기
        delete_num = delete[-2] # 숫자(같은 문자)만 가져오기
        delete_num = float(delete_num)
        delete_num = delete_num/100
        data_list.append(delete_num) #data_list에 추가하기

    else:
        delete = delete.replace('=', ':')  # 분리하기
        delete = delete.replace('$', ':')
        delete = delete.replace('T', ':')
        delete = delete.replace('%', ':')
        delete = delete.split(':')
        delete_num = delete[-2] # 숫자(같은 문자)만 가져오기
        delete_num = float(delete_num)
        data_list.append(delete_num)


sheet.append(data_list)

wb.save('crawling.xlsx')