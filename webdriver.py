import jin as j
import os 
import time

#web
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

#ACCESS
import pyodbc
import pandas as pd


options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ['enable-automation'])





CUR_path=os.getcwd()
DB_path1 = "\\tr\dfs\ELデバイス部\書類\書類登録\アプリ\all-MDB\書類採番.mdb"

# データベースに接続します
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=\\tr\dfs\ELデバイス部\書類\書類登録\アプリ\all-MDB\書類採番.mdb;'
    )
conn = pyodbc.connect(conn_str)
  
# データベースの指定したテーブルのデータのうち、指定された条件のものを抽出します

sql = 'SELECT YOMIGANA,ID,BUILT_USER FROM T04  ORDER BY [YOMIGANA] ASC'
df = pd.read_sql(sql, conn)
#print(df)

nmp=df.values
# データベースの接続を閉じます
conn.close()


'''
print(nmp.shape[0])         #1次　配列数
print(nmp.shape[1])         #2次　配列数
print(kazu)                 #1次　配列数

print(nmp[15][0])           #読み出し
print(nmp[15][1])
print(nmp[15][2])
'''
  

print(nmp.shape[0])         #1次　配列数
print(nmp.shape[1])         #2次　配列数

name_table={}                   #'従業員番号=名前
name_table2={}                  #'名前=従業員番号

strA=""
for i in range(nmp.shape[0]):
        name_table[nmp[i][1]]=nmp[i][2] #'従業員番号=名前
        name_table2[nmp[i][2]]=nmp[i][1] #'名前=従業員番号
                
print(name_table["A0044"])
print(name_table2["神川　晃幸"])

'''
driver = webdriver.Chrome(executable_path='chromedriver.exe', chrome_options=options,service=ChromeService(ChromeDriverManager().install()))
driver.get('https://google.com')


while True:
        time.sleep(1)	
	
'''
