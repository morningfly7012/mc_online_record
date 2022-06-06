import requests
import time
import json
import os
from cProfile import run
from distutils.log import debug
from openpyxl import load_workbook
#用來匯入模組


# 讀取 Excel 檔案
wb = load_workbook('status.xlsx') #載入status.xlsx這個檔案
sheet = wb['狀態'] #選擇工作列 狀態
  
ipp = input('請輸入連線位置：')
api = requests.get("https://api.mcsrvstat.us/2/"+ipp) #設定API
apiii = json.loads(api.text) #載入API

while True:
    m = time.strftime("%M",time.localtime())
    print(m)
    if m == "59":
        now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        print(now+"紀錄中 請稍侯...")

        #更改excel檔案
        with open ("system.json",mode="r",encoding="utf-8") as filt:
            data = json.load(filt)     
            timess = str(data["time"])
            statuss = str(data["status"])
            onliness = str(data["online"])
            timein = sheet["A"+timess]
            timein.value = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            statusin = sheet["B"+statuss]
            statusin.value = apiii["online"]
            onlinein = sheet["C"+onliness]
            onlinein.value = apiii["players"]["online"]
            wheretime = data["time"] + 1 #morningfly版權所有
            wherestatus = data["status"] + 1
            whereonline = data["online"] + 1
            wb.save('status.xlsx')
        
        #更改json參數 來調等
        with open ("system.json",mode="w",encoding="utf-8") as filt:
            datas = {"time":wheretime,"status":wherestatus,"online":whereonline}
            json.dump(datas,filt)
        print(now+"紀錄完成")
        time.sleep(61)
