#導入所需模組
import requests
import time
import pandas as pd
from datetime import datetime, timedelta

#宣告
co=[0,0,0,0,0,1,1,0]
res_0=[0,0,0,0,0,0,0,0]
entry=0

#每三秒執行
while True:
#爬蟲
    all=requests.get("https://api.thingspeak.com/channels/1510233/feeds.json?api_key=N27RYM4ET5AE8DT2&results=2").json()

#最近兩次時間
    res_1=[time for time in all["feeds"][0]["created_at"]]
    res_2=[time for time in all["feeds"][1]["created_at"]]
    Day1=int(res_1[8]+res_1[9])
    Day2=int(res_2[8]+res_2[9])
    Hour1=int(res_1[11]+res_1[12])
    Hour2=int(res_2[11]+res_2[12])
    Min1=int(res_1[14]+res_1[15])
    Min2=int(res_2[14]+res_2[15])

#如果有新數據
    if entry!=all["feeds"][1]["entry_id"] and entry!=0:

    #如果那兩次數據為同一分鐘內的數據
        if res_1[14]+res_1[15]==res_2[14]+res_2[15]:
        #計算數據總和及次數
            for a in range(1,9):
            #風向與UV不能參加運算
                if a==6 or a==7:
                    res_0[a-1]=int(all["feeds"][1]["field"+str(a)])
                else:
                    res_0[a-1]+=float(all["feeds"][1]["field"+str(a)])
                    co[a-1]+=1
        
    #如果那兩次數據為不同分鐘的數據
        else:
        #如果從未入過數據
            if res_0==[0,0,0,0,0,0,0,0]:
            #入一次舊數據
                for a in range(1,9):
                    if a==6 or a==7:
                        res_0[a]=int(all["feeds"][0]["field"+str(a)])
                    else:
                        res_0[a-1]+=float(all["feeds"][0]["field"+str(a)])
                        co[a-1]+=1

            
        #計算所有數據一分鐘的平均
            for c in range(8):
            #5,6不能進行運算,否則輸入不到正確數據
                co[5]=1
                co[6]=1
                res_0[c]/=co[c]
            
        #調整數據
            res_0[0]=int(res_0[0]*100)/100
            res_0[1]=int(res_0[1]/693.68*1.0083703*100)/100
            res_0[2]=int(res_0[2])
            res_0[3]=int(res_0[3]*45.415*100)/100
            res_0[4]=int(res_0[4]*3.257*100)/100
            res_0[7]=int(res_0[7])
            if res_0[5]==45:
                res_0[5]="東北"
            elif res_0[5]==90:
                res_0[5]="東"
            elif res_0[5]==135:
                res_0[5]="東南"
            elif res_0[5]==180:
                res_0[5]="南"
            elif res_0[5]==225:
                res_0[5]="西南"
            elif res_0[5]==270:
                res_0[5]="西"
            elif res_0[5]==315:
                res_0[5]="西北"
            elif res_0[5]==360:
                res_0[5]="北"
            
        #當前要輸入時間
            if Hour1!=Hour2:
                if Day1!=Day2:
                    timeloop=60-Min1+Min2+60*(24-Hour1+Hour2-1)
                else:
                    timeloop=60-Min1+Min2+60*(Hour2-Hour1-1)
            else:
                timeloop=Min2-Min1

        #時間超1分鐘的會生成報告(偵錯用)
            if Min2-Min1>1:
                with open("bug_report.txt",mode="a",encoding="utf-8") as file2:
                    file2.write("時間:"+time.strftime("%Y/%m/%d %H時")+str(Min1)+"分\n")

        #輸出txt,excel檔
            for i in range(0,timeloop):
                date=datetime.now()+timedelta(minutes=-timeloop+i)
                with open("data.txt",mode="a",encoding="utf-8") as file1:
                    file1.write(str(entry)+"\n時間:"+date.strftime("%Y/%m/%d %H時%M分")+"\n温度°C:"+str(res_0[0])+"\n大氣壓hPa:"+str(res_0[1])+"\n濕度%:"+str(res_0[2])+"\n降兩量mm:"+str(res_0[3])+"\n平均風速km/h:"+str(res_0[4])+"\n風向:"+str(res_0[5])+"\n紫外線指數:"+str(res_0[6])+"\nPM2.5:"+str(res_0[7])+"\n\n")

                file=pd.read_excel("target.xlsx")
                data=pd.DataFrame({
                    "時間:":date.strftime("%Y/%m/%d %H時%M分"),
                    "温度°C:":res_0[0],
                    "大氣壓hPa:":res_0[1],
                    "濕度%:":res_0[2],
                    "降兩量mm:":res_0[3],
                    "平均風速km/h:":res_0[4],
                    "風向:":res_0[5],
                    "紫外線指數:":res_0[6],
                    "PM2.5:":res_0[7],
                    },index=[0])
                new=file.append(data,ignore_index=True)
                new.to_excel("target.xlsx",index=False)

                file2=pd.read_excel("real.xlsx")
                data1=pd.DataFrame({
                    "時間:":date.strftime("%Y/%m/%d %H時%M分"),
                    "温度°C:":res_0[0],
                    "大氣壓hPa:":res_0[1],
                    "濕度%:":res_0[2],
                    "降兩量mm:":res_0[3],
                    "平均風速km/h:":res_0[4],
                    "風向:":res_0[5],
                    "紫外線指數:":res_0[6],
                    "PM2.5:":res_0[7],
                    },index=[0])
                new1=file2.append(data1,ignore_index=True)
                new1.to_excel("real.xlsx",index=False)
                
        #輸出google excel


        #重設數據
            for a in range(1,9):
                if a==6 or a==7:
                    res_0[a-1]=int(all["feeds"][1]["field"+str(a)])
                else:
                    res_0[a-1]=float(all["feeds"][1]["field"+str(a)])
                    co[a-1]=1
        
    #人為偵錯用
        print(str(entry)+"  "+res_1[14]+res_1[15]+"    "+res_2[14]+res_2[15])
        print(res_0)
        print(co)
   
#重新輸入數據編號
    entry=all["feeds"][1]["entry_id"]
    
#三秒
    time.sleep(3)