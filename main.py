import pandas as pd
import numpy as np
import datetime
import time
import os
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_GSPAnalysis")

def timeCheck(timeA,timeB):
    return timeB-timeA <= delta

print("Read and Preprocess File\n")

df = pd.read_excel("input.xlsx",sheet_name=2,usecols=[6,7,8])
df1 = df.dropna(subset=["Equipment Code"]).reset_index(drop=True).copy()

for i in range(len(df1)):
    # 將日期分項取出
    D, M, Y = df1['Application Date'][i].split("/")
    h, m = df1['Application time'][i].split(":")
    # 將取出的日期轉成datatime格式保存進新的欄位
    df1.at[i,"datetime"] = datetime.datetime(int(Y),int(M),int(D),int(h),int(m))

print("Done\n\n")

#判斷時間差
delta = datetime.timedelta(minutes=10)

#複製一份資料
df2 = df1.copy()
# df2 = df1.head(500).copy()

print("Unique the Data\n")
equipmentList = pd.unique(df2['Equipment Code']).tolist()

print("Done\n\n")

# 新增一個全都為零的二維陣列，兩維長度皆為Equipment Code的種類
resultMatrix = np.array([[0]*len(equipmentList)]*len(equipmentList),dtype=int)

print("Main Process...\n\n")

#計時
start_time = time.time()
#判斷用
equipmentSet = set()
#存放結果的DataFrame
resultdf = pd.DataFrame(columns=equipmentList)

for indexA , rowA in df2.iterrows():
    equipmentNum = [0]*len(equipmentList)
    print("A Round : ",indexA,"  Start")
    for indexB , rowB in df2.iloc[indexA+1:].iterrows():
        if( timeCheck(rowA['datetime'],rowB['datetime']) ):
            equipmentNum[equipmentList.index(rowB['Equipment Code'])] += 1
        else:
            break

    if(rowA['Equipment Code'] in equipmentSet):
        resultdf[rowA['Equipment Code']] += equipmentNum
    else:
        resultdf[rowA['Equipment Code']] = equipmentNum
        equipmentSet.add(rowA['Equipment Code'])

    print("A Round : ",indexA,"  End")


print("Process Time : ",time.time()-start_time,"sec\n\n")

print("Output...\n")

resultdf.index = equipmentList

with pd.ExcelWriter('resultdf.xlsx') as writer:

    df2.to_excel(writer,sheet_name='Data',encoding='utf_8_sig')
    resultdf.to_excel(writer,sheet_name='Result',encoding='utf_8_sig')
    

print("Done\n\n")

exit()