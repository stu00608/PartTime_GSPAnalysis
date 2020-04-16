import pandas as pd
import numpy as np
import datetime
import time
import os
import progressbar as pb
os.chdir(r"C:\Users\user\Desktop\PartTime_GSPAnalysis") 

def timeCheck(timeA,timeB):
    #此函數用作計算時間差，以datetime格式相減，最後與delta時間比較
    return timeB-timeA <= delta

print("Read and Preprocess File\n")

df = pd.read_excel("input.xlsx",sheet_name=2,usecols=[6,7,8]) #讀input.xlsx內第三個工作表(sheet_name)的第7,8,9行(usecols)
df1 = df.dropna(subset=["Equipment Code"]).reset_index(drop=True).copy() #將讀到的表格中在Equipment Code欄位為空的一整列刪除

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

print("Unique the Data\n")
equipmentList = pd.unique(df2['Equipment Code']).tolist() #將Equipment Code種類列出並轉為list

print("Done\n\n")

# 新增一個全都為零的二維陣列，兩維長度皆為Equipment Code的種類
resultMatrix = np.array([[0]*len(equipmentList)]*len(equipmentList),dtype=int)

print("Main Process...\n\n")

# 計算每個設備壞掉的次數，列出Pi 按Pi大小排序

broken_count = df2['Equipment Code'].value_counts().reset_index() #計算df2的Equipment Code欄位內每個Equipment Code出現的次數(value_counts)，隨後重置索引

total_break = int(broken_count['Equipment Code'].sum()) #計算出現次數的總和即為全部壞掉的次數

broken_count = broken_count.set_index(['index']).astype(int) #將作為欄位的Equipment Code改成index，並將數值調整成int形式
broken_probability = broken_count / total_break #計算Pi：i設備壞掉的機率= Ni(i設備壞掉的次數)/Sum(Ni)i=1 to n(所有設備壞掉的次數)
broken_probability.columns = ['Pi'] #更改欄位名稱
broken_probability_output = broken_probability.copy().sort_values(ascending=False,by=['Pi']) #將計算完成的資料由大到小排序
#以下四行為將broken_count照著equipmentList內的順序擺放
new_broken_count = pd.DataFrame(index=equipmentList,columns=['count']) 
for index,row in new_broken_count.iterrows():
    row['count'] = broken_count.loc[index,'Equipment Code']
broken_count = new_broken_count.copy()


#判斷用
equipmentSet = set()
#存放結果的DataFrame
resultdf = pd.DataFrame(columns=equipmentList)

print("計算在A設備故障時其他設備故障的Matrix\n\n")

#progressbar初始化 (這是可以在terminal當中看到進度條的插件)
total = len(df2)
t = 0
pbar = pb.ProgressBar(widgets = ['Progress: ',pb.Percentage(), ' ', pb.Bar('>'),' ', pb.Timer(),
      ' ', pb.ETA(), ' ', pb.FileTransferSpeed()]).start()

#df2第一層迴圈(index是從0開始到最後，row是現在迴圈的那一橫行，可以用中括號欄位名稱去取特定值)
for indexA , rowA in df2.iterrows():

    #progressbar更新
    pbar.update(int((t / (total - 1)) * 100))
    t += 1

    equipmentNum = [0]*len(equipmentList) #產生一個全部Equipment Code種類長度的0陣列
    #df2第二層迴圈，使用從第一層迴圈的位置後開始到最後
    for indexB , rowB in df2.iloc[indexA+1:].iterrows():
        #判斷時間差
        if( timeCheck(rowA['datetime'],rowB['datetime']) ):
            equipmentNum[equipmentList.index(rowB['Equipment Code'])] += 1 #若時間差小於delta，則記錄一次
        else:
            break #因為時間是有排序過的，所以一旦沒有在時間差內，這之後的資料都不會在時間差內，故可以直接break

    #判斷此設備是否有檢測過，沒有就直接取代原本欄位，有就加上原本欄位
    if(rowA['Equipment Code'] in equipmentSet):
        resultdf[rowA['Equipment Code']] += equipmentNum
    else:
        resultdf[rowA['Equipment Code']] = equipmentNum
        equipmentSet.add(rowA['Equipment Code'])

#progressbar結束
pbar.finish()
print("\n\n")


# 計算Pij (矩陣) 
print("計算在A設備故障時其他設備故障的關聯機率\n\n")

resultdf.index = equipmentList #將resultdf(上面計算的結果)的索引從數字改成Equipment Code

Pij = resultdf.copy() #copy一份資料給Pij
col = list(Pij.columns) #取Pij的欄位轉為list
count = list(broken_count['count']) #取broken_count的數字欄位轉為list

# Pij = Nij(resultdf) / Ni(broken_count)
for item in col:
    Pij[item] /= count
    
print("\n\n")

#計算Ci 重要指標

Ii = 1
Ij = 1

sum_of_Pij =  Pij.sum(axis=1) #計算Pij的總和(橫向)
sum_of_Pij = pd.DataFrame(data=sum_of_Pij) #新增一個DataFrame存進總和
sum_of_Pij = sum_of_Pij.sort_index() #照索引排序資料
broken_probability = broken_probability.sort_index() #照索引排序資料
Ci_index = list(broken_probability.index) #取index出來轉為list 

#重要指標計算
Ci = np.array(broken_probability['Pi']) * Ii + np.array(broken_probability['Pi']) * Ii * np.array(sum_of_Pij[0]) * Ij
#將最後結果組成DataFrame存放到Ci
Ci = pd.DataFrame(columns = ['Ci'], index = Ci_index , data = Ci).sort_values(ascending=False,by=['Ci'])

print("Output...\n")


#輸出Excel
with pd.ExcelWriter('resultdf.xlsx') as writer:

    df2.to_excel(writer,sheet_name='Data',encoding='utf_8_sig')
    resultdf.to_excel(writer,sheet_name='Result',encoding='utf_8_sig')
    broken_probability.to_excel(writer,sheet_name='Pi',encoding='utf_8_sig')
    Pij.to_excel(writer,sheet_name='Pij',encoding='utf_8_sig')
    Ci.to_excel(writer,sheet_name='Ci',encoding='utf_8_sig')
    

print("Done\n\n")

exit()