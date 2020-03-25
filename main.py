import pandas as pd
import numpy as np
import datetime
import time
import os
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_GSPAnalysis")

def timeCheck(rowA,rowB):
    return rowB['datetime']-rowA['datetime'] <= delta

print("Read and Preprocess File\n")

df = pd.read_excel("input.xlsx",sheet_name=2,usecols=[6,7,8])
df1 = df.dropna(subset=["Equipment Code"]).reset_index(drop=True).copy()

for i in range(len(df1)):
    D, M, Y = df1['Application Date'][i].split("/")
    h, m = df1['Application time'][i].split(":")
    df1.at[i,"datetime"] = datetime.datetime(int(Y),int(M),int(D),int(h),int(m))
    #df1.at[i,"Equipment Code"] = [df1.loc[i,"Equipment Code"]]
    #df1.at[i,"time_string"] = [str(Y)+str(M)+str(D)+str(h)+str(m)]

    #print("YMDhm",str(Y)+str(M)+str(D)+str(h)+str(m))

print("Done\n\n")

time_base_data = list()
delta = datetime.timedelta(minutes=10)
df2 = df1.copy()
Range = len(df2)

print("Unique the Data\n")

equipmentList = pd.unique(df2['Equipment Code']).tolist()

print("Done\n\n")

resultMatrix = np.array([[0]*len(equipmentList)]*len(equipmentList),dtype=int)


print("Main Process...\n\n")
start_time = time.time()


resultdf = pd.DataFrame(columns=equipmentList)

for indexA , rowA in df2.iterrows():
    equipmentNum = [0]*len(equipmentList)
    print("A Round",indexA,"  Start")
    for indexB , rowB in df2.iterrows():
        print(indexA,"B Round",indexB)
        if(indexB<=indexA):
            print("B",indexB,"pass")
            continue
        elif( timeCheck(rowA,rowB) ):
            equipmentNum[equipmentList.index(rowB['Equipment Code'])] += 1
        else:
            break
    resultdf[rowA['Equipment Code']] = equipmentNum
    print("A Round",indexA,"  End")



for i in range(Range):
    for j in range(Range):

        if(df2['datetime'][j]-df2['datetime'][i] > delta):
            break
        elif(df2['datetime'][j]-df2['datetime'][i] <= delta):
            resultMatrix[equipmentList.index(df2.at[i,'Equipment Code'])][equipmentList.index(df2.at[j,'Equipment Code'])] += 1
        elif(df2['datetime'][j] == df2['datetime'][i]):
            resultMatrix[equipmentList.index(df2.at[i,'Equipment Code'])][equipmentList.index(df2.at[j,'Equipment Code'])] = np.nan




print("Process Time : ",time.time()-start_time,"sec\n")
print("Done\n\n")

print("Output\n")

result = pd.DataFrame(data=resultMatrix,columns=equipmentList,index=equipmentList)

result.to_excel('result.xlsx',encoding='utf_8_sig')

print("Done\n\n")