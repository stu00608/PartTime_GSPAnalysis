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
# df2 = df1.head(500).copy()

Range = len(df2)

print("Unique the Data\n")

equipmentList = pd.unique(df2['Equipment Code']).tolist()

print("Done\n\n")

resultMatrix = np.array([[0]*len(equipmentList)]*len(equipmentList),dtype=int)


print("Main Process...\n\n")
start_time = time.time()
equipmentSet = set()
resultdf = pd.DataFrame(columns=equipmentList)

for indexA , rowA in df2.iterrows():
    equipmentNum = [0]*len(equipmentList)
    print("A Round : ",indexA,"  Start")
    for indexB , rowB in df2.iloc[indexA+1:].iterrows():
        # print(indexA,"B Round",indexB)
        if( timeCheck(rowA['datetime'],rowB['datetime']) ):
            # print('rowB[\'Equipment Code\'] : ',rowB['Equipment Code'])
            # print("equipmentList.index(rowB['Equipment Code']) : ",equipmentList.index(rowB['Equipment Code']))
            equipmentNum[equipmentList.index(rowB['Equipment Code'])] += 1
        else:
            break
    # print('rowA[\'Equipment Code\'] : ',rowA['Equipment Code'])

    if(rowA['Equipment Code'] in equipmentSet):
        resultdf[rowA['Equipment Code']] += equipmentNum
    else:
        resultdf[rowA['Equipment Code']] = equipmentNum
        equipmentSet.add(rowA['Equipment Code'])

    print("A Round : ",indexA,"  End")

print("main loop done \n")

print("Check...\n")

if(len(equipmentList) != len(equipmentSet)):
    print("Error!! len(equipmentList) != len(equipmentSet)")
    

print("Process Time : ",time.time()-start_time,"sec\n")
print("Done\n\n")

print("Output\n")

resultdf.index = equipmentList

with pd.ExcelWriter('resultdf.xlsx') as writer:

    df2.to_excel(writer,sheet_name='Data',encoding='utf_8_sig')
    resultdf.to_excel(writer,sheet_name='Result',encoding='utf_8_sig')
    

print("Done\n\n")

exit()