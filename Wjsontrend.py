import pandas as pd
import openpyxl 

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/NPTGERP/0428/result_0428.xlsx", sheet_name="RV13D1AMAZU.ABWEUUS_result")
read_excel.index=read_excel['Unnamed: 0']
read_excel=read_excel.drop(['Unnamed: 0'],axis=1)
extract_data=read_excel.T
extract_data["PO + Substitute"]=extract_data["PO Price Change"]+extract_data["Substitute Change"]
extract_data=extract_data.round(1)

#column_list
column_list=list(extract_data.index)
column_list.insert(0,"index")

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','""')
print("")
print(column_json+",")

#row_list
extract_data=extract_data.T
extract_data.reset_index(inplace=True)
extract_data=extract_data.T
extract_data.reset_index(inplace=True, drop=True)
extract_data=extract_data.T

#row_list - index
bb=pd.DataFrame()
for i in range(len(extract_data.index)):
    bb.at[0,i]='"index":"'+str(extract_data.at[i,0])+'",'

print('"rows":[')
#row_list - values
for i in range(len(extract_data.index)-1):
    if i==len(extract_data.index)-2:
        aa=str(list(extract_data.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"}"
        print(AA)
    else:
        aa=str(list(extract_data.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"},"
        print(AA)
print(']')







