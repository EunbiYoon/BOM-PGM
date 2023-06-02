import pandas as pd
import openpyxl 

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/NPTGERP/0602/result_0602.xlsx", sheet_name="RV13D1AMAZU.ABWEUUS_worst item")
read_excel.index=["",1,2,3,4,5,6,7]
read_excel.columns=["NPT","","","","","","","NPT vs GERP","","GERP","","","","","",""]
read_excel=read_excel.round(1)
print(read_excel)

#column_list
column_list=list(read_excel.columns)
column_list.insert(0,"index")

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','""')
print("")
print(column_json+",")

# #row_list - index
# bb=pd.DataFrame()
# for i in range(len(read_excel.index)):
#     bb.at[0,i]='"index":"'+str(read_excel.at[i,0])+'",'

print('"rows":[')
#row_list - values
for i in range(len(read_excel.index)):
    print(read_excel.iloc[i])
    # if i==len(read_excel.index)-1:
    #     aa=str(list(read_excel.iloc[i+1][1:])).replace("'",'"').replace('nan','""')
    #     aA='"values":'+aa
    #     AA="{"+str(bb.at[0,i])+aA+"}"
    #     print(AA)
    # else:
    #     aa=str(list(read_excel.iloc[i][1:])).replace("'",'"').replace('nan','""')
    #     aA='"values":'+aa
    #     AA="{"+str(bb.at[0,i])+aA+"},"
    #     print(AA)
print(']')







