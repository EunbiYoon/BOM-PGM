import pandas as pd
import numpy as np

###########################GERP###############################
gerp=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/gerp.xlsx')
# #파일에 따라 추가하기
# gerp=gerp[gerp['Model'].str.contains('F3')] #total 행 제거
# gerp.reset_index(drop=True, inplace=True)
for i in range(len(gerp)):
    level=gerp.at[i,'Level']
    if level==0:
        gerp.at[i,'level0']=gerp.at[i,'Child Item']
    elif level=='*1':
        gerp.at[i,'level0']=gerp.at[i,'Parent Item']
        gerp.at[i,'level1']=gerp.at[i,'Child Item']
    elif level=='**2':
        gerp.at[i,'level0']=gerp.at[0,'Parent Item'] #가장 첫번째
        gerp.at[i,'level1']=gerp.at[i,'Parent Item']
        gerp.at[i,'level2']=gerp.at[i,'Child Item']
    elif level=='***3':
        gerp.at[i,'level0']=gerp.at[0,'Parent Item'] #가장 첫번째
        gerp.at[i,'level2']=gerp.at[i,'Parent Item']
        gerp.at[i,'level3']=gerp.at[i,'Child Item']
    elif level=='****4':
        gerp.at[i,'level0']=gerp.at[0,'Parent Item'] #가장 첫번째
        gerp.at[i,'level3']=gerp.at[i,'Parent Item']
        gerp.at[i,'level4']=gerp.at[i,'Child Item']
    elif level=='*****5':
        gerp.at[i,'level0']=gerp.at[0,'Parent Item'] #가장 첫번째
        gerp.at[i,'level4']=gerp.at[i,'Parent Item']
        gerp.at[i,'level5']=gerp.at[i,'Child Item']
    else:
        pass

#Level 3 - level1
empty_list=gerp.loc[gerp['Level']=='***3'].reset_index()
full_list=gerp.loc[gerp['Level']=='**2'].reset_index()
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level2']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level2']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#Level 3 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value=empty_list.at[i,"level1"]
    for j in range(len(gerp)):
        if empty_index==j:
            gerp.at[j,"level1"]=empty_value
        else:
            pass
#Level 4 - level2, level1
empty_list=gerp.loc[gerp['Level']=='****4'].reset_index()
full_list=gerp.loc[gerp['Level']=='***3'].reset_index()   
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level3']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level3']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#Level 4 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    for j in range(len(gerp)):
        if empty_index==j:
            gerp.at[j,"level1"]=empty_value1
            gerp.at[j,"level2"]=empty_value2
        else:
            pass
#Level 5 - level3,level2, level1
empty_list=gerp.loc[gerp['Level']=='*****5'].reset_index()
full_list=gerp.loc[gerp['Level']=='****4'].reset_index() 
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level4']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level4']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level3"]=full_list.at[j,'level3']
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']

#Level 5 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    empty_value3=empty_list.at[i,"level3"]
    for j in range(len(gerp)):
        if empty_index==j:
            gerp.at[j,"level1"]=empty_value1
            gerp.at[j,"level2"]=empty_value2
            gerp.at[j,"level3"]=empty_value3
        else:
            pass

gerp.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/gerpresult.xlsx')

###########################NPT###############################
npt=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/npt.xlsx')

# #파일에 따라 추가하기 (Test-1 O, Test-3 X)
npt=npt.drop([0],axis=0) #total 행 제거
npt.reset_index(drop=True, inplace=True)

#NPT level 숫자로 정리
for i in range(len(npt)):
    #level 정리
    level=str(npt.at[i,"Lvl"])
    level=int(level.replace('.',''))
    npt.at[i,"level"]=level

#parent part No 생성하기
for i in range(len(npt)):
    level=npt.at[i,"level"]
    #level로 순위 메기기
    if level!=0:
        for j in range(len(npt)):
            compare_level=npt.at[i-j,"level"]
            if compare_level<level:
                Parent_Part=npt.at[i-j,"Part No"]
                break
            else:
                continue
        npt.at[i,"Parent Part"]=Parent_Part
    else:
        #부모는 자기자신
        npt.at[i,"Parent Part"]=npt.at[i,"Part No"]

#level0,level1,level2,level3, level4, level5 생성하기
for i in range(len(npt)):
    level=npt.at[i,'level']
    if level==0:
        npt.at[i,'level0']=npt.at[i,'Part No']
    elif level==1:
        npt.at[i,'level0']=npt.at[i,'Parent Part']
        npt.at[i,'level1']=npt.at[i,'Part No']
    elif level==2:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level1']=npt.at[i,'Parent Part']
        npt.at[i,'level2']=npt.at[i,'Part No']
    elif level==3:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level2']=npt.at[i,'Parent Part']
        npt.at[i,'level3']=npt.at[i,'Part No']
    elif level==4:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level3']=npt.at[i,'Parent Part']
        npt.at[i,'level4']=npt.at[i,'Part No']
    elif level==5:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level4']=npt.at[i,'Parent Part']
        npt.at[i,'level5']=npt.at[i,'Part No']
    elif level==6:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level5']=npt.at[i,'Parent Part']
        npt.at[i,'level6']=npt.at[i,'Part No']
    elif level==7:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level6']=npt.at[i,'Parent Part']
        npt.at[i,'level7']=npt.at[i,'Part No']
    else:
        pass

#Level 3 - level1
empty_list=npt.loc[npt['level']==3].reset_index()
full_list=npt.loc[npt['level']==2].reset_index()
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level2']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level2']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 3 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value=empty_list.at[i,"level1"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value
        else:
            pass

#level 4 - level2, level1
empty_list=npt.loc[npt['level']==4].reset_index()
full_list=npt.loc[npt['level']==3].reset_index()   
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level3']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level3']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 4 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value1
            npt.at[j,"level2"]=empty_value2
        else:
            pass

#level 5 - level3,level2, level1
empty_list=npt.loc[npt['level']==5].reset_index()
full_list=npt.loc[npt['level']==4].reset_index() 
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level4']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level4']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level3"]=full_list.at[j,'level3']
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 5 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    empty_value3=empty_list.at[i,"level3"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value1
            npt.at[j,"level2"]=empty_value2
            npt.at[j,"level3"]=empty_value3
        else:
            pass

#level 6 - level5,level4,level3,level2,level1
empty_list=npt.loc[npt['level']==6].reset_index()
full_list=npt.loc[npt['level']==5].reset_index() 
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level5']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level5']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level4"]=full_list.at[j,'level4']
    empty_list.at[i,"level3"]=full_list.at[j,'level3']
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 6 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    empty_value3=empty_list.at[i,"level3"]
    empty_value4=empty_list.at[i,"level4"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value1
            npt.at[j,"level2"]=empty_value2
            npt.at[j,"level3"]=empty_value3
            npt.at[j,"level4"]=empty_value4
        else:
            pass

#level 7 - level6,level5,level4,level3,level2,level1
empty_list=npt.loc[npt['level']==7].reset_index()
full_list=npt.loc[npt['level']==6].reset_index() 
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level6']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level6']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level5"]=full_list.at[j,'level5']
    empty_list.at[i,"level4"]=full_list.at[j,'level4']
    empty_list.at[i,"level3"]=full_list.at[j,'level3']
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 7 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    empty_value3=empty_list.at[i,"level3"]
    empty_value4=empty_list.at[i,"level4"]
    empty_value5=empty_list.at[i,"level5"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value1
            npt.at[j,"level2"]=empty_value2
            npt.at[j,"level3"]=empty_value3
            npt.at[j,"level4"]=empty_value4
            npt.at[j,"level5"]=empty_value5
        else:
            pass

#비교할 NPT 데이터만 남기기 (Assembly Pull Type)
npt=npt[npt['Supply Type']=='Assembly Pull']
npt.reset_index(inplace=True,drop=True)

npt.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/nptresult.xlsx')

####################### compare npt and gerp #######################
#Match with GERP parent
#Change Parent Part As F3 (NPT and GERP match) remove TAW
f3_parent=gerp.at[0,"Parent Item"]
for i in range(len(npt)):
    parent_part=npt.at[i,'Parent Part'][:3]
    if parent_part=="TAW":
        npt.at[i,'Parent Part']=f3_parent
    else:
        pass

#match_list 시작
match_list=pd.DataFrame()
unique_gerp=gerp
model_name=gerp.at[0,"Parent Item"][:1]

#Screw,Tapping / Screw,Taptite / Duct / Packing / Resin,EPS / Damper 우선 배치 혼돈을 줄수 있음
for i in range(len(npt)):
    npt_part=npt.at[i,"Part No"]
    npt_parent=npt.at[i,"Parent Part"]
    npt_des=npt.at[i,"Desc."]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    npt_qty=npt.at[i,"Unit Qty"]
    match_number=npt.at[i,"Seq."]
    if npt_des=="Screw,Tapping": #Screw,tapping     
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part and gerp_qty==npt_qty and gerp_parent==npt_parent: #완전 일치문 (parent는 일치해야 함 드라이어)
                if gerp_price==npt_price:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_true"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_price"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)


    elif npt_des=="PCB Assembly,Complex":
        count=0
        used_list=pd.DataFrame()
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and count==0:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_re"]=gerp_data
                used_list.at[count,0]=j
                count=count+1
            elif gerp_des==npt_des and count==1:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_exc"]=gerp_data
                used_list.at[count,0]=j
                count=count+1
            elif gerp_des==npt_des and count==2:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_sub"]=gerp_data
                used_list.at[count,0]=j 
                count=count+1  
        used_list=used_list[0].to_list()
        unique_gerp=unique_gerp.drop(used_list,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)

    #Screw,Taptite 
    elif npt_des=='Screw,Taptite':
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part and npt_qty==gerp_qty: #완전 일치문 (parent는 일치하지 않는 경우가 있어 제외)
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_exc"]=gerp_data
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)

    #Duct Assembly 
    elif npt_des=='Duct Assembly':
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_sub_part=str(unique_gerp.at[j,"Child Item"])[:-1] # 끝자리 한자리까지 일치
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            if npt_part.__contains__(gerp_sub_part): # 끝자리 한자리까지 일치
                if gerp_part==npt_part: #완전일치
                    if gerp_price==npt_price:
                        gerp_data=unique_gerp.at[j,"Seq"]
                        match_list.at[match_number,"gerp_true"]=gerp_data
                    else:
                        gerp_data=unique_gerp.at[j,"Seq"]
                        match_list.at[match_number,"gerp_price"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_sub"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)

    #Packing 
    elif npt_des=='Packing':
        for j in range(len(unique_gerp)):
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part: #부모랑 des일치하면 우선 배정, 파트넘버 전부 같음
                if gerp_price==npt_price:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_true"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_price"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)

    #Resin,EPS -> DR 
    elif npt_des=='Resin,EPS':
        for j in range(len(unique_gerp)):
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part and npt_parent==gerp_parent: #부모랑 des일치하면 우선 배정, 파트넘버 전부 같음
                if gerp_price==npt_price:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_true"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_price"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)
    else:
        pass

#Screw,Customized -> 부모, description, Qty 만보고 먼저 매치
for i in range(len(npt)):
    npt_part=npt.at[i,"Part No"]
    npt_des=npt.at[i,"Desc."]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    npt_qty=npt.at[i,"Unit Qty"]
    match_number=npt.at[i,"Seq."]
    if npt_des=="Screw,Customized":
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part and gerp_qty==npt_qty: #완전 일치문 (parent는 일치하지 않는 경우가 있어 제외)
                if gerp_price==npt_price:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_true"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_price"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)
    else:
        pass


#Hanger,Upper -> 부모, description 만보고 먼저 매치
sub_count=0
count=0
gerp_drop=pd.DataFrame()
for i in range(len(npt)):
    npt_des=npt.at[i,"Desc."]
    npt_part=npt.at[i,"Part No"]
    npt_sub=str(npt.at[i,"Part No"])[:-2] #part넘버가 모두 숫자인 경우
    npt_price=npt.at[i,"Material Cost (LOC)"]
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_sub=str(unique_gerp.at[j,"Child Item"])[:-2] #part넘버가 모두 숫자인 경우
        #완전 일치문 (parent는 일치하지 않는 경우가 있어 제외)
        if gerp_des=='Hanger,Upper' and gerp_des==npt_des and gerp_part==npt_part: 
            if gerp_price==npt_price:
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_true"]=unique_gerp.at[j,"Seq"] #gerp는 j
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1
            else:
                match_numbre=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_price"]=unique_gerp.at[j,"Seq"] #gerp는 j
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1

        #sub로 매치
        elif gerp_des=='Hanger,Upper' and gerp_des==npt_des and gerp_sub==npt_sub: 
            if sub_count==0: #한번도 sub 매치 안된 경우
                match_number=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_sub"]=unique_gerp.at[j,"Seq"] #gerp는 j
                sub_count=sub_count+1
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1
            else:
                match_number=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_exc"]=unique_gerp.at[j,"Seq"] #gerp는 j
                sub_count=0
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1

        #door glass -> parent 다른게 두게
        elif gerp_des=="Door,Glass" and npt_des==gerp_des and gerp_sub==npt_sub:
            if gerp_part==npt_part: #parent 다르지만 part 일치
                match_number=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_price"]=unique_gerp.at[j,"Seq"] #gerp는 j
                sub_count=sub_count+1
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1
            else:
                match_number=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_exc"]=unique_gerp.at[j,"Seq"] #gerp는 j
                sub_count=0
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1

        else:
            continue

### HOW TO DROP 
unique_gerp.index=unique_gerp["Seq"] #index reset
if len(gerp_drop)==0:
    unique_gerp.reset_index(drop=True, inplace=True)
else:
    drop_index=gerp_drop.gerp.tolist()
    unique_gerp=unique_gerp.drop(drop_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)



# Holder -> 파트넘버, des 보고 바로 매치
for i in range(len(npt)):
    npt_parent=npt.at[i,"Parent Part"]
    npt_part=npt.at[i,"Part No"]
    npt_des=npt.at[i,"Desc."]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    npt_level=int(npt.at[i,"Lvl"][-1:])
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        if npt_des=="Holder" and npt_level>2 and npt_des==gerp_des and npt_part==gerp_part:
            if gerp_price==npt_price:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_true"]=gerp_data
            else:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_price"]=gerp_data
        else:
            continue
    else:
        pass


    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)

remain_match=pd.DataFrame()
count=0

for i in range(len(npt)):
    npt_parent=npt.at[i,"Parent Part"]
    npt_part=npt.at[i,"Part No"]
    npt_subpart=npt.at[i,"Part No"][:-2]
    npt_des=npt.at[i,"Desc."]
    npt_level=int(npt.at[i,"Lvl"][-1:])
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_subpart=unique_gerp.at[j,"Child Item"][:-2]
        gerp_des=unique_gerp.at[j,"Description"]
        remain_seq=unique_gerp.at[j,"Seq"]

        ############### Heater Assembly ###############
        if gerp_des=="Heater Assembly" and npt_des==gerp_des and count==0:
            gerp_data=unique_gerp.at[j,"Seq"]
            match_list.at[match_number,"gerp_true"]=gerp_data
            count=count+1

        elif gerp_des=="Heater Assembly" and npt_des==gerp_des and count==1:
            gerp_data=unique_gerp.at[j,"Seq"]
            match_list.at[match_number,"gerp_true"]=gerp_data
            count=count+1

    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)
####################################### Arrange - Matching Perfectly #######################################
for i in range(len(npt)):
    npt_part=npt.at[i,"Part No"]
    npt_des=npt.at[i,"Desc."]
    npt_parent=npt.at[i,"Parent Part"]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_parent=unique_gerp.at[j,"Parent Item"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        if gerp_part==npt_part and gerp_des==npt_des and gerp_parent==npt_parent:#완전 일치문
            if gerp_price==npt_price:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_true"]=gerp_data
            else:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_price"]=gerp_data
        else:
            continue
    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)


####################################### Arrange - Matching Substitute #######################################
for i in range(len(npt)):
    npt_part=str(npt.at[i,"Part No"])
    npt_des=npt.at[i,"Desc."]
    npt_parent=npt.at[i,"Parent Part"]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=str(unique_gerp.at[j,"Child Item"])
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_parent=unique_gerp.at[j,"Parent Item"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        if gerp_part[:-2]==npt_part[:-2] and gerp_des==npt_des and gerp_parent==npt_parent:
            gerp_data=unique_gerp.at[j,"Seq"]
            match_list.at[match_number,"gerp_sub"]=gerp_data
        else:
            continue
    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)    


#일치 조건문 - substitute   (part no 완전 다른 것 ==> 둘다 parent와 description)
for i in range(len(npt)):
    #Label,Barcode의 경우 매칭되지 않도록 주의 
    npt_des=npt.at[i,"Desc."]
    if npt_des=='Label,Barcode': #남는 데이터로 다시 돌리는 조건문에서 매칭 될 것
        pass
    else:
        npt_part=npt.at[i,"Part No"]
        npt_des=npt.at[i,"Desc."]
        npt_parent=npt.at[i,"Parent Part"]
        npt_price=npt.at[i,"Material Cost (LOC)"]
        match_number=npt.at[i,"Seq."]
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            if gerp_des==npt_des and gerp_parent==npt_parent:
                if gerp_part!=npt_part:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_exc"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)   


####################################### Arrange - Matching except parent #######################################
for i in range(len(npt)):
    npt_part=npt.at[i,"Part No"]
    npt_des=npt.at[i,"Desc."]
    npt_parent=npt.at[i,"Parent Part"]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_parent=unique_gerp.at[j,"Parent Item"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        if gerp_part==npt_part and gerp_des==npt_des:
            gerp_data=unique_gerp.at[j,"Seq"]
            match_list.at[match_number,"gerp_parent"]=gerp_data
        else:
            continue
    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True) 

####################################### Arrange - Matching except parent #######################################
### 매칭 안된 데이터 - 남는 데이터 가지고 다시 비교
#price, parent -> parent는 드롭 -> unique_add
match_list=match_list.reset_index()
unique_parent=pd.DataFrame()
count=0
for i in range(len(match_list)):
    price=str(match_list.at[i,"gerp_price"])
    parent=str(match_list.at[i,"gerp_parent"])
    if price!='nan' and parent!='nan':
        unique_parent.at[count,"Seq"]=match_list.at[i,"gerp_parent"]
        count=count+1
    else:
        pass

#unique_parent 존재 유무
unique_parent_count=len(unique_parent)
if unique_parent_count==0:
    remain_gerp=unique_gerp.dropna(subset=['Seq']) #seq빈 애들 drop
    remain_gerp.reset_index(drop=True, inplace=True)
else:
    A=unique_parent["Seq"].tolist()
    unique_add=gerp
    unique_add.index=unique_add.Seq
    unique_add=unique_add.iloc[A]
    unique_add.index=range(len(unique_gerp), len(unique_gerp)+len(unique_add))
    #unique_add 를 unique_gerp에 더하기 & 내림 차순
    remain_gerp=pd.concat([unique_gerp, unique_add], axis=0).sort_values(by=['Seq'],ascending=True)
    remain_gerp.reset_index(drop=True, inplace=True)
    #남은 것 다시 매칭하기 위해 행렬 정렬 -> 아래 조건문에서 자기 자리 찾아감
    match_list.index=match_list['index']

####################################### Arrange - remain gerp #######################################
### 매칭 안된 데이터 - 남는 데이터 가지고 다시 비교
############################ 2) NPT Sequence -> index #########################
match_list.index=match_list['index']
remain_match=pd.DataFrame()
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=remain_gerp.at[i,"Description"]
    remain_subdes=remain_gerp.at[i,"Description"][:-2]
    remain_parent=remain_gerp.at[i,"Parent Item"]
    remain_subparent=remain_gerp.at[i,"Parent Item"][:-2]
    remain_part=remain_gerp.at[i,"Child Item"]
    remain_subpart=remain_gerp.at[i,"Child Item"][:-2]
    remain_seq=remain_gerp.at[i,"Seq"]

    for j in range(len(npt)):
        npt_des=npt.at[j,"Desc."]
        npt_subdes=npt.at[j,"Desc."][:-2]
        npt_parent=npt.at[j,"Parent Part"]
        npt_part=npt.at[j,"Part No"]
        npt_subpart=npt.at[j,"Part No"][:-2]
        match_number=npt.at[j,"Seq."]

        ############### coil,sheet(sts) ###############
        if remain_des=="Coil,Steel(STS)" and npt_des=="Sheet,Steel(STS)" and npt_parent==remain_parent:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1
        
        ############### Coil,Steel(GI) ###############
        elif remain_des=='Coil,Steel(GI)':
            if npt_parent==remain_parent and npt_des=='Sheet,Steel(GA)':
                match_list.at[match_number,"gerp_re"]=remain_seq
                remain_match.at[count,"index"]=i
                count=count+1
            elif npt_part==remain_part and npt_des==remain_des:
                match_list.at[match_number,"gerp_re"]=remain_seq
                remain_match.at[count,"index"]=i
                count=count+1
            elif npt_des=="Sheet,Steel(GI)" and remain_parent==npt_parent:
                match_list.at[match_number,"gerp_re"]=remain_seq
                remain_match.at[count,"index"]=i
                count=count+1

        ############### Sheet,Steel(PCM) ###############
        elif remain_des=="Sheet,Steel(PCM)" and npt_des=="Sheet,Steel(PCM)" and npt_subpart==remain_subpart:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Sheet,Steel(GI) ###############
        elif remain_des=='Sheet,Steel(GI)' and npt_parent==remain_parent and npt_des=='Coil,Steel(GI)':
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Sheet,Steel(STS) ###############
        elif npt_des=='Sheet,Steel(STS)' and remain_des=='Coil,Steel(STS)':
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### resin,asa ###############
        elif npt_des=='Resin,A' and npt_parent==remain_parent:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### rotor, stator ###############
        elif remain_des.__contains__('Assembly,Combined') and remain_des.__contains__(npt_des):
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### rotor assembly ###############
        elif remain_des=='Rotor Assembly' and npt_des.__contains__('Rotor'):
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### stator assembly ###############
        elif remain_des=='Stator Assembly' and npt_des.__contains__('Stator'):
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### hanger pivot ###############
        elif remain_des.__contains__(',Pivot') and remain_des==npt_des and remain_subpart==npt_subpart:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### handle ###############
        elif remain_des=='Handle' and npt_des=='Handle Assembly' and npt_parent==remain_parent:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1
        
        ############### handle assembly ###############
        elif remain_des=='Handle Assembly' and npt_des=='Handle' and npt_parent==remain_parent:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Door Glass ###############
        elif remain_des=='Door,Glass'and gerp_parent==remain_parent and gerp_des=='Door,Glass':
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Barcode Label ###############
        elif remain_des=="Label,Barcode" and npt_subpart==remain_subpart and npt_des==remain_des:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Complex PCB ###############
        elif remain_des=="PCB Assembly,Complex" and npt_subpart==remain_subpart:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Wrap ###############
        elif remain_des=="Wrap" and npt_subpart==remain_subpart:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Control Panel ###############
        elif remain_des=="Panel Assembly,Control" and npt_subpart==remain_subpart:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Heater ###############
        elif remain_des=="Heater Assembly" and npt_subpart==remain_subpart:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        ############### Heater ###############
        elif remain_des=="Screw,Taptite" and npt_parent=="TAV35812101" and npt_part==remain_part:
            match_list.at[match_number,"gerp_re"]=remain_seq
            remain_match.at[count,"index"]=i
            count=count+1

        else:
            pass
    
match_list["index"]=match_list.index
match_list.reset_index(drop=True, inplace=True)

#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)


############### remain gerp 매칭 안되고 missing 된것
remain_gerp.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/remaingerp.xlsx')


####################################### match_list "match_digit" #######################################
sub_matchlist=match_list

####### columns not all contain #######
all_columns=set(["gerp_price","gerp_sub","gerp_exc","gerp_true","gerp_parent","gerp_re","index"])
submatchlist_columns=set(sub_matchlist.columns.tolist())
exclude_list=pd.DataFrame(all_columns.difference(submatchlist_columns))

for i in range(len(sub_matchlist)):
    match_list.at[i,"match_digit"]=0
    for j in range(len(exclude_list)):
        exclude_column=exclude_list.at[j,0]
        match_list.at[i,exclude_column]=np.nan

# add one more column
for i in range(len(match_list)):
    match_column=match_list.iloc[i]
    if str(match_column.gerp_price)!="nan":
        match_list.at[i,"match_digit"]=match_list.at[i,"match_digit"]+100000
    if str(match_column.gerp_sub)!="nan":
        match_list.at[i,"match_digit"]=match_list.at[i,"match_digit"]+10000
    if str(match_column.gerp_exc)!="nan":
        match_list.at[i,"match_digit"]=match_list.at[i,"match_digit"]+1000
    if str(match_column.gerp_true)!="nan":
        match_list.at[i,"match_digit"]=match_list.at[i,"match_digit"]+100
    if str(match_column.gerp_parent)!="nan":
        match_list.at[i,"match_digit"]=match_list.at[i,"match_digit"]+10
    if str(match_column.gerp_re)!="nan":
        match_list.at[i,"match_digit"]=match_list.at[i,"match_digit"]+1

####### columns not all contain -> column delete #######
if exclude_list.empty:
    pass
else:
    match_list=match_list.drop(exclude_list[0].values.tolist(),axis=1)

match_list=match_list.rename(columns={"index": "Seq."})
match_list.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/matchlist.xlsx')

######################################## submatchlist matching with match_digit
sub_matchlist=pd.DataFrame()
change_count=0

for i in range(len(match_list)):
    match_digit=match_list.at[i,"match_digit"]
    
    ############### true ###############
    if match_digit==100:
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_true']
        sub_matchlist.at[change_count,"gerp_true"]=match_list.at[i,'gerp_true']
        change_count=change_count+1

    ############### true, re ###############
    elif match_digit==101:
        #true
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_true']
        sub_matchlist.at[change_count,"gerp_true"]=match_list.at[i,'gerp_true']
        change_count=change_count+1
        #re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        change_count=change_count+1
    
    ############### true, price ###############
    elif match_digit==100100:
        #true
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_true']
        sub_matchlist.at[change_count,"gerp_true"]=match_list.at[i,'gerp_true']
        change_count=change_count+1
        #re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        change_count=change_count+1


    ############### true, parent ###############
    elif match_digit==110:
        #true
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_true']
        sub_matchlist.at[change_count,"gerp_true"]=match_list.at[i,'gerp_true']
        change_count=change_count+1
        #re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_parent']
        sub_matchlist.at[change_count,"gerp_parent"]=match_list.at[i,'gerp_parent']
        change_count=change_count+1
    
    ############### price ###############
    elif match_digit==100000:
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        change_count=change_count+1

    ############### parent ###############
    elif match_digit==10:
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_parent']
        sub_matchlist.at[change_count,"gerp_parent"]=match_list.at[i,'gerp_parent']
        change_count=change_count+1

    ############### parent, price ###############
    elif match_digit==100010:
        # price
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        change_count=change_count+1
        #parent
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_parent"]=match_list.at[i,'gerp_parent']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_parent']
        change_count=change_count+1
    
    ############### parent, re ###############
    elif match_digit==11:
        #parent
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_parent"]=match_list.at[i,'gerp_parent']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_parent']
        change_count=change_count+1
        # re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        change_count=change_count+1

    ############### parent, re,sub  ###############
    elif match_digit==10011:
        #parent
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_parent"]=match_list.at[i,'gerp_parent']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_parent']
        change_count=change_count+1
        # re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        change_count=change_count+1
        # sub
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1
    
    ############### parent, exc ###############
    elif match_digit==1010:
        #parent
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_parent"]=match_list.at[i,'gerp_parent']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_parent']
        change_count=change_count+1
        # exc
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_exc"]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_exc']
        change_count=change_count+1
    ############### sub ###############
    elif match_digit==10000:
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1

    ############### sub, price ###############
    elif match_digit==110000:
        #price
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        change_count=change_count+1
        #sub
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1

    ############### sub, exc ###############
    elif match_digit==11000:
        #sub
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1
        #exc
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_exc"]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_exc']
        change_count=change_count+1

    ############### sub, exc, re ###############
    elif match_digit==11001:
        #sub
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1
        #exc
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_exc"]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_exc']
        change_count=change_count+1
        #re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        change_count=change_count+1

    ############### sub, price,exc ###############
    elif match_digit==111000:
        #price
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        change_count=change_count+1
        #sub
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1
        #exc
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_exc"]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_exc']
        change_count=change_count+1
    
    ############### sub, price,re ###############
    elif match_digit==110001:
        #price
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        change_count=change_count+1
        #sub
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1
        #re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        change_count=change_count+1


    ############### sub, parent ###############
    elif match_digit==10010:
        #price
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_parent"]=match_list.at[i,'gerp_parent']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_parent']
        change_count=change_count+1
        #sub
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1

    ############### re ###############
    elif match_digit==1:
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        change_count=change_count+1

    ############### re, price ###############
    elif match_digit==100001:
        #price
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        change_count=change_count+1
        #re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        change_count=change_count+1

    ############### exc ###############
    elif match_digit==1000:
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_exc"]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_exc']
        change_count=change_count+1

    ############### exc, price ###############
    elif match_digit==101000:
        #price
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"count"]=1 #final table false column
        change_count=change_count+1
        #exc
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_exc"]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"count"]=1 #final table false column
        change_count=change_count+1

    ############### exc, price, re ###############
    elif match_digit==101001:
        #price
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_price"]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_price']
        sub_matchlist.at[change_count,"count"]=1 #final table false column
        change_count=change_count+1
        #exc
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_exc"]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"count"]=1 #final table false column
        change_count=change_count+1
        #re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"count"]=1 #final table false column
        change_count=change_count+1

    ############### parent, re, sub, exc ###############
    elif match_digit==11011:
        #parent
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_parent"]=match_list.at[i,'gerp_parent']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_parent']
        change_count=change_count+1
        #sub
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_sub"]=match_list.at[i,'gerp_sub']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_sub']
        change_count=change_count+1
        #re
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_re"]=match_list.at[i,'gerp_re']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_re']
        change_count=change_count+1
        #exc
        sub_matchlist.at[change_count,"Seq."]=match_list.at[i,'Seq.']
        sub_matchlist.at[change_count,"gerp_exc"]=match_list.at[i,'gerp_exc']
        sub_matchlist.at[change_count,"gerpSeq."]=match_list.at[i,'gerp_exc']
        change_count=change_count+1
    
    else:
        print("[ERROR] match digit another login found - check match_digit: "+str(match_digit))
        print(match_digit)


# match=False -> count=0 matching
match_join = pd.merge(npt['Seq.'], sub_matchlist, on ='Seq.', how ='left')    

#save file
match_join.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/matchjoin.xlsx')

######################################### final table
#final table
final_table=pd.DataFrame()
match_column=pd.DataFrame()

for i in range(len(match_join)):
    ######################################### NPT #########################################
    npt_seq=match_join.at[i,"Seq."]
    #sequence 일치 칼럼을 뽑아내고
    npt_column=npt[npt["Seq."]==npt_seq]
    npt_column.reset_index(inplace=True, drop=True) #index 넘버 상관 X하고 0으로 
    #npt list
    final_table.at[i,"Seq."]=npt_seq
    final_table.at[i,"level"]=npt_column.at[0,"level"]
    final_table.at[i,"Parent Part"]=npt_column.at[0,"Parent Part"]
    final_table.at[i,"Part No"]=npt_column.at[0,"Part No"]
    final_table.at[i,"Desc."]=npt_column.at[0,"Desc."]
    final_table.at[i,"UIT"]=npt_column.at[0,"UIT"]
    final_table.at[i,"UOM"]=npt_column.at[0,"UOM"]
    final_table.at[i,"Unit Qty"]=npt_column.at[0,"Unit Qty"]
    final_table.at[i,"Accum Qty"]=npt_column.at[0,"Accum Qty"]
    final_table.at[i,"Supplier Name"]=npt_column.at[0,"Supplier Name"]
    final_table.at[i,"Curr."]=npt_column.at[0,"Curr."]
    final_table.at[i,"Unit Price"]=npt_column.at[0,"Unit Price"]
    final_table.at[i,"Exchange Rate(LOC)"]=npt_column.at[0,"Exchange Rate(LOC)"]
    final_table.at[i,"Unit Price (USD)"]=npt_column.at[0,"Unit Price (USD)"]
    final_table.at[i,"Material Cost (LOC)"]=npt_column.at[0,"Material Cost (LOC)"]
    #match list
    final_table.at[i,"npt end"]=''
    final_table.at[i,"match"]=''
    final_table.at[i,"price match"]=''
    final_table.at[i,"gerp start"]=''

    ######################################### GERP #########################################
    #GERP Matching
    gerp_seq=match_join.at[i,"gerpSeq."]
    gerp_seq_exist=str(gerp_seq)
    match_sub=match_join.at[i,"gerp_sub"]
    #######################GERP Match False####################
    if gerp_seq_exist=='nan': #empty dataframe -> count0
        final_table.at[i,"Seq"]=''
        final_table.at[i,"Level"]=''
        final_table.at[i,"Parent Item"]=''
        final_table.at[i,"Child Item"]=''
        final_table.at[i,"Description"]=''
        final_table.at[i,"Item Type"]=''
        final_table.at[i,"Qty Per Assembly"]=''
        final_table.at[i,"Material Cost"]=''
        final_table.at[i,"QPA*Material Cost"]=''
        #match -> not match
        final_table.at[i,'match']=False
        #나중에 채울 price match
        final_table.at[i,"price match"]=''
    #######################GERP Match True,PriceChange,Substitute####################
    else:
        #gerp information extract
        gerp_column=gerp[gerp["Seq"]==gerp_seq]
        gerp_column.reset_index(inplace=True, drop=True) #index 넘버 상관 X하고 0으로 
        
        final_table.at[i,"Seq"]=gerp_seq
        final_table.at[i,"Level"]=int(str(gerp_column.at[0,"Level"])[-1:])
        final_table.at[i,"Parent Item"]=gerp_column.at[0,"Parent Item"]
        final_table.at[i,"Child Item"]=gerp_column.at[0,"Child Item"]
        final_table.at[i,"Description"]=gerp_column.at[0,"Description"]
        final_table.at[i,"Item Type"]=gerp_column.at[0,"Item Type"]
        final_table.at[i,"Qty Per Assembly"]=gerp_column.at[0,"Qty Per Assembly"]
        final_table.at[i,"Material Cost"]=gerp_column.at[0,"Material Cost"]
        final_table.at[i,"QPA*Material Cost"]=gerp_column.at[0,"QPA*Material Cost"]
        final_table.at[i,"match"]=True
        final_table.at[i,"price match"]=''


#######################Add MTL Column####################
mtl=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/gerp.xlsx',sheet_name='MTL Cost')
mtl=mtl[['Item No','Item Cost','Material Overhead Cost','Creation Period']]
#mtl 부분 데이터 정리
mtl=mtl.rename(columns={'Item No':'Child Item','Creation Period':'PAC Creation','Item Cost':'MTL Cost','Material Overhead Cost':'MTL OH'})
mtl['MTL OH']=mtl['MTL OH'].round(8)
mtl['Net Material']=(mtl['MTL Cost']-mtl['MTL OH']).round(8)
mtl['mtl start']=''

#칼럼 순서 바꾸기
mtl=mtl[['mtl start','Child Item','PAC Creation','MTL Cost','Net Material','MTL OH']]

# vlookup 하고 내림차순 정리
mtl_final=pd.merge(final_table, mtl, on='Child Item', how='left')
final_table = mtl_final.sort_values(by=['Seq.'],ascending=True)


######################## Match Type : Substitute / True ########################
############ Substitute
for i in range(len(final_table)):
    final_match=str(final_table.at[i,"match"])
    npt_part=str(final_table.at[i,"Part No"])
    gerp_part=str(final_table.at[i,"Child Item"])
    if gerp_part!='':
        if npt_part!=gerp_part:
            final_table.at[i,"match"]="Substitute"

############################ Price Match #########################
############## price match - general ##############
final_table["price match"]=(final_table['Net Material']-final_table['Unit Price (USD)'])*final_table['Qty Per Assembly']
cast_to_type = {'price match': float} # round error -> datatype
final_table = final_table.astype(cast_to_type) # round error -> datatype
final_table["price match"]=final_table["price match"].round(8)


############## price match - exceptional - Heater Substitute ##############
## 자신의 child 다른 parent 매칭 but substitute part
for i in range(len(final_table)):
    child_part=str(final_table.at[i,"Child Item"])[:-2]  #gerp
    part_des=final_table.at[i,"Desc."] #npt
    price_match=final_table.at[i,"match"]
    if part_des.__contains__('Heater Assembly') and price_match=='Substitute':
        for j in range(len(final_table)):
            another_parent=final_table.at[j,"Parent Part"][:-2]#npt
            another_qty=final_table.at[j,"Qty Per Assembly"]
            another_price=final_table.at[j,"Material Cost (LOC)"]
            another_des=final_table.at[j,"Desc."]
            another_match=final_table.at[j,"match"]
            if another_des=='Heater Assembly' and another_parent==child_part and another_match!=("Substitute"):
                final_table.at[i,"price match"]=(final_table.at[i,"Net Material"]-another_price)*final_table.at[i,"Qty Per Assembly"]
    else:
        pass

# ############## price match Npt T, Child NPT G,D,S ##############
# for i in range(len(final_table)):
#     npt_uit=final_table.at[i,"UIT"]
#     npt_part=final_table.at[i,"Part No"]
#     npt_match=final_table.at[i,"match"]
#     npt_part=final_table.at[i,"Part No"]
#     npt_gerp_part=final_table.at[i,"Child Item"]
#     npt_child_uit=final_table.at[i,"Item Type"]
#     npt_cost=final_table.at[i,"Net Material"]
#     npt_qty=final_table.at[i,"Qty Per Assembly"]
#     child_seq=int(final_table.at[i,"Seq."])+1
#     child_index=final_table[final_table['Seq.']==child_seq].index

#     # T & Substitute
#     if npt_uit=="T" and npt_gerp_part!='' and npt_part!=npt_gerp_part:
#         child_seq=int(final_table.at[i,"Seq."])+1
#         child_row=final_table.index[final_table['Seq.']==child_seq].tolist()
#         child_index=child_row[0]
#         child_part=final_table.at[child_index,"Parent Part"]
#         child_cost=final_table.at[child_index,"Material Cost (LOC)"]
#         child_g_cost=final_table.at[child_index,"Unit Price (USD)"]
#         child_gd_cost=final_table.at[child_index,"Net Material"]
#         child_g_qty=final_table.at[child_index,"Qty Per Assembly"]
#         child_uit=final_table.at[child_index,"UIT"]

#         if child_uit=="G" and child_part==npt_part:
#             final_table.at[i,"price match"]=round((npt_cost-child_cost)*npt_qty,8)
#             final_table.at[child_index,"price match"]=round((child_gd_cost-child_g_cost)*child_g_qty*final_table.at[j-1,"Qty Per Assembly"],8)
#         elif child_uit=="D" and child_part==npt_part:
#             final_table.at[i,"price match"]=round((npt_cost-child_cost)*npt_qty,8)
#         elif child_uit=="S" and child_part==npt_part:
#             final_table.at[i,"price match"]=round((npt_cost-child_cost)*npt_qty,8)


######################## Match Type : Price Change / True / Substitute ########################
############ True, Price Change
for i in range(len(final_table)):
    price_match=round(final_table.at[i,'price match'],2)
    final_match=final_table.at[i,"match"]
    if final_match==True:
        if abs(price_match)<0.001 or str(price_match)=='nan': 
            final_table.at[i,'match']=True
        else:
            final_table.at[i,'match']="Price Change"


######################## Total Price Calculation ########################
#price table -> final result
final_table['final result']=''
final_table.at[0,'Price Change']=0
final_table.at[0,'Substitute Price Change']=0
for i in range(len(final_table)):
    final_match=str(final_table.at[i,"match"])
    if final_match=='Substitute':
        price_value=str(final_table.at[i,"price match"])
        if price_value=='nan':
            pass
        else:
            final_table.at[0,'Substitute Price Change']=final_table.at[0,'Substitute Price Change']+final_table.at[i,"price match"]
    elif final_match=="Price Change":
        price_value=str(final_table.at[i,"price match"])
        if price_value=='nan':
            pass
        else:
            final_table.at[0,'Price Change']=final_table.at[0,'Price Change']+final_table.at[i,"price match"]

# #price match column round 2 -> data
final_table['Price Change']=round(final_table['Price Change'],2)
final_table['Substitute Price Change']=round(final_table['Substitute Price Change'],2)


######################## NPT Matching with several GERP -> NPT Empty ########################
for i in range(1,len(final_table)): #0은 -1과 비교할 수 없음으로
    npt_seq1=final_table.at[i,"Seq."]
    npt_seq2=final_table.at[i-1,"Seq."]
    if npt_seq1==npt_seq2:
        #sequence남기기
        final_table.at[i,"level"]=np.nan
        final_table.at[i,"Parent Part"]=np.nan
        final_table.at[i,"Part No"]=np.nan
        final_table.at[i,"Desc."]=np.nan
        final_table.at[i,"UOM"]=np.nan
        final_table.at[i,"Unit Qty"]=np.nan
        final_table.at[i,"Accum Qty"]=np.nan
        final_table.at[i,"Supplier Name"]=np.nan
        final_table.at[i,"Curr."]=np.nan
        final_table.at[i,"Unit Price"]=np.nan
        final_table.at[i,"Exchange Rate(LOC)"]=np.nan
        final_table.at[i,"Unit Price (USD)"]=np.nan
        final_table.at[i,"Material Cost (LOC)"]=np.nan


final_table.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0602/DR/final_table.xlsx')