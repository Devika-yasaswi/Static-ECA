from pandas import *
from tkinter.filedialog import *
def branch_calculation(data):
    df=DataFrame(columns=["subject","No.of students registered","No.of students appeared","Absentees","Pass Percentage"])
    for i in list(data.columns)[1:-7]:
        temp_list=[i]
        sub=data[i].tolist()
        temp_list.append(len(sub))
        total=len(sub)-(sub.count("AB")+sub.count("ABSENT"))
        temp_list.append(total)
        temp_list.append(sub.count("AB")+sub.count("ABSENT"))
        passed=sub.count("A+")+sub.count("A")+sub.count("B")+sub.count("C")+sub.count("D")+sub.count("E")+sub.count("COMPLE")+sub.count("COMPLETED")
        temp_list.append(passed/total*100)
        df.loc[len(df.index)]=temp_list
    count=0
    count1=0
    for i in range(len(data)):
        new=data.iloc[i,1:-7]    
        new=list(set(new))
        if len(new)==1 and (new[0]=="AB" or new[i]=="ABSENT"):
            count+=1
        if len(new)!=1 and ("F" in new or "MP" in new or  "AB" in new or "ABSENT" in new or "MP" in new):
            count1+=1
    df1=DataFrame(columns=["subject","No.of students registered","No.of students appeared","Absentees","Pass Percentage"])
    new=["Total",len(data),len(data)-count,count,(len(data)-count-count1)/(len(data)-count)*100]
    df1.loc[len(df1.index)]=new
    data=data.sort_values(by=["Points"])
    df2=DataFrame(columns=["Place","Roll No","Points","SGPA"])
    x=1
    temp=data.iloc[-1,-1]
    for i in range(1,len(data)):
        if x<3 or temp==data.iloc[-i,-1]:
            if temp!=data.iloc[-i,-1]:
                x+=1
            new=[x,data.iloc[-i,0],data.iloc[-i,-2],data.iloc[-i,-1]]
            temp=data.iloc[-i,-1]
            df2.loc[len(df2.index)]=new     
        elif x==3:
            break
    df=concat([df,df1],axis=0)
    return df,df2
    

def get_statistics(file):
    try:
        civil_data=read_excel(file,sheet_name=["CE"])
        civil_data=DataFrame(civil_data["CE"])
    except:
        pass
    try:
        eee_data=read_excel(file,sheet_name=["EEE"])
        eee_data=eee_data["EEE"]
    except:
        pass
    try:
        mech_data=read_excel(file,sheet_name=["ME"])
        mech_data=mech_data["ME"]
    except:
        pass
    try:
        ece_data=read_excel(file,sheet_name=["ECE"])
        ece_data=ece_data["ECE"]
    except:
        pass
    try:
        cse_data=read_excel(file,sheet_name=["CSE"])
        cse_data=cse_data["CSE"]
    except:
        pass
    try:
        civil_data,civil_top=branch_calculation(civil_data)
    except:
        pass
    try:
        eee_data,eee_top=branch_calculation(eee_data)
    except:
        pass
    try:
        mech_data,mech_top=branch_calculation(mech_data)
    except:
        pass
    try:
        ece_data,ece_top=branch_calculation(ece_data)
    except:
        pass
    try:
        cse_data,cse_top=branch_calculation(cse_data)
    except:
        pass
    with ExcelWriter(file,engine='openpyxl',mode='a',if_sheet_exists="overlay") as output:
        try:
            civil_data.to_excel(output,sheet_name="CE stats",index=False)
            civil_top.to_excel(output,sheet_name="CE stats",index=False,startcol=7)
        except:
            pass
        try:
            eee_data.to_excel(output,sheet_name="EEE stats",index=False)
            eee_top.to_excel(output,sheet_name="EEE stats",index=False,startcol=7)
        except:
            pass
        try:
            mech_data.to_excel(output,sheet_name="ME stats",index=False)
            mech_top.to_excel(output,sheet_name="ME stats",index=False,startcol=7)
        except:
            pass
        try:
            ece_data.to_excel(output,sheet_name="ECE stats",index=False)
            ece_top.to_excel(output,sheet_name="ECE stats",index=False,startcol=7)
        except:
            pass
        try:
            cse_data.to_excel(output,sheet_name="CSE stats",index=False)
            cse_top.to_excel(output,sheet_name="CSE stats",index=False,startcol=7)
        except:
            pass