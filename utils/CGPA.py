from pandas import *
from tkinter.filedialog import *
def excel_to_dataframe(data):
    try:
        civil=read_excel(data,sheet_name=["Civil"])
        civil=civil["Civil"]
    except:
        civil=DataFrame()
    try:
        eee=read_excel(data,sheet_name=["EEE"])
        eee=eee["EEE"]
    except:
        eee=DataFrame()
    try:
        mech=read_excel(data,sheet_name=["Mechanical"])
        mech=mech["Mechanical"]
    except:
        mech=DataFrame()
    try:
        ece=read_excel(data,sheet_name=["ECE"])
        ece=ece["ECE"]
    except:
        ece=DataFrame()
    try:
        cse=read_excel(data,sheet_name=["CSE"])
        cse=cse["CSE"]
    except:
        cse=DataFrame()
    return civil,eee,mech,ece,cse

def CGPA_cal(sem_selection,sem_file_list):    
    civil_final_df=DataFrame()
    eee_final_df=DataFrame()
    mech_final_df=DataFrame()
    ece_final_df=DataFrame()
    cse_final_df=DataFrame()
    
    def final_df_cal(final_df,df,i):
        df=df[["Roll No","Points","Total Credits","SGPA","Backlogs"]]
        df=df.rename(columns={"Total Credits":"Total Credits sem"+str(i+1),"Points":"Points sem"+str(i+1),"SGPA":"SGPA sem"+str(i+1)})
        if len(final_df.columns)==0:
            final_df=df
        else:
            final_df=merge(final_df,df,"outer",on="Roll No")
        final_df=final_df.fillna(0)
        return final_df
    for i in range(len(sem_selection)):
        if sem_selection[i]==1:
            try:
                civil_df,eee_df,mech_df,ece_df,cse_df=excel_to_dataframe(sem_file_list[i])
                if civil_df.empty and eee_df.empty and mech_df.empty and ece_df.empty and cse_df.empty:
                    raise Exception
                try:
                    civil_final_df=final_df_cal(civil_final_df,civil_df,i)
                except:
                    pass
                try:
                    eee_final_df=final_df_cal(eee_final_df,eee_df,i)
                except:
                    pass
                try:
                    mech_final_df=final_df_cal(mech_final_df,mech_df,i)
                except:
                    pass
                try:
                    ece_final_df=final_df_cal(ece_final_df,ece_df,i)
                except:
                    pass
                try:
                    cse_final_df=final_df_cal(cse_final_df,cse_df,i)
                except:
                    pass
            except :
                return i+1
            try:
                if "Roll No" not in civil_df.columns and "Total Credits" not in civil_df.columns and "SGPA" not in civil_df.columns and "Backlogs" not in civil_df.columns and not civil_df.empty:
                    return i+1
            except:
                pass
            try:
                if "Roll No" not in eee_df.columns and "Total Credits" not in eee_df.columns and "SGPA" not in eee_df.columns and "Backlogs" not in eee_df.columns and not eee_df.empty:
                    return i+1
            except:
                pass
            try:
                if "Roll No" not in mech_df.columns and "Total Credits" not in mech_df.columns and "SGPA" not in mech_df.columns and "Backlogs" not in mech_df.columns and not mech_df.empty:
                    return i+1
            except:
                pass
            try:
                if "Roll No" not in ece_df.columns and "Total Credits" not in ece_df.columns and "SGPA" not in ece_df.columns and "Backlogs" not in ece_df.columns and not ece_df.empty:
                    return i+1
            except:
                pass
            try:
                if "Roll No" not in cse_df.columns and "Total Credits" not in cse_df.columns and "SGPA" not in cse_df.columns and "Backlogs" not in cse_df.columns and not cse_df.empty:
                    return i+1
            except:
                pass
                
    def CGPA_calculations(df):   
        if df.empty:
            return
        df["CGPA"]=0
        df["Total backlogs"]=0
        x=len(df.columns)//4
        value=0
        gpa=0
        backlogs=0
        for i in range(len(df)):
            for j in range(x):
                gpa+=(df.iloc[i,1+4*j])
                value+=(df.iloc[i,2+(4*j)])
                backlogs+=df.iloc[i,4+(4*j)]
            df.loc[i,"CGPA"]=gpa/value
            df.loc[i,"Total backlogs"]=backlogs
            value=0
            gpa=0
            backlogs=0
        df=df.drop(df.columns[4::4],axis=1)
        return df  
    try:  
        civil_final_df=CGPA_calculations(civil_final_df)
    except:
        pass
    try:
        eee_final_df=CGPA_calculations(eee_final_df)
    except:
        pass
    try:
        mech_final_df=CGPA_calculations(mech_final_df)
    except:
        pass
    try:
        ece_final_df=CGPA_calculations(ece_final_df)
    except:
        pass
    try:
        cse_final_df=CGPA_calculations(cse_final_df)
    except:
        pass
    files=[('xlsx files','*.xlsx')]
    file=asksaveasfile(mode='wb',filetypes = files,defaultextension=files)
    with ExcelWriter(file,engine='openpyxl',mode='w') as output:
        try:
            civil_final_df.to_excel(output,sheet_name="Civil",index=False)
        except:
            pass
        try:
            eee_final_df.to_excel(output,sheet_name="EEE",index=False)
        except:
            pass
        try:
            mech_final_df.to_excel(output,sheet_name="Mechanical",index=False)
        except:
            pass
        try:
            ece_final_df.to_excel(output,sheet_name="ECE",index=False)
        except:
            pass
        try:
            cse_final_df.to_excel(output,sheet_name="CSE",index=False)
        except:
            pass
    return 0