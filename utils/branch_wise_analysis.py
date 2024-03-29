from pandas import *
import matplotlib 
matplotlib.use('agg')
import matplotlib.pyplot as plt
from openpyxl import load_workbook,drawing
import os
def delete_branch (file,branch):
    wb = load_workbook(file)
    sheets=wb.get_sheet_names()
    for i in sheets:
        if branch not in i:            
            wb.remove(wb[i])
        elif branch !=i and i != branch + ' stats':
            wb.remove(wb[i])
    wb.save(file)

def branchwise_analysis(file,branch):
    data=read_excel(file,sheet_name=[branch])
    data=DataFrame(data[branch])
    delete_branch(file,branch)
    for i in list(data.columns)[1:-7]:
        final_labels=[]
        final_grades=[]
        grades=data[i].tolist()
        df=DataFrame({'Roll No':data['Roll No'],i:grades})
        new=[grades.count("A+")]
        new.append(grades.count("A"))
        new.append(grades.count("B"))
        new.append(grades.count("C"))
        new.append(grades.count("D"))
        new.append(grades.count("E"))
        new.append(grades.count("F"))
        new.append(grades.count("AB")+grades.count("ABSENT"))
        new.append(grades.count("MP"))
        new.append(grades.count("COMPLE")+grades.count("COMPLETED"))
        my_labels=["A+","A","B","C","D","E","F","AB","MP","COMPLETED"]
        for j in range(len(new)):
            if new[j]!=0:
                final_grades.append(new[j])
                final_labels.append(my_labels[j])
        plt.pie(final_grades, labels=final_labels, autopct="%.2f%%")
        strFile="./Piechart.png"
        plt.savefig(strFile)
        plt.clf()
        plt.close()
        df1=DataFrame(columns=['Grades','No.of Students'])
        for k in range(len(new)):
            df1.loc[len(df1.index)]=[my_labels[k],new[k]]
        writer=ExcelWriter(file,engine='openpyxl',mode='a',if_sheet_exists="overlay")
        df.to_excel(writer,sheet_name=i[0:31],index=False)
        df1.to_excel(writer,sheet_name=i[0:31],index=False,startrow=len(df)+3)
        workbook =writer.book
        ws = workbook[i[0:31]]
        img = drawing.image.Image(strFile)
        img.anchor = 'D'+str(len(df)+3)
        ws.add_image(img)
        writer.save()