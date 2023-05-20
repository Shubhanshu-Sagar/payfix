import pandas as pd
import numpy as np
import csv
from openpyxl import Workbook
import openpyxl
import os
import jinja2
import pdfkit
import shutil
def clear():
    os.system("cls")
#Defining function
def increment(bp,ittr):
    t_list=[]
    for i in range(1,ittr+1):
            increasebp = ((bp+(bp*0.03)))
            inc_bp =round(increasebp,-2)
            t_list.append(inc_bp)
            bp=inc_bp
    return t_list[ittr-1]


def option_availed(lev,bp):
    print(lev)
    print(type(bp))
    temp_bp = bp
    for i in ptt[lev]:
        if bp==i:
            
            print("Basic Pay is in existing pay level")
            break
    
    
    
    
    
#    after availing option         
    obp_sl=increment(bp,2)
    print(f" new basic after excercising option in same level is : {obp_sl}")
    obp_nl=""
    next_level=str(int(lev)+1)
    for i in ptt[next_level]:
        if i >= obp_sl:
            obp_nl=i
            break
            
    
    row_next_l=ptt[ptt[next_level]==obp_nl].index.item()
    
    for k in ptt[next_level]:
        if k >= bp:
            limit_fix=k
            break
    spl_dict={
         "tilldatebasic":limit_fix,
        "nextlevel":next_level,
        "newrow":row_next_l,
        "newbasic":obp_nl,
       }
    return spl_dict
def dnicalc(date):
    macp_dlist=date.split("/")
    macp_date=int(macp_dlist[0])
    macp_month=int(macp_dlist[1])
    macp_year=int(macp_dlist[2])
    inc_month=input("Increment Month /n press '7' for july /n press '1' for january")
    if inc_month =='7':
        inc_dlist=[1,7,macp_year]
        print(f"your Increment date is {inc_dlist[0]} / {inc_dlist[1]} / {inc_dlist[2]}")
        l_day=30
        l_month=6
        l_year=macp_year
        dni_date=[1,1,macp_year+1]
    elif inc_month =='1':
        inc_dlist=[1,1,macp_year+1]
        print(f"your Increment date is {inc_dlist[0]} / {inc_dlist[1]} / {inc_dlist[2]}")
        l_day=31
        l_month=12
        l_year=macp_year
        dni_date=[1,7,macp_year+1]
    else:
        print("Enter valid input")
    limitdate=str(f"{l_day} / {l_month} / {l_year}")
    macpdate=str(f"{macp_date} / {macp_month} / {macp_year}")
    incdate =str(f"{inc_dlist[0]} / {inc_dlist[1]} / {inc_dlist[2]}")
    dnidate = str(f"{dni_date[0]} / {dni_date[1]} / {dni_date[2]}")
    return [macpdate,limitdate,incdate,dnidate]

def nooption(lev,bp):
    newbp=increment(bp,1)
    new_fix_bp=""
    next_level=str(int(lev)+1)
    for i in ptt[next_level]:
        if i>=newbp:
            new_fix_bp=i
            break
    nextrow=ptt[ptt[next_level]==new_fix_bp].index.item()
    print(f"""New Basic After Availing No option : 
              {new_fix_bp} in level {next_level} row :{nextrow} """)
# MAKING PAY MATRIX TABLE 
filename = "pm_f.xlsx"
raw_table=np.array(pd.read_excel(filename,header=0))

# arr1=np.array(raw_table)
temp = """1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
37
38
39
40
"""
rows=temp.split()
# print(rows)
temp2="""
1	2	3	4	5	6	7	8	9	10	11	12	13	13A	14	15	16	17	18

"""
cols=temp2.split()
# print(cols)

ptt=pd.DataFrame(data=raw_table,index=rows,columns=cols)
#Collecting Data
emp_data={}
emp_data["emp_code"]=""
emp_data["existing_basic"]=""
try:
    clear()
    temp_var=int(input("Enter Employee Code \n :"))
    
except ValueError:
    print("Error : Employee code should be in numerical digits")
    
    
else:

    if type(temp_var)==int:
        if len(str(temp_var))==8:
            print("Valid Employee Code")
            emp_data["emp_code"]=temp_var
        else:
            print(" Input Error :- The Employee Code Must be 8 digit long , Try Again")
    else:
        pass
emp_data["user_name"]=input(" Enter Employee's User Name : \n ")
emp_data["desig"]=input(" Enter Employee's Designation : \n ")
emp_data["father"]=input(" Enter Employee's Father Name : \n ")
try:
    temp_var=float(input("Enter Employee Existing Basic Pay \n :"))
except ValueError:
    print("Error : Employee Basic should be in numerical digits")
    exit()
else:
    emp_data["existing_basic"]=temp_var
    print("Valid Basic Pay")
    
emp_data["exist_level"]=input("Enter Employee's Existing Pay Level")
emp_data["macp_date"] = input("Enter The Macp date ''(DD/MM/YYYY)''")
option=input("Eligible For Option ? (Yes / No)")
emp_data["option_choice"]=option.lower()
clear()
if emp_data["option_choice"]=="yes" :
    print("option availed")
    basic_list=option_availed(emp_data["exist_level"],emp_data["existing_basic"])
    date_list=dnicalc(emp_data["macp_date"])
    print(date_list)
    print(basic_list)
    finalsheet=f"""
    --------------------------------------------------------------------------------------------------------------------------------------------------------------------
    1. Existing Basic Pay : {emp_data["existing_basic"]} (level {emp_data["exist_level"]})
    2. Basic pay after granting macp
        from {date_list[0]} to {date_list[1]} : {basic_list["tilldatebasic"]} (Level {basic_list["nextlevel"]}) 
    
    3. option availed : YES 
    4. Refixation of pay after availing option on {date_list[2]} is : {basic_list['newbasic']} ( Level : {basic_list["nextlevel"]} , Row : {basic_list["newrow"]})
    5. DNI : {date_list[3]}


    --------------------------------------------------------------------------------------------------------------------------------------------------------------------
    """
    f_data_list=[emp_data["emp_code"],emp_data["user_name"],emp_data["desig"],emp_data["father"],emp_data["existing_basic"],emp_data["exist_level"],date_list[0],date_list[1],basic_list["tilldatebasic"],date_list[2],basic_list['newbasic'],basic_list["nextlevel"],basic_list["newrow"],date_list[3]]
    print(finalsheet)
    template_dict={
        'e_name':emp_data["user_name"],
        'e_code':emp_data["emp_code"],
        'desig_name':emp_data["desig"],
        'father_name':emp_data["father"],
        'basic_pay':int(emp_data["existing_basic"]),
        'exist_level':emp_data["exist_level"],
        'macp_date':date_list[0],
        'user_response':emp_data["option_choice"],
        'date_1':date_list[0],
        'date_2':date_list[1],
        'enhanced_increment':int(basic_list["tilldatebasic"]),
        'level_name1':basic_list["nextlevel"],
        'new_basic':int(basic_list['newbasic']),
        'new_level':basic_list["nextlevel"],
        'next_idate':date_list[3]


    }
    filename=input("Enter File Name Please : ")
    extension_name=filename+".pdf"
    template_loader = jinja2.FileSystemLoader('./')
    template_env = jinja2.Environment(loader=template_loader)

    template = template_env.get_template('template1.html')
    output_text = template.render(template_dict)

    config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')
    pdfkit.from_string(output_text, extension_name, configuration=config, css='style.css')

    
    os.rename(f'./{extension_name}', f'C:/Users/Shubhanshu sagar/Desktop/macpreports/{extension_name}')
    with open("output.csv","a",newline="") as File:
        writer = csv.writer(File)
        writer.writerow(f_data_list)
    File.close()
    


    
    
    
    
    
    
    
    
elif emp_data["option_choice"]=="no" :
    print("no option availed")
    nooption(emp_data["exist_level"],emp_data["existing_basic"])
    
else:
    print("Input Valid Option")
