import tkinter as tk
from tkinter import *
from tkinter import messagebox
import pandas as pd, sys, os
import xlsxwriter
import xlrd
import webbrowser
import time

WorkspacePath = ""
TargetFileName = ""
ResultFileName = ""
NumberOfFormulas = ""



window_main = tk.Tk(className=' Apply Formulas To Group Of CSV Files', )
window_main.geometry("582x420")
window_main.resizable(0,0)


bg_data = tk.Button(window_main, height= 12, width=49)
bg_data.pack()
bg_data.place(relx = 0.005, rely = 0.005)
bg_data.config(state='disabled')

bg_formula = tk.Button(window_main, height= 27, width=30)
bg_formula.pack()
bg_formula.place(relx = 0.615, rely = 0.005)
bg_formula.config(state='disabled')

bg_exe_stat = tk.Button(window_main, height= 10, width=49)
bg_exe_stat.pack()
bg_exe_stat.place(relx = 0.005, rely = 0.47)
bg_exe_stat.config(state='disabled')




label_data_section = tk.Label(window_main, text='REQUIRED DATA:', font=("Calibri",12,"bold"))
label_data_section.place(relx = 0.025, rely = 0.02)

label_formula_section = tk.Label(window_main, text='FORMULAS:', font=("Calibri",12,"bold"))
label_formula_section.place(relx = 0.635, rely = 0.02)

label_formula_section = tk.Label(window_main, text='EXECUTION STATUS:', font=("Calibri",12,"bold"))
label_formula_section.place(relx = 0.025, rely = 0.49)




label_1 = tk.Label(window_main, text='Workspace Path')
label_1.place(relx = 0.025, rely = 0.1)


entry_1 = tk.StringVar()
entry_widget_1 = tk.Entry(window_main, textvariable=entry_1, width=29)
entry_widget_1.pack()
entry_widget_1.place(relx = 0.29, rely = 0.1)



label_2 = tk.Label(window_main, text='Target File Name (.csv)')
label_2.place(relx = 0.025, rely = 0.16)


entry_2 = tk.StringVar()
entry_widget_2 = tk.Entry(window_main, textvariable=entry_2, width=29)
entry_widget_2.pack()
entry_widget_2.place(relx = 0.29, rely = 0.16)




label_3 = tk.Label(window_main, text='Result File Name (.csv)')
label_3.place(relx = 0.025, rely = 0.22)


entry_3 = tk.StringVar()
entry_widget_3 = tk.Entry(window_main, textvariable=entry_3, width=29)
entry_widget_3.pack()
entry_widget_3.place(relx = 0.29, rely = 0.22)



label_4 = tk.Label(window_main, text='Number of Formulas (1-10)')
label_4.place(relx = 0.025, rely = 0.28)


entry_4 = tk.StringVar()
entry_widget_4 = tk.Entry(window_main, textvariable=entry_4, width=29)
entry_widget_4.pack()
entry_widget_4.place(relx = 0.29, rely = 0.28)



label_5 = tk.Label(window_main, text='version 0.1')
label_5.place(relx = 0.025, rely = 0.94)


label_6 = tk.Label(window_main, text='By Nitin D. Patwari')
label_6.place(relx = 0.34, rely = 0.86)


def callback(url):
    webbrowser.open_new_tab(url)
    
label_7 = tk.Label(window_main, text='LinkedIn:')
label_7.place(relx = 0.2, rely = 0.9)
label_8 = tk.Label(window_main, text='nitin-d-patwari-60368796', fg='blue', cursor='hand2')
label_8.place(relx = 0.34, rely = 0.9)
label_8.bind("<Button-1>", lambda e: callback('https://linkedin.com/in/nitin-d-patwari-60368796'))


label_9 = tk.Label(window_main, text='GitHub:')
label_9.place(relx = 0.2, rely = 0.94)
label_10 = tk.Label(window_main, text='patwarind', fg='blue', cursor='hand2')
label_10.place(relx = 0.34, rely = 0.94)
label_10.bind("<Button-1>", lambda e: callback('https://github.com/patwarind'))



def submitValues():
    global WorkspacePath
    global TargetFileName
    global ResultFileName
    global NumberOfFormulas

    WorkspacePath = r"C:\Users\HP\Documents\GitHub\Excel_automation_tools\ApplyFormulaToGroupOfFiles_v0.1\workspace"#entry_1.get()
    TargetFileName = "test_file.csv"#entry_2.get()
    ResultFileName = "result.csv"#entry_3.get() 
    NumberOfFormulas = "1"#entry_4.get()

    try:
        if (0<int(NumberOfFormulas)<11):               
            for k in range(0, int(NumberOfFormulas)):
                globals()['label_formulas_%s' %k] = tk.Label(window_main, text='Formula '+str(k+1)+':')
                globals()['label_formulas_%s' %k].place(relx = 0.635, rely = 0.1+k*0.08)
                globals()['entry_formulas_%s' %k] = tk.StringVar()
                globals()['entry_widget_formulas_%s' %k] = tk.Entry(window_main, textvariable=globals()['entry_formulas_%s' %k])
                globals()['entry_widget_formulas_%s' %k].pack()
                globals()['entry_widget_formulas_%s' %k].place(relx = 0.755, rely = 0.1+k*0.08)

            label_11 = tk.Label(window_main, text='1) REQUIRED DATA FOR PROCESSING CAPTURED!!', fg='green')
            label_11.place(relx = 0.025, rely = 0.56)
            label_12 = tk.Label(window_main, text='2) PLEASE ENTER FORMULA/S', fg='green')
            label_12.place(relx = 0.025, rely = 0.62)
            
            submit.config(state='disabled')
            entry_widget_1.config(state='disabled')
            entry_widget_2.config(state='disabled')
            entry_widget_3.config(state='disabled')
            entry_widget_4.config(state='disabled')
            process.config(state='active')        
        else:
            tk.messagebox.showwarning(title="Enter valid number of formulas", message="Number of formulas must have value in between 1 to 10 including 1 and 10")
            pass        
    except:
        tk.messagebox.showwarning(title="Enter valid data", message="Enter valid values of each field.")
        





def processData():
    

##    try:
    file_count = 0
    workbook = xlsxwriter.Workbook(WorkspacePath+'\\'+'temp.xlsx')
    worksheet = workbook.add_worksheet('Summary1')
    worksheet1 = workbook.add_worksheet('Summary2')
    

    for root, dist_list, files_list in os.walk(WorkspacePath):
        for file_name in files_list:
            if TargetFileName in file_name:
                
                file_name_path = os.path.join(root, file_name)
                print(file_name_path)
                df = pd.read_csv(file_name_path, low_memory=False)
##                print(list(df.columns))

                for column in range(0, len(list(df.columns))):
                    worksheet.write(0, column, list(df.columns)[column])
                    
                col = 0
                for label in list(df.columns):  
                    for row in range(1, len(list(df[label]))+1):
                        worksheet.write(row, col, list(df[label])[row-1])
                    col = col + 1
              

                worksheet1.write_string(1+file_count,0, file_name_path)


                for k in range(0, int(NumberOfFormulas)):
                    if(globals()['entry_formulas_%s' %k].get()==""):
                        raise Exception("Enter valid formulas")
                    print(globals()['entry_formulas_%s' %k].get())
                    worksheet.write_formula('ZZ'+str(k+1), globals()['entry_formulas_%s' %k].get())
                    worksheet.activate()
                    globals()['entry_widget_formulas_%s' %k].config(state='disabled')
                    if(file_count==0): 
                        worksheet1.write_string(0,k+1, 'Formula '+str(k))
                    else:
                        pass
                    
                    worksheet1.write_formula(1+file_count,k+1, 'Summary1!'+'ZZ'+str(k+1))
                    worksheet1.activate()
                    
##                    worksheet.hide()
                    file_count = file_count + 1
                    
    workbook.close()
                
    
    result_data = df
    result_data.to_csv(WorkspacePath+"\\" + ResultFileName, index=False)

    process.config(state='disabled')
    label_13 = tk.Label(window_main, text='3) TASK COMPLETED SUCCESSFULLY', fg='green')
    label_13.place(relx = 0.025, rely = 0.68)
    label_14 = tk.Label(window_main, text='TASK COMPLETED!!', bg='green', fg='white')
    label_14.place(relx = 0.22, rely = 0.78)
    msg_answer=tk.messagebox.askyesno(title="Task status", message="Task completed successfully. Do you want to exit?")
    if (msg_answer):
        window_main.destroy()
    else:
        pass
    pass


##    except:
##        tk.messagebox.showwarning(title="Enter formulas correctly", message="Enter valid formulas.")
                    



    
submit = tk.Button(window_main, text="Submit", command=submitValues)
submit.pack()
submit.place(relx = 0.025, rely = 0.35)


process = tk.Button(window_main, text="Apply", command=processData)
process.pack()
process.place(relx = 0.635, rely = 0.9)
process.config(state='disabled')

 
window_main.mainloop()







