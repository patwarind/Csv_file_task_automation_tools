"""


"""












import tkinter as tk

from tkinter import *

from tkinter import messagebox

import pandas as pd, sys, os

import xlsxwriter

import webbrowser







def callback(url):

    webbrowser.open_new_tab(url)

    pass









def submitValues():

    global WorkspacePath

    global TargetFileName

    global ResultFileName

    global NumberOfFormulas



    WorkspacePath = entry_widget_WorkspacePath.get()

    TargetFileName = entry_widget_TargetFileName.get()

    ResultFileName = entry_widget_ResultFileName.get() 

    NumberOfFormulas = entry_widget_NumberOfFormulas.get()




    try:

        if (0<int(NumberOfFormulas)<11):               

            for k in range(0, int(NumberOfFormulas)):

                globals()['label_formulas_%s' %k] = tk.Label(window_main, text='Formula '+str(k+1)+':')

                globals()['label_formulas_%s' %k].place(relx = 0.635, rely = 0.1+k*0.08)


                globals()['entry_formulas_%s' %k] = tk.StringVar()

                globals()['entry_widget_formulas_%s' %k] = tk.Entry(window_main, textvariable=globals()['entry_formulas_%s' %k])

                globals()['entry_widget_formulas_%s' %k].pack()

                globals()['entry_widget_formulas_%s' %k].place(relx = 0.755, rely = 0.1+k*0.08)




            label_status_1 = tk.Label(window_main, text='1) REQUIRED DATA FOR PROCESSING CAPTURED!!', fg='green')

            label_status_1.place(relx = 0.025, rely = 0.56)


            label_status_2 = tk.Label(window_main, text='2) PLEASE ENTER FORMULA/S', fg='green')

            label_status_2.place(relx = 0.025, rely = 0.62)
            


            submit.config(state='disabled')

            entry_widget_WorkspacePath.config(state='disabled')

            entry_widget_TargetFileName.config(state='disabled')

            entry_widget_ResultFileName.config(state='disabled')

            entry_widget_NumberOfFormulas.config(state='disabled')



            apply.config(state='active')        



        else:

            tk.messagebox.showwarning(title="Enter valid number of formulas", message="Number of formulas must have value in between 1 to 10 including 1 and 10")

            pass        



    except:

        tk.messagebox.showwarning(title="Enter valid data", message="Enter valid values of each field.")



        








def applyFormulas():
    
    try:

        file_count = 0

        workbook = xlsxwriter.Workbook(WorkspacePath+'\\'+ResultFileName)

        worksheet_summary = workbook.add_worksheet('Summary')


        

        for root, dist_list, files_list in os.walk(WorkspacePath):

            for file_name in files_list:

                if TargetFileName in file_name:


                    globals()['worksheet%s' %file_count] = workbook.add_worksheet('file_'+str(file_count))

                    file_name_path = os.path.join(root, file_name)

                    df = pd.read_csv(file_name_path, low_memory=False)


                    for column in range(0, len(list(df.columns))):

                        globals()['worksheet%s' %file_count].write(0, column, list(df.columns)[column])

                        


                    col = 0


                    for label in list(df.columns):  

                        for row in range(1, len(list(df[label]))+1):

                            globals()['worksheet%s' %file_count].write(row, col, list(df[label])[row-1])

                        col = col + 1
                  


                    
                    worksheet_summary.write_string(1+file_count,0, file_name_path)




                    for k in range(0, int(NumberOfFormulas)):

                        if(globals()['entry_formulas_%s' %k].get()==""):

                            raise Exception("Enter valid formulas")


                        globals()['worksheet%s' %file_count].write_formula('ZZ'+str(k+1), globals()['entry_formulas_%s' %k].get())

                        globals()['worksheet%s' %file_count].activate()

                        globals()['entry_widget_formulas_%s' %k].config(state='disabled')


                        if(file_count==0): 

                            worksheet_summary.write_string(0,k+1, 'Formula '+str(k+1))

                        else:

                            pass


                        
                        worksheet_summary.write_formula(1+file_count,k+1, 'file_'+str(file_count)+'!'+'ZZ'+str(k+1))

                        worksheet_summary.activate()
                        
                        globals()['worksheet%s' %file_count].hide()





                    file_count = file_count + 1
                        



        workbook.close()
                    
        



        apply.config(state='disabled')



        label_status_3 = tk.Label(window_main, text='3) TASK COMPLETED SUCCESSFULLY', fg='green')

        label_status_3.place(relx = 0.025, rely = 0.68)


        label_status_final = tk.Label(window_main, text='TASK COMPLETED!!', bg='green', fg='white')

        label_status_final.place(relx = 0.22, rely = 0.78)



        msg_answer=tk.messagebox.askyesno(title="Task status", message="Task completed successfully. Do you want to exit?")



        if (msg_answer):

            window_main.destroy()

        else:

            pass

        

    except:

        tk.messagebox.showwarning(title="Enter formulas correctly", message="Enter valid formulas.")
                    


















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






label_WorkspacePath = tk.Label(window_main, text='Workspace Path')

label_WorkspacePath.place(relx = 0.025, rely = 0.1)




entry_WorkspacePath = tk.StringVar()

entry_widget_WorkspacePath = tk.Entry(window_main, textvariable=entry_WorkspacePath, width=29)

entry_widget_WorkspacePath.pack()

entry_widget_WorkspacePath.place(relx = 0.29, rely = 0.1)





label_TargetFileName = tk.Label(window_main, text='Target File Name (.csv)')

label_TargetFileName.place(relx = 0.025, rely = 0.16)





entry_TargetFileName = tk.StringVar()

entry_widget_TargetFileName = tk.Entry(window_main, textvariable=entry_TargetFileName, width=29)

entry_widget_TargetFileName.pack()

entry_widget_TargetFileName.place(relx = 0.29, rely = 0.16)







label_ResultFileName = tk.Label(window_main, text='Result File Name (.xlsx)')

label_ResultFileName.place(relx = 0.025, rely = 0.22)





entry_ResultFileName = tk.StringVar()

entry_widget_ResultFileName = tk.Entry(window_main, textvariable=entry_ResultFileName, width=29)

entry_widget_ResultFileName.pack()

entry_widget_ResultFileName.place(relx = 0.29, rely = 0.22)






label_NumberOfFormulas = tk.Label(window_main, text='Number of Formulas (1-10)')

label_NumberOfFormulas.place(relx = 0.025, rely = 0.28)




entry_NumberOfFormulas = tk.StringVar()

entry_widget_NumberOfFormulas = tk.Entry(window_main, textvariable=entry_NumberOfFormulas, width=29)

entry_widget_NumberOfFormulas.pack()

entry_widget_NumberOfFormulas.place(relx = 0.29, rely = 0.28)





label_version = tk.Label(window_main, text='version 0.1')

label_version.place(relx = 0.025, rely = 0.94)





label_name = tk.Label(window_main, text='By Nitin D. Patwari')

label_name.place(relx = 0.34, rely = 0.86)






    
label_LinkedIn = tk.Label(window_main, text='LinkedIn:')

label_LinkedIn.place(relx = 0.24, rely = 0.9)





label_link_LinkedIn = tk.Label(window_main, text='nitin-d-patwari-60368796', fg='blue', cursor='hand2')

label_link_LinkedIn.place(relx = 0.34, rely = 0.9)

label_link_LinkedIn.bind("<Button-1>", lambda e: callback('https://linkedin.com/in/nitin-d-patwari-60368796'))






label_GitHub = tk.Label(window_main, text='GitHub:')

label_GitHub.place(relx = 0.24, rely = 0.94)




label_link_GitHub = tk.Label(window_main, text='patwarind', fg='blue', cursor='hand2')

label_link_GitHub.place(relx = 0.34, rely = 0.94)

label_link_GitHub.bind("<Button-1>", lambda e: callback('https://github.com/patwarind'))





    
submit = tk.Button(window_main, text="Submit", command=submitValues)

submit.pack()

submit.place(relx = 0.025, rely = 0.35)





apply = tk.Button(window_main, text="Apply", command=applyFormulas)

apply.pack()

apply.place(relx = 0.635, rely = 0.9)

apply.config(state='disabled')

 




window_main.mainloop()







