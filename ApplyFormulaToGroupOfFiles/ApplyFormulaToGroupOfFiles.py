import tkinter as tk
import pandas as pd, sys, os
import xlsxwriter

global WorkspacePath
global TargetFileName
global ResultFileName
global Formula



window_main = tk.Tk(className=' Apply Formula To Group Of Files', )
window_main.geometry("400x200")


label_1 = tk.Label(window_main, text='Workspace Path')
label_1.place(relx = 0.04, rely = 0)


entry_1 = tk.StringVar()
entry_widget_1 = tk.Entry(window_main, textvariable=entry_1)
entry_widget_1.pack()
entry_widget_1.place(relx = 0.37, rely = 0)




label_2 = tk.Label(window_main, text='Target File Name (.csv)')
label_2.place(relx = 0.04, rely = 0.15)


entry_2 = tk.StringVar()
entry_widget_2 = tk.Entry(window_main, textvariable=entry_2)
entry_widget_2.pack()
entry_widget_2.place(relx = 0.37, rely = 0.15)




label_3 = tk.Label(window_main, text='Result File Name (.csv)')
label_3.place(relx = 0.04, rely = 0.3)


entry_3 = tk.StringVar()
entry_widget_3 = tk.Entry(window_main, textvariable=entry_3)
entry_widget_3.pack()
entry_widget_3.place(relx = 0.37, rely = 0.3)



label_4 = tk.Label(window_main, text='Formula')
label_4.place(relx = 0.04, rely = 0.45)


entry_4 = tk.StringVar()
entry_widget_4 = tk.Entry(window_main, textvariable=entry_4)
entry_widget_4.pack()
entry_widget_4.place(relx = 0.37, rely = 0.45)




def submitValues():
    WorkspacePath = r"C:\Users\HP\Documents\GitHub\Excel_automation_tools\ApplyFormulaToGroupOfFiles\workspace"#entry_1.get()
    TargetFileName = "test_file.csv"#entry_2.get()
    ResultFileName = "result.csv"#entry_3.get() 
    Formula = "=SUM(B2:B15)"#entry_4.get()

    for root, dist_list, files_list in os.walk(WorkspacePath):
        for file_name in files_list:
            if TargetFileName in file_name:
                file_name_path = os.path.join(root, file_name)
                print(file_name_path)
                df = pd.read_csv(file_name_path, low_memory=False)
##                print(list(df.columns))
                workbook = xlsxwriter.Workbook(WorkspacePath+'\\'+'clone_result_file.xlsx')
                worksheet = workbook.add_worksheet('Summary1')

                for column in range(0, len(list(df.columns))):
                    worksheet.write(0, column, list(df.columns)[column])
                    
                col = 0
                for label in list(df.columns):
                    
                    for row in range(1, len(list(df[label]))+1):
                        worksheet.write(row, col, list(df[label])[row-1])
                    col = col + 1

                worksheet1 = workbook.add_worksheet('Summary2')
                worksheet1.write_formula('K8', Formula)
                worksheet1.activate()
                worksheet.hide()
                workbook.close()
                


    result_data = df
    result_data.to_csv(WorkspacePath+"\\" + ResultFileName, index=False)

    

    
submit = tk.Button(window_main, text="Submit", command=submitValues)
submit.pack()
submit.place(relx = 0.04, rely = 0.60)


 
window_main.mainloop()







