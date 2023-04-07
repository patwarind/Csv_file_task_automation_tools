import tkinter as tk
 
window_main = tk.Tk(className='Tkinter - TutorialKart', )
window_main.geometry("400x200")


label_1 = tk.Label(window_main, text='.csv files path')
label_1.place(relx = 0.04, rely = 0)


entry_1 = tk.StringVar()
entry_widget_1 = tk.Entry(window_main, textvariable=entry_1)
entry_widget_1.pack()
entry_widget_1.place(relx = 0.35, rely = 0)




label_2 = tk.Label(window_main, text='Result file path')
label_2.place(relx = 0.04, rely = 0.15)


entry_2 = tk.StringVar()
entry_widget_2 = tk.Entry(window_main, textvariable=entry_2)
entry_widget_2.pack()
entry_widget_2.place(relx = 0.35, rely = 0.15)




label_3 = tk.Label(window_main, text='Formula')
label_3.place(relx = 0.04, rely = 0.3)


entry_3 = tk.StringVar()
entry_widget_3 = tk.Entry(window_main, textvariable=entry_3)
entry_widget_3.pack()
entry_widget_3.place(relx = 0.35, rely = 0.3)




def submitValues():
    print(entry_1.get())
    print(entry_2.get())
    print(entry_3.get())
 
submit = tk.Button(window_main, text="Submit", command=submitValues)
submit.pack()
submit.place(relx = 0.04, rely = 0.45)
 
window_main.mainloop()
