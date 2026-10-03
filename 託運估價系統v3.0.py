# -*- coding: utf-8 -*-
"""
Created on Wed Aug  7 20:40:03 2024

@author: 陳皓均
"""

import tkinter as tk
import pandas as pd

def on_submit():
    try:
        df = pd.read_excel(entry.get() + ".xlsx")
    except :
        label2.config(text="找不到檔案")
        return
    df1 = pd.read_excel(entry.get() + ".xlsx", sheet_name="注意事項")

    root.destroy()
    
    car = tk.Tk()
    car.geometry("850x600")
    car.resizable(False, False) 
    car.title("車資估算表")

    car_left = tk.Frame(car, width=100, height=400)
    car_right = tk.Frame(car, width=100, height=400)
    car_bottom = tk.Frame(car, width=100, height=400)
    
    car_left.grid(row=0, column=0, sticky="NW")
    car_right.grid(row=0, column=1, sticky="NW")
    car_bottom.grid(row=0, column=2, sticky="NW")
    
    areas = df.columns[1:9]
    
    area_title=tk.Label(car_left,text="請選擇區域",font=("TkDefaultFont",16))
    area_title.grid(sticky="nw")
    
    area_var = tk.StringVar()
    area_var.set("0")
    
    for i in range(len(areas)):
        area= tk.Radiobutton(car_left,text=areas[i],value=i,font=("TkDefaultFont",12),variable=area_var)  
        area.grid(sticky="W")  
        
    blank_title = tk.Label(car_left,text="",font=("TkDefaultFont",16))
    blank_title.grid(sticky="SW")
    
    area_title2 = tk.Label(car_left,text="中南部",font=("TkDefaultFont",16))
    area_title2.grid(sticky="SW")
    
    areas2 = df.columns[10:12]
    
    area_var2 = tk.StringVar()
    area_var2.set("0")
    
    for i in range(len(areas2)):
        area2= tk.Radiobutton(car_left,text=areas2[i],value=i,font=("TkDefaultFont",12),variable=area_var2)  
        area2.grid(sticky="W")  
    
    blank_title1 = tk.Label(car_left,text="",font=("TkDefaultFont",16))
    blank_title1.grid(sticky="SW")
    
    area_title3 = tk.Label(car_left,text="請輸入重量(KG)",font=("TkDefaultFont",16))
    area_title3.grid(sticky="SW")
    
    area_kg=tk.Entry(car_left)
    area_kg.grid(sticky="W")  
    
    muchs = df.iloc[:,0]
    
    much_title=tk.Label(car_right,text="請選擇重量或板數",font=("TkDefaultFont",16))
    much_title.grid(sticky="W")
    much_var = tk.StringVar()  
    much_var.set("0")
    
    for i in range(len(muchs)):
        
        much= tk.Radiobutton(car_right,text=muchs[i],value=i,font=("TkDefaultFont",12),variable=much_var)  
        much.grid(sticky="w")
    
    manual=tk.Label(car_bottom,text="其他說明",font=("TkDefaultFont",16))
    manual.grid(sticky="W")
    
    ps=df1["注意事項"]
    
    for i in range(len(ps)):
        if i == 9 or i == 10:
            much= tk.Label(car_bottom,text=ps[i],font=("TkDefaultFont",12),fg="#D26900")  
            much.grid(sticky="w")
        elif i == 11 or i == 12:
            much= tk.Label(car_bottom,text=ps[i],font=("TkDefaultFont",12),fg="#AE0000")  
            much.grid(sticky="w")
        else:
            much= tk.Label(car_bottom,text=ps[i],font=("TkDefaultFont",12))  
            much.grid(sticky="w")
    
    maneys=tk.Label(car_bottom,text="估算表：",font=("TkDefaultFont",16))
    maneys.grid(sticky="W")
    
    maneys1=[]
    
    for j in range(1,9):
        maneys1.append(df.columns[j])
        
    # maneys1 = df.columns[1]
    
    # maneys2 = df.columns[2]
        
    # maneys3 = df.columns[3]
        
    # maneys4 = df.columns[4]
        
    # maneys5 = df.columns[5]
    
    # maneys6 = df.columns[6]
    
    # maneys7 = df.columns[7]
    
    # maneys8 = df.columns[8]

    def submitOpen():
        submit.config(state=tk.DISABLED)
        submit3.config(state=tk.DISABLED)
        submit4.config(state=tk.DISABLED)
        
        much = much_var.get()
        area = area_var.get()
        global money

        for i in range(8):
            if area == str(i):
                area1=areas[int(area)]
                much1=muchs[int(much)]
                maney1=df[maneys1[i]][int(much)]
        
        # if area == "0" :
        #     area1=areas[int(area)]
        #     much1=muchs[int(much)]
        #     maney1=df[maneys1][int(much)]
        # elif area == "1" :
        #     area1=areas[int(area)]
        #     much1=muchs[int(much)]
        #     maney1=df[maneys2][int(much)]
        # elif area == "2" :
        #     area1=areas[int(area)]
        #     much1=muchs[int(much)]
        #     maney1=df[maneys3][int(much)]
        # elif area == "3"  :
        #     area1=areas[int(area)]
        #     much1=muchs[int(much)]
        #     maney1=df[maneys4][int(much)]
        # elif area == "4"  :
        #     area1=areas[int(area)]
        #     much1=muchs[int(much)]
        #     maney1=df[maneys5][int(much)]
        # elif area == "5"  :
        #     area1=areas[int(area)]
        #     much1=muchs[int(much)]
        #     maney1=df[maneys6][int(much)]
        # elif area == "6"  :
        #     area1=areas[int(area)]
        #     much1=muchs[int(much)]
        #     maney1=df[maneys7][int(much)]
        # elif area == "7"  :
        #     area1=areas[int(area)]
        #     much1=muchs[int(much)]
        #     maney1=df[maneys8][int(much)]
        # else:
        #     money = tk.Label(car_bottom, text="輸入錯誤")
        #     money.grid(row=2, column=0, sticky="SNWE")
    
        money = tk.Label(car_bottom,text=f"\n區域：{area1}\n\n重量/板數：{much1}\n\n預估金額：{maney1}元",font=("TkDefaultFont",18))
        money.grid(row=16, column=0, sticky="W")

        
    def submitOpen2():
        global money
        money.destroy()
        submit.config(state=tk.NORMAL)
        submit3.config(state=tk.NORMAL)
        submit4.config(state=tk.NORMAL)
    
    def submitOpen3():
        global area4,area3

        if area_kg.get().isdigit() == False:
            submit3.config(state=tk.DISABLED)
            submit.config(state=tk.DISABLED)
            submit2.config(state=tk.DISABLED)
            text="請輸入數字"
        else:
            submit3.config(state=tk.DISABLED)
            submit.config(state=tk.DISABLED)
            submit2.config(state=tk.DISABLED)
            
            area_kg_g=int(area_kg.get())
            area4=area_var2.get()
            
            if area4 == "0" :
                if area_kg_g >= (df[df.columns[9]][6]-1)*2:
                    price = "重量太大無法計算"
                else:
                    if area_kg_g >= df[df.columns[9]][6]:
                        area_kg_g2 = area_kg_g - df[df.columns[9]][6]+1
                        if area_kg_g2 >= df[df.columns[9]][5]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[10]][5])+int(area_kg_g2*df[df.columns[10]][5])
                        elif area_kg_g2 >= df[df.columns[9]][4]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[10]][5])+int(area_kg_g2*df[df.columns[10]][4])
                        elif area_kg_g2 >= df[df.columns[9]][3]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[10]][5])+int(area_kg_g2*df[df.columns[10]][3])        
                        elif area_kg_g2 >= df[df.columns[9]][2]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[10]][5])+int(area_kg_g2*df[df.columns[10]][2])
                        elif area_kg_g2 >= df[df.columns[9]][1]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[10]][5])+1200
                        else:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[10]][5])+1100
                    elif area_kg_g >= df[df.columns[9]][5]:
                        price=int(area_kg_g*df[df.columns[10]][5])
                    elif area_kg_g >= df[df.columns[9]][4]:
                        price=int(area_kg_g*df[df.columns[10]][4])
                    elif area_kg_g >= df[df.columns[9]][3]:
                        price=int(area_kg_g*df[df.columns[10]][3])        
                    elif area_kg_g >= df[df.columns[9]][2]:
                        price=int(area_kg_g*df[df.columns[10]][2])
                    elif area_kg_g >= df[df.columns[9]][1]:
                        price=1200 
                    else:
                        price=1100
            else:
                if area_kg_g >= (df[df.columns[9]][6]-1)*2:
                    price = "重量太大無法計算"
                else:
                    if area_kg_g >= df[df.columns[9]][6]:
                        area_kg_g2 = area_kg_g - df[df.columns[9]][6]+1
                        if area_kg_g2 >= df[df.columns[9]][5]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[11]][5])+int(area_kg_g2*df[df.columns[11]][5])
                        elif area_kg_g2 >= df[df.columns[9]][4]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[11]][5])+int(area_kg_g2*df[df.columns[11]][4])
                        elif area_kg_g2 >= df[df.columns[9]][3]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[11]][5])+int(area_kg_g2*df[df.columns[11]][3])        
                        elif area_kg_g2 >= df[df.columns[9]][2]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[11]][5])+int(area_kg_g2*df[df.columns[11]][2])
                        elif area_kg_g2 >= df[df.columns[9]][1]:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[11]][5])+1600
                        else:
                            price=int((df[df.columns[9]][6]-1)*df[df.columns[11]][5])+1400
                    elif area_kg_g >= df[df.columns[9]][5]:
                        price=int(area_kg_g*df[df.columns[11]][5])
                    elif area_kg_g >= df[df.columns[9]][4]:
                        price=int(area_kg_g*df[df.columns[11]][4])
                    elif area_kg_g >= df[df.columns[9]][3]:
                        price=int(area_kg_g*df[df.columns[11]][3])        
                    elif area_kg_g >= df[df.columns[9]][2]:
                        price=int(area_kg_g*df[df.columns[11]][2])
                    elif area_kg_g >= df[df.columns[9]][1]:
                        price=1600
                    else:
                        price=1400
                        
            if area_kg_g >= (df[df.columns[9]][6]-1)*2:
                text= price
            elif area_kg_g >= df[df.columns[9]][6]:
                text=f"\n區域：{areas2[int(area4)]}\n重量：{area_kg_g} KG\n\n超過最大{df[df.columns[9]][6]}重量\n已兩台車計算價錢\n\n預估金額：{price}元"
            else:
                text=f"\n區域：{areas2[int(area4)]}\n\n重量：{area_kg_g} KG\n\n預估金額：{price}元"
        area3= tk.Label(car_bottom,text=text,font=("TkDefaultFont",18))
        area3.grid(row=15, column=0, sticky="W")  
        
        
    def submitOpen4():
        global area3
        area3.destroy()
        area_kg.delete(0, tk.END)
        submit3.config(state=tk.NORMAL)
        submit.config(state=tk.NORMAL)
        submit2.config(state=tk.NORMAL)
    
    submit3 = tk.Button(car_left,text="計算",command=submitOpen3)
    submit3.grid(sticky="SW")
    
    submit4 = tk.Button(car_left,text="清除",command=submitOpen4)
    submit4.grid(row=16, column=0)    
        
    submit = tk.Button(car_right,text="計算",command=submitOpen)
    submit.grid(sticky="SW")
    
    submit2 = tk.Button(car_right,text="清除",command=submitOpen2)
    submit2.grid(row=8,column=0)
    
    car.mainloop()
    
root = tk.Tk()
root.title("讀取 Excel 檔案")

root.geometry("320x50")
root.resizable(False, False)

label = tk.Label(root, text="請輸入要讀取的Excel名稱:")
label.grid(row=0, column=0)

labelblank = tk.Label(root, text="    ")
labelblank.grid(row=0, column=1)

entry = tk.Entry(root,width=10)
entry.grid(row=0, column=2)

labelblank2 = tk.Label(root, text="        ")
labelblank2.grid(row=0, column=3)

submit_button = tk.Button(root, text="確認", command=on_submit)
submit_button.grid(row=0, column=4)

root.bind('<Return>', lambda event: on_submit())

label2 = tk.Label(root, text="")
label2.grid(sticky='w')

root.mainloop()