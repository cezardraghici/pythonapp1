
from tkinter import *
import tkinter as tk
import pandas as pd
import bkdb
from xlsxwriter.workbook import Workbook
from tkinter.filedialog import askopenfilename, asksaveasfilename
import tkinter.ttk as ttk 

window = tk.Tk()
window.geometry("1350x600")
window.title("Jhon")
width, height = window.winfo_screenwidth(), window.winfo_screenheight()
window.geometry('%dx%d+0+0' % (width,height))

def view_command():
    clear_command()
    for row in bkdb.view():
        tree.insert("",tk.END,values=row)
        
def delete_command():
    bkdb.delete()

def search_command():
    for row in bkdb.search(nume_text.get(),firma_produs_text.get()):
        tree.insert("",tk.END,values=row)
        clear_text()
        
def insert_command():
    bkdb.insert(nume_text.get(),firma_produs_text.get(), unitate_text.get(), nr_bucati_text.get(), pret_text.get())
    tree.insert("",tk.END,values=(nume_text.get(),firma_produs_text.get(), unitate_text.get(), nr_bucati_text.get(), pret_text.get()))
    clear_text()

def clear_command():
    clear_text()
    for i in tree.get_children():
        tree.delete(i)
    
def insertExel_command():
    bkdb.insert_from_excel()
   
def save_command():
    savefile = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") )) 
    wb = Workbook(savefile+ ".xlsx",{'strings_to_numbers': True})
    ws = wb.add_worksheet()
    bold = wb.add_format({'bold': True})
    number_format = wb.add_format({'num_format': '#,##0.00'})
    header = ('Nume produs','Firma Produs','Um','Nr. buc','Pret','Pret cu TVA','Total')
    i = 0
    j = 0
    for vl in (header):
        ws.write(i,j,vl,bold)
        j+=1 
    i=1
    j=0
    for child in tree.get_children():
        ws.write_row(i,0,tree.item(child)["values"])
        ws.write(i,5,'=E'+str(i+1)+'*1.19')
        ws.write(i,6,'=D'+str(i+1)+'*F'+str(i+1))
        i+=1 
    
    ws.set_column('E:G', cell_format=number_format)
    ws.set_column('A:A', width=70)
    ws.set_column('B:B', width=15)
    ws.set_column('C:Z', width=10)
    ws.write(i, 0, 'Total',bold)
    ws.set_column('H:H', cell_format=number_format)
    ws.write(i, 7, '=SUM($G:$G)',bold)
    ws.write(i, 8, 'Lei',bold)
    wb.close()
  
def remove_item_command():
   selected_items = tree.selection()        
   for selected_item in selected_items:          
      tree.delete(selected_item)
     
def clear_text():
    e1.delete(0,'end')
    e2.delete(0,'end')
    e3.delete(0,'end')
    e4.delete(0,'end')
    e5.delete(0,'end')

menu = Menu(window)
window.config(menu=menu)
filemenu = Menu(menu)
menu.add_cascade(label="File", menu=filemenu)
filemenu.add_command(label="Open...", command=insertExel_command)
filemenu.add_command(label="Save", command=save_command)
filemenu.add_command(label="All Products", command=view_command)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=window.quit)

helpmenu = Menu(menu)
menu.add_cascade(label="Help", menu=helpmenu)
helpmenu.add_command(label="Suna-l pe Cezarica")

tree= ttk.Treeview(window, height=30, column=("column1", "column2", "column3","column4","column5",), show='headings')
tree.pack(expand=True, fill='both')
style = ttk.Style()
style.configure("Treeview.Heading", font=(None, 13))

tree.heading("#1", text="Nume Produs")
tree.column("#1", minwidth=0, width=500, stretch=True)
tree.heading("#2", text="Firma Produs")
tree.column("#2", minwidth=0, width=250, stretch=True, anchor=tk.CENTER)
tree.heading("#3", text="Um")
tree.column("#3", minwidth=0, width=110, stretch=True, anchor=tk.CENTER)
tree.heading("#4", text="Nr. buc")
tree.column("#4", minwidth=0, width=110, stretch=True, anchor=tk.CENTER)
tree.heading("#5", text="Pret")
tree.column("#5", minwidth=0, width=110, stretch=True, anchor=tk.CENTER)
tree.grid(row=2,column=0,rowspan=40,columnspan=10)

l1=tk.Label(window,text="Nume Produs ")
l1.grid(row=0, column=0)

l2=tk.Label(window,text="Firma Produs")
l2.grid(row=0,column=2)

l4=tk.Label(window,text="   Unitate  ")
l4.grid(row=0,column=4)

l5=tk.Label(window,text=" Nr. bucati ")
l5.grid(row=0,column=6)

l6=tk.Label(window,text="   Pret    ")
l6.grid(row=0,column=8)

nume_text=StringVar()
e1=ttk.Entry(window,textvariable=nume_text)
e1.grid(row=0,column=1)

firma_produs_text=StringVar()
e2=ttk.Entry(window,textvariable=firma_produs_text)
e2.grid(row=0,column=3)

unitate_text=StringVar()
e3=ttk.Entry(window,textvariable=unitate_text)
e3.grid(row=0,column=5)

nr_bucati_text=IntVar()
e4=ttk.Entry(window,textvariable=nr_bucati_text)
e4.grid(row=0,column=7)

pret_text=DoubleVar()
e5=ttk.Entry(window,textvariable=pret_text)
e5.grid(row=0,column=9)

sb1=tk.Scrollbar(window,orient="vertical",command=tree.yview)
sb1.grid(row=2,column=12,rowspan=40,sticky=tk.NS)

b2=tk.Button(window,text="Find",width=15, command=search_command)
b2.grid(row=4,column=15)

b7=tk.Button(window,text="Clear Screen", command=clear_command,width=15)
b7.grid(row=8,column=15)

b8=tk.Button(window,text="Delete", command=remove_item_command,width=15)
b8.grid(row=6,column=15)

b10=tk.Button(window,text="Insert", command=insert_command,width=15)
b10.grid(row=10,column=15)

window.mainloop()