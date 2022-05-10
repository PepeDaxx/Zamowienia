import string
import tkinter as tk
from tkinter import messagebox,ttk
from tkinter import *
from tkinter.filedialog import askopenfilename, asksaveasfilename

from requests import delete

from Comparator import Comparator
from Orders import OrderReader
#Variables
version = "0.3"
file1_path = ""
file2_path = ""
order_file_path = ""
differences = {}
new_items = {}
deleted_items = {}
buy_differences = {}
order_products = {}
comp = Comparator()
order_reader = OrderReader()
loaded = False
ord_loaded = False
files = [("Arkusz kalkulacyjny","*.xlsx")]
# Table sorting function
def treeview_sort(tv,col,reverse,type):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    if type == string:
        l.sort(reverse=reverse)
    elif type == int:
        l.sort(key=lambda t: int(t[0]),reverse=reverse)
    elif type == float:
        l.sort(key=lambda t: float(t[0]),reverse=reverse)
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)
    tv.heading(col, text=col, command=lambda _col=col: treeview_sort(tv, _col, not reverse,type))
#Choose first file (old)
def choose_file1():
    global file1_path
    file1_path = askopenfilename(filetypes=files)
    file_name = file1_path[file1_path.rfind("/")+1:]
    file1_label.config(text=file_name)
#Choose second file (new)
def choose_file2():
    global file2_path
    file2_path = askopenfilename(filetypes=files)
    file_name = file2_path[file2_path.rfind("/")+1:]
    file2_label.config(text=file_name)
#Choose file contaning sell data, current stock for checking sale possibilities
def choose_order_file():
    global order_file_path
    order_file_path = askopenfilename(filetypes=files)
    file_name = order_file_path[order_file_path.rfind("/")+1:]
    order_file_label.config(text=file_name)
#Main function for both functionalities. First sends item paths to Comparator.py which extracts, compares and delivers rresults about old and dew but prices
#For second function application takes given file path and sends it to Orders.py to calculate sales and how long it will take to sell out given products
def generate_output(notebook = int, **args):
    #Notebook 1 - Buy differences
    if notebook == 1:
        global differences, new_items, deleted_items, loaded
        if file1_path == "" or file2_path == "":
            messagebox.showerror("Nie moge rozpocząć pracy","Musisz wskazać obydwa cenniki!")
        else:
            differences, new_items, deleted_items, buy_differences = comp.read_files(file1_path,file2_path)
            put_data_into_table(differences,new_items,deleted_items,buy_differences)
            messagebox.showinfo("Sukces","Operacja wykonana pomyślnie!")
            loaded = True
    #Notebook 2 - Sales
    elif notebook == 2:
        global order_products
        if order_file_path =="":
            messagebox.showerror("Nie moge rozpocząć pracy","Musisz wskazać plik!")
        else:
            order_products = order_reader.read_file(order_file_path)
            update_orders_table(order_products)
            messagebox.showinfo("Sukces","Operacja wykonana pomyślnie!")
#Fill tables in notebook 1 
def put_data_into_table(difference,new,deleted,buy_diff):
    difference_table.delete(*difference_table.get_children())
    new_table.delete(*new_table.get_children())
    deleted_table.delete(*deleted_table.get_children())
    buy_diff_table.delete(*buy_diff_table.get_children())
    id = 0
    for key in difference:
        list = difference[key]
        difference_table.insert(parent="",index="end",iid=id,text="",values=(id+1,key,list[0],list[1],list[2],list[3]))
        id += 1
    id = 0
    for key in new:
        list = new[key]
        new_table.insert(parent="",index="end",iid=id,text="",values=(id+1,key,list[0],list[1],list[2],list[3]))
        id += 1
    id = 0
    for key in deleted:
        list = deleted[key]
        deleted_table.insert(parent="",index="end",iid=id,text="",values=(id+1,key,list[0]))
        id += 1
    id = 0
    for key in buy_diff:
        buy_list = buy_diff[key]
        buy_diff_table.insert(parent="",index="end",iid=id,text="",values=(id+1,key,buy_list[0],buy_list[1],buy_list[2],buy_list[3],buy_list[4]))
        id += 1
#Fill tables in notebook 2
def update_orders_table(order_products_dict):
    global ord_loaded
    order_table.delete(*order_table.get_children())
    id = 0
    for key in order_products_dict:
        list = order_products_dict[key]
        order_table.insert(parent="",index="end",iid=id,text="",values=(id+1,key,list[0],list[1],list[2],list[3],list[4],list[5],list[6]))
        id +=1
    ord_loaded = True
#Save file with notebook 1 results
def export_to_file():
    if loaded == False:
        messagebox.showerror("Błąd","Brak wyników do zapisania!")
    else:
        file_path = asksaveasfilename(filetypes=files,defaultextension=files)
        comp.write_file(file_path)
        messagebox.showinfo("Sukces","Plik został pomyślnie zapisany!")
#Save file with notebook 2 results
def export_orders():
    if ord_loaded == False:
        messagebox.showerror("Błąd","Brak wyników do zapisania!")
    else:
        file_path = asksaveasfilename(filetypes=files,defaultextension=files)
        order_reader.write_file(file_path)
        messagebox.showinfo("Sukces","Plik został pomyślnie zapisany!")
def quit_program():
    win.quit()
    quit()    
#GUI
win = tk.Tk()
win.title("Podręcznik Tomcia")
win.geometry("1800x900")
win.resizable(False,False)
#Notebooks 
notebook = ttk.Notebook(win)
notebook.pack()
n1_frame = ttk.Frame(notebook)
n1_frame.pack()
notebook.add(n1_frame,text="Generator")
n2_frame = ttk.Frame(notebook)
n2_frame.pack()
notebook.add(n2_frame,text="Zamówienia")
version_label = ttk.Label(win,text=f"Wersja {version}")
version_label.pack()
### NOTEBOOK 1 - n1_frame
#Gen Frame
gen_frame = ttk.LabelFrame(n1_frame,text="Generator")
gen_frame.pack(expand=True,fill=BOTH)
file1_label = ttk.Label(gen_frame,text="...",width=50)
file1_label.grid(column=0,row=0,padx=0)

file1_button = ttk.Button(gen_frame,text="Wybierz stary cennik",command=choose_file1,width=20)
file1_button.grid(column=1,row=0)

file2_label = ttk.Label(gen_frame,text="...",width=50)
file2_label.grid(column=0,row=1)

file2_button = ttk.Button(gen_frame,text="Wybierz nowy cennik",command=choose_file2,width=20)
file2_button.grid(column=1,row=1)

generate_button = ttk.Button(gen_frame,text="Porównaj",command=lambda: generate_output(notebook = 1))
generate_button.grid(column=0,row=2)

save_button=ttk.Button(gen_frame,text = "Zapisz wyniki",command=export_to_file)
save_button.grid(column=1,row=2)
first_line_frame = Frame(n1_frame)
first_line_frame.pack(fill=BOTH)
second_line_frame = Frame(n1_frame)
second_line_frame.pack(fill=BOTH)
third_line_frame = Frame(n1_frame)
third_line_frame.pack(fill=BOTH)
#Diff, new and deleted frame
diff_frame = ttk.LabelFrame(first_line_frame,text="Różnice w RRP")
diff_frame.pack(expand=True,fill=BOTH)
buy_diff_frame = ttk.LabelFrame(second_line_frame,text="Różnica w zakupie")
buy_diff_frame.pack(expand=True,fill=BOTH)
new_frame = ttk.LabelFrame(third_line_frame,text="Nowe produkty")
new_frame.pack(expand=True,fill=BOTH,side=LEFT)
deleted_frame = ttk.LabelFrame(third_line_frame,text="Usunięte produkty")
deleted_frame.pack(expand=True,fill=BOTH,side=RIGHT)
#Create Difference table
diff_columns = ["LP","PN","Nazwa Produktu","RRP","Stare RRP","Różnica"]
difference_table = ttk.Treeview(diff_frame,columns=diff_columns,show='headings')
for col in diff_columns:
    if col == "LP":
        difference_table.heading(col,text=col,command=lambda c = col : treeview_sort(difference_table,c,False,int))
    elif col == "RRP" or col =="Stare RRP" or col == "Różnica":
        difference_table.heading(col,text=col,command=lambda c = col : treeview_sort(difference_table,c,False,float))
    else:
        difference_table.heading(col,text=col,command=lambda c = col : treeview_sort(difference_table,c,False,string))
diff_scr = ttk.Scrollbar(diff_frame,orient="vertical",command=difference_table.yview)
diff_scr.pack(side="right",fill=Y)
difference_table.configure(yscrollcommand=diff_scr.set)
difference_table.pack(fill=BOTH)
#Buy price difference table
buy_diff_columns = ["LP","PN","Nazwa","Zakup","Poprzedni Zakup","Różnica","Promocja"]
buy_diff_table = ttk.Treeview(buy_diff_frame,columns=buy_diff_columns,show='headings')
for col in buy_diff_columns:
    if col == "LP":
        buy_diff_table.heading(col,text=col,command=lambda c = col : treeview_sort(buy_diff_table,c,False,int))
    elif col == "Zakup" or col =="Poprzedni Zakup" or col == "Różnica":
        buy_diff_table.heading(col,text=col,command=lambda c = col : treeview_sort(buy_diff_table,c,False,float))
    else:
        buy_diff_table.heading(col,text=col,command=lambda c = col : treeview_sort(buy_diff_table,c,False,string))
buy_diff_scr = ttk.Scrollbar(buy_diff_frame,orient="vertical",command=buy_diff_table.yview)
buy_diff_scr.pack(side="right",fill=Y)
buy_diff_table.configure(yscrollcommand=buy_diff_scr.set)
buy_diff_table.pack(fill=BOTH)
#Create new items table
new_columns = ["LP","PN","Nazwa produktu","RRP","Zakup","Promocja"]
new_table = ttk.Treeview(new_frame,columns=new_columns,show='headings')
for col in new_columns:
    if col == "LP":
        new_table.heading(col,text=col,command=lambda c = col : treeview_sort(new_table,c,False,int))
    elif col == "RRP" or col =="Zakup":
        new_table.heading(col,text=col,command=lambda c = col : treeview_sort(new_table,c,False,float))
    else:
        new_table.heading(col,text=col,command=lambda c = col : treeview_sort(new_table,c,False,string))
new_scr = ttk.Scrollbar(new_frame,orient="vertical",command=new_table.yview)
new_scr.pack(side="right",fill=Y)
new_table.configure(yscrollcommand=new_scr.set)
new_table.pack(fill=BOTH)
#Create deleted items frame
deleted_columns = ["LP","PN","Nazwa Produktu"]
deleted_table = ttk.Treeview(deleted_frame,columns=deleted_columns,show='headings')
for col in deleted_columns:
    if col == "LP":
        deleted_table.heading(col,text=col,command=lambda c = col : treeview_sort(deleted_table,c,False,int))
    else:
        deleted_table.heading(col,text=col,command=lambda c = col : treeview_sort(deleted_table,c,False,string))
new_scr = ttk.Scrollbar(new_frame,orient="vertical",command=new_table.yview)
deleted_scr = ttk.Scrollbar(deleted_frame,orient="vertical",command=deleted_table.yview)
deleted_scr.pack(side="right",fill=Y)
deleted_table.configure(yscrollcommand=deleted_scr.set)
deleted_table.pack(fill=BOTH)
### NOTEBOOK 2 - n2_frame
top2_frame = ttk.LabelFrame(n2_frame,text="Zamówienia")
top2_frame.pack(fill=BOTH)
order_file_label = ttk.Label(top2_frame,text="...",width=50)
order_file_label.grid(column=0,row=0,padx=0)
order_file_button = ttk.Button(top2_frame,text="Wybierz plik",command=choose_order_file,width=20)
order_file_button.grid(column=1,row=0)

order_generate_button = ttk.Button(top2_frame,text="Wczytaj dane",command=lambda: generate_output(notebook=2))
order_generate_button.grid(column=1,row=1)
order_save_button=ttk.Button(top2_frame,text = "Zapisz wyniki",command=export_orders)
order_save_button.grid(column=1,row=2)
# Table for ordering usage
table_n2_frame = ttk.LabelFrame(n2_frame,text="Wyniki")
table_n2_frame.pack(fill=BOTH,expand=True)
order_columns = ('LP','PN',"Nazwa produktu","Stan","Minimum","Różnica","Sprzedaż","Do Min","Do Zera")
order_table = ttk.Treeview(table_n2_frame,columns=order_columns,show='headings')
for col in order_columns:
    if col == "LP" or col =="Stan" or col =="Minimum" or col =="Różnica" or col =="Sprzedaż":
        order_table.heading(col,text=col,command=lambda c = col : treeview_sort(order_table,c,False,int))
    elif col == "Do Min" or col == "Do Zera":
        order_table.heading(col,text=col,command=lambda c = col : treeview_sort(order_table,c,False,float))
    else:
        order_table.heading(col,text=col,command=lambda c = col : treeview_sort(order_table,c,False,string))
order_scr = ttk.Scrollbar(table_n2_frame,orient="vertical",command=order_table.yview)
order_scr.pack(side="right",fill=Y)
order_table.configure(yscrollcommand=order_scr.set)
order_table.pack(fill=BOTH,expand=True)


win.mainloop()