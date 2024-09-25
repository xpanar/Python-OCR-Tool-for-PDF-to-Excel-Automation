import pdfplumber
import json
import tkinter as tk
from tkinter import * 
from tkinter.filedialog import askopenfilename,asksaveasfilename
import tkinter.ttk as ttk
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill



filedata = []               
line_item = 0               
installation_total = 0.00   
Transport_cost = 0          
Lead_and_Lift = 0           

def get_column_values(treeview, column_index):
    column_values = []                      
    for item_id in treeview.get_children(): 
        item = treeview.item(item_id)       
        values = item['values']             
        print(f"Retrieved values: {values}")  # Debugging
        if column_index < len(values):      
            column_values.append(str(values[column_index]))  # Convert to string explicitly
    return column_values                    

def insert_row_above():
    selected_item = treeview.selection()
    if selected_item:
        item = selected_item[0]
        index = treeview.index(item)
        treeview.insert('', index, text="New Row", values=("", "", "", "", "" ))

def insert_row_below():
    selected_item = treeview.selection()
    if selected_item:
        item = selected_item[0]
        index = treeview.index(item) + 1
        treeview.insert('', index, text="New Row", values=("", "", "", "", "" ))

def delete_row():
    selected_item = treeview.selection()
    if selected_item:
        item = selected_item[0]
        treeview.delete(item)



def on_double_click(event):
    # Identify the row and column
    row_id = treeview.identify_row(event.y)
    column = treeview.identify_column(event.x)
    
    if not row_id or not column:
        return
    
    # Get the bounding box of the cell
    x, y, width, height = treeview.bbox(row_id, column)

    # Get the current value of the cell
    col_index = int(column[1:]) - 1
    current_value = treeview.item(row_id, 'values')[col_index]

    # Create an Entry widget
    entry = ttk.Entry(master)
    entry.place(x=x, y=y, width=width, height=height)
    entry.insert(0, current_value)
    entry.focus()

    def on_enter(event):
        new_value = entry.get()
        values = list(treeview.item(row_id, 'values'))
        values[col_index] = new_value
        treeview.item(row_id, values=values)
        entry.destroy()
        total = 0
        for child in treeview.get_children():
            values = treeview.item(child, 'values')
            total += float(values[5])
        formatted_total = f"{total:.2f}"
        text_insta.delete(0, END)
        text_insta.insert(0, formatted_total)
        

    entry.bind('<Return>', on_enter)
    entry.bind('<FocusOut>', on_enter)


def treeview_to_dict(tree):
    tree_data = []
    for child in tree.get_children():
        values = tree.item(child, 'values')
        row_data = {tree['columns'][i]: values[i] for i in range(len(values))}
        tree_data.append(row_data)
    return tree_data

def file_selector():
    filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
    print(filename)
    #pdf_path = 'C:/Users/xpana/OneDrive/Desktop/big qoutation.pdf'
    extract_text_from_pdf(filename,installation_total)

def file_saver():
    filename = asksaveasfilename(initialfile = "Quote_extracted_values.json",  title="Save Extracted Data", filetypes = (("for send to Google From data","*.json"),("all files","*.*")))

    filedata = treeview_to_dict(treeview)
    filedata.append({"installation_total": text_insta.get() , "Transport_cost": text_Transport_Cost.get(), "Lead_and_Lift": text_Lead_and_Lift.get()})
    print(filedata)
    with open(filename, 'w') as json_file:
        json.dump(filedata, json_file, indent=4)
        print("Write successful")

def extract_text_from_pdf(pdf_path, installation_total):
    all_text = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            all_text += text if text else ""
    line = all_text.split("\n")
    Table = [['0000', '000', '0', '₹0000.00', '₹0000.00']]
    tempT = ["null","null","null","null","null"]

    for i in treeview.get_children():
        treeview.delete(i)

    for i in range(len(line)):
        temp_line = line[i]
        if temp_line.find("Window Code:") >= 0:
            QT = temp_line.split(" ")
            if QT[0] == "Frame":
                continue
            tempT[0] = QT[0]
            line_item =str(QT[0])
            for a in range(len(QT)):
                if QT[a] == "Code:":
                    tempT[1] = QT[a+1]
            #print(tempD)
        if temp_line.find("Qty Rate Discounted Rate Amount") >= 0:
            QT = line[i+1].split(" ")

            tempT[2] = QT[0]

            tempT[3] = QT[len(QT)-4]

            tempT[4] = QT[len(QT)-1]
            

            Table.append(tempT)
            y = {"Sales Line":str(tempT[0]), "Window Code:": tempT[1], "Qty": tempT[2], "Rate": str(tempT[len(tempT)-2]) ,"Amount": str(tempT[len(tempT)-1])[1:]  }
            filedata.append({"windows": y})

            treeview.insert("", "end", values=(str(y["Sales Line"]),y["Window Code:"], y["Qty"],y["Rate"], y["Amount"],0.00))

        if temp_line.find("Transport Cost") >= 0:
            QT = line[i].split(" ")
            Transport_cost = str(QT[len(QT)-1])[1:]
            print("temp_Transport  == " +Transport_cost)
            QT = line[i+1].split(" ")
            Lead_and_Lift = str(QT[len(QT)-1])[1:]
            print("temp_Lead_and_Lift  == "+Lead_and_Lift)

            text_Transport_Cost.config(state=NORMAL)
            text_Transport_Cost.delete(0, END)
            text_Transport_Cost.insert(0, Transport_cost)
            #text_Transport_Cost.config(state=DISABLED)

            text_Lead_and_Lift.config(state=NORMAL)
            text_Lead_and_Lift.delete(0, END)
            text_Lead_and_Lift.insert(0, Lead_and_Lift)
            #text_Lead_and_Lift.config(state=DISABLED)

            break

    print("total line item = "+str(line_item)+"  "+str(int(line_item)))
    line_item = int(line_item)
    installation = [0] * line_item
    
    child = treeview.get_children()
    sale_line = get_column_values(treeview, 0)
    sale_line_lenth = len(sale_line)


    print(sale_line_lenth)
    print(f"Values in 'ID' column: {sale_line}")

    for i in range(len(line),0,-1):
        temp_line = line[i-1]
        print(str("00"+sale_line[sale_line_lenth-1]))
        if temp_line.find(str("00"+str(sale_line[sale_line_lenth-1]))) >= 0:
            print(line[i-1])

            


            temp_line_value = line[i-1].split(" ")

            installation[line_item-1] = temp_line_value[len(temp_line_value)-1]

            print(installation)

            print(f"child: {child}")
            print(f"Length of child: {len(child)}")

            print(treeview.item(child[sale_line_lenth-1], 'values'))
            values = treeview.item(child[sale_line_lenth-1], 'values')
            treeview.item(child[sale_line_lenth-1], values=(values[0], values[1], values[2], values[3],values[4], temp_line_value[len(temp_line_value)-1] ))
            installation_total = float(temp_line_value[len(temp_line_value)-1]) + installation_total
            line_item = line_item - 1
            sale_line_lenth = sale_line_lenth - 1
            if sale_line_lenth <= 0:
                
                break

    text_insta.config(state=NORMAL)
    text_insta.delete(0, END)
    text_insta.insert(0, "{:.2f}".format(installation_total))
    #text_insta.config(state=DISABLED)      

    # Convert Treeview data to a dictionary



        


master = Tk()
master.geometry("640x400")
master.resizable(0, 0)
master.title("ORwing Quotation Data Extractor")

# logo = tk.PhotoImage(file="C:/Users/xpana/OneDrive/Desktop/Python_ws/Dai-cut.png")  # Replace with your image file path
# master.iconphoto(False, logo)

treeview = ttk.Treeview(master, show="headings", columns=("Sales Line", "Window Code:", "Qty", "Rate", "Amount", "Installation"))
v_scroll = ttk.Scrollbar(master, orient='vertical', command=treeview.yview)
treeview.configure(yscroll=v_scroll.set)

treeview.bind('<Double-1>', on_double_click)
treeview.pack(fill='both', expand=True)


treeview.heading("#1", text="Sales Line")
treeview.column("#1", width=70, anchor=CENTER)
treeview.heading("#2", text="Window Code:")
treeview.column("#2", width=100, anchor=CENTER)
treeview.heading("#3", text="Qty")
treeview.column("#3", width=50, anchor=CENTER)
treeview.heading("#4", text="Rate", anchor=W)
treeview.column("#4", width=150)
treeview.heading("#5", text="Amount", anchor=W)
treeview.column("#5", width=150)
treeview.heading("#6", text="Installation", anchor=W)
treeview.column("#6", width=150)

treeview.grid(row=0, column=0,columnspan="4")
v_scroll.grid(row=0, column=4, sticky='ns')

lable_insta = Label(text="Installation Chg.")
lable_insta.grid(row=2, column=2,sticky=E) 
lable_insta = Label(text="Transport Cost")
lable_insta.grid(row=3, column=2,sticky=E) 
lable_insta = Label(text="Lead and Lift")
lable_insta.grid(row=4, column=2,sticky=E) 

text_insta = Entry(width=20 )
text_insta.config(state=DISABLED)
text_insta.grid(row=2, column=3,sticky=W)
text_Transport_Cost = Entry(width=20 )
text_Transport_Cost.config(state=DISABLED)
text_Transport_Cost.grid(row=3, column=3,sticky=W)
text_Lead_and_Lift = Entry(width=20 )
text_Lead_and_Lift.config(state=DISABLED)
text_Lead_and_Lift.grid(row=4, column=3,sticky=W)


btn_insert_above = Button(master, text ="Insert Row Above", command = insert_row_above, width=15)
btn_insert_above.grid(row="2",column="0", padx=1)
btn_insert_below = Button(master, text ="Insert Row Below", command = insert_row_below, width=15)
btn_insert_below .grid(row="3",column="0", padx=1)
btn_delete = Button(master, text ="Delete Row", command = delete_row, width=10)
btn_delete.grid(row="3",column="1")

B_open = Button(master, text ="Select File", command = file_selector,width=20)
B_open.grid(row="5",column="0")
B_Save = Button(master, text ="Save Json File", command = file_saver,width=20)
B_Save.grid(row="5",column="1",)


mainloop()