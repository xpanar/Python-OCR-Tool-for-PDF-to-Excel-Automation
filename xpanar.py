import pdfplumber       
import json             
import os , sys               
import tkinter as tk    
from tkinter import *   
from tkinter.filedialog import askopenfilename,asksaveasfilename
import tkinter.ttk as ttk   
import openpyxl
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation  
import shutil
from datetime import datetime



def resource_path(relative_path):
    """ Get the absolute path to a resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

xlsm_file = resource_path("sample.xlsm")
json_file = resource_path("presets.json")

shutil.copy2(xlsm_file, resource_path('modified_file.xlsm'))
original_file = resource_path('modified_file.xlsm')
workbook = openpyxl.load_workbook(original_file, keep_vba=True)    

sheet1 = workbook['Client Master']
comboboxes = {}
listboxes = {}


def drag_copy(source_sheet, source_cell, target_range):
    # Get the value and data validation from the source cell
    value = source_sheet[source_cell].value
    validations = [
        validation for validation in source_sheet.data_validations.dataValidation
        if source_cell in validation.ranges
    ]

    # Apply value and validation to the entire target range
    for row in source_sheet[target_range]:
        for cell in row:
            cell.value = value
            for validation in validations:
                new_validation = DataValidation(
                    type=validation.type,
                    formula1=validation.formula1,
                    formula2=validation.formula2,
                    showDropDown=validation.showDropDown,
                    showErrorMessage=validation.showErrorMessage,
                    errorTitle=validation.errorTitle,
                    error=validation.error
                )
                source_sheet.add_data_validation(new_validation)
                new_validation.add(cell)

# Load custom items from the file
def load_custom_items():
    if os.path.exists(json_file):
        with open(json_file, "r") as file:
            return json.load(file)
    return {"OrderType": [], "SalesExecutive": [], "dropdown3": []}

# Save custom items to the file
def save_custom_items(data):
    with open("presets.json", "w") as file:
        json.dump(data, file)

def add_item(combobox_name, entry_widget):
    new_item = entry_widget.get()
    if new_item and new_item not in items[combobox_name]:
        items[combobox_name].append(new_item)
        comboboxes[combobox_name]['values'] = items[combobox_name]
        save_custom_items(items)
        entry_widget.delete(0, tk.END)
        update_listbox(combobox_name)

# Delete an item from the specified combobox
def delete_item(combobox_name, item):
    items[combobox_name].remove(item)
    comboboxes[combobox_name]['values'] = items[combobox_name]
    save_custom_items(items)
    update_listbox(combobox_name)

# Update the listbox below the combobox
def update_listbox(combobox_name):
    listboxes[combobox_name].delete(0, tk.END)
    for item in items[combobox_name]:
        listboxes[combobox_name].insert(tk.END, item)
        listboxes[combobox_name].see(tk.END)


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
    entry = ttk.Entry(root)
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
    sheet1['B1'] = Client_Name.get()
    sheet1['B2'] = Address.get()
    sheet1['B3'] = Email_ID.get()
    sheet1['H1'] = Contact_No.get()
    sheet1['H2'] = Order_Type.get()
    sheet1['H3'] = Sales_Executive.get()
    sheet1['M1'] = Client_Final_Price.get()
    sheet1['M2'] = Client_Discount.get()
    sheet1['M3'] = Dealer_Fenesta_Price_Basic.get()
    sheet1['L13'] = float(text_Lead_and_Lift.get())
    sheet1['L13'].number_format = '₹#,##0.00'
    sheet1['L12'] = float(text_Transport_Cost.get())
    sheet1['L12'].number_format = '₹#,##0.00'
    new_file_name = 'renamed_example.xlsm'

    excel_row = 5
    for row in treeview.get_children():
        row_values = treeview.item(row)["values"]
        print(row_values)
        
        sheet1['A'+str(excel_row)] = row_values[0]
        sheet1['A'+str(excel_row)].number_format = '0000'
        sheet1['B'+str(excel_row)] = row_values[1]
        sheet1['C'+str(excel_row)] = row_values[2]
        sheet1['D'+str(excel_row)] = float(row_values[3])
        sheet1['D'+str(excel_row)].number_format = '₹#,##0.00'
        sheet1['E'+str(excel_row)] = float(row_values[4])
        sheet1['E'+str(excel_row)].number_format = '₹#,##0.00'
        sheet1['F'+str(excel_row)] = float(row_values[5])
        sheet1['F'+str(excel_row)].number_format = '₹#,##0.00'
        
        excel_row = excel_row + 1

    drag_copy(sheet1, 'G5', 'G5:G'+str(excel_row-1))
    sheet1['C'+str(excel_row)] = str("=SUM(C5:C"+str(excel_row-1)+")")
    sheet1['C'+str(excel_row)].font = Font(bold=True)
    sheet1['D'+str(excel_row)] = str("=SUM(D5:D"+str(excel_row-1)+")")
    sheet1['D'+str(excel_row)].font = Font(bold=True)
    sheet1['E'+str(excel_row)] = str("=SUM(E5:E"+str(excel_row-1)+")")
    sheet1['E'+str(excel_row)].font = Font(bold=True)
    sheet1['L9'] = str("=SUM(E5:E"+str(excel_row-1)+")")
    sheet1['F'+str(excel_row)] = str("=SUM(F5:F"+str(excel_row-1)+")")
    sheet1['F'+str(excel_row)].font = Font(bold=True)
    sheet1['L14'] = str("=SUM(F5:F"+str(excel_row-1)+")")

    current_date = datetime.now().strftime("%Y-%m-%d")
    file_path = asksaveasfilename(
        defaultextension=".xlsm",
        filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm")],
        title="Save as",
        initialfile= str(Client_Name.get()) +"  "+ current_date+ ".xlsm"
    )

    # Rename the file
    #modified_file = 'modified_file.xlsm'

    #if os.path.exists(new_file_name):
       # os.remove(new_file_name)  # Remove the existing file if necessary
    workbook.save(original_file)
    os.rename(original_file, file_path)


def extract_text_from_pdf(pdf_path, installation_total):
    all_text = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            all_text += text if text else ""
    line = all_text.split("\n")
    Table = [['0000', '000', '0', '₹0000.00', '₹0000.00']]
    tempT = ["null","null","null","null","null"]

    excel_row = 5

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

            if len(QT) == 6 :
                tempT[3] = QT[len(QT)-4]
            elif len(QT) == 4 :
                tempT[3] = str(QT[len(QT)-2])[1:]
            else :
                tempT[3] = QT[len(QT)-4]

            tempT[4] = QT[len(QT)-1]
            
            print("check hear --")
            print(QT)


            Table.append(tempT)
            
            y = {"Sales Line":str(tempT[0]), "Window Code:": tempT[1], "Qty": tempT[2], "Rate": str(tempT[len(tempT)-2]) ,"Amount": str(tempT[len(tempT)-1])[1:]  }
            filedata.append({"windows": y})

            # sheet1['A'+str(excel_row)] = str(tempT[0])
            # sheet1['B'+str(excel_row)] = tempT[1]
            # sheet1['C'+str(excel_row)] = round(int(tempT[2]),0)
            # sheet1['D'+str(excel_row)] = round(float(str(tempT[len(tempT)-2])),2)
            # sheet1['E'+str(excel_row)] = round(float(str(tempT[len(tempT)-1])[1:]),2)
            # excel_row = excel_row + 1

            treeview.insert("", "end", values=(str(y["Sales Line"]),y["Window Code:"], y["Qty"],y["Rate"], y["Amount"],0.00))

        if temp_line.find("Transport Cost") >= 0:
            QT = line[i].split(" ")
            Transport_cost = str(QT[len(QT)-1])[1:]
            #print("temp_Transport  == " +Transport_cost)
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
            

    #print(filedata)
    #print("total line item = "+str(line_item)+"  "+str(int(line_item)))
    line_item = int(line_item)
    installation = [0] * line_item
    excel_row = 5
    
    child = treeview.get_children()
    sale_line = get_column_values(treeview, 0)
    sale_line_lenth = len(sale_line)

    #print(sale_line_lenth)
    #print(f"Values in 'ID' column: {sale_line}")

    valid_values = ["1st LOT", "2nd LOT", "3rd LOT", "CANCELLED"]

    # Create a data validation object for a dropdown list
    dv = DataValidation(type="list", formula1=f'"{",".join(valid_values)}"', showDropDown=True)
    sheet1.add_data_validation(dv)

    for i in range(len(line),0,-1):
        temp_line = line[i-1]
        # print(str("00"+sale_line[sale_line_lenth-1]))
        if temp_line.find(str("00"+str(sale_line[sale_line_lenth-1]))) >= 0:
            #print(line[i-1])

            temp_line_value = line[i-1].split(" ")

            installation[line_item-1] = temp_line_value[len(temp_line_value)-1]

            #print(treeview.item(child[sale_line_lenth-1], 'values'))
            values = treeview.item(child[sale_line_lenth-1], 'values')
            treeview.item(child[sale_line_lenth-1], values=(values[0], values[1], values[2], values[3],values[4], temp_line_value[len(temp_line_value)-1] ))

            #print("+++++++++++++++++++++++")
            #print(temp_line_value)
            #print("=======================")
            installation_total = float(temp_line_value[len(temp_line_value)-1]) + installation_total
            line_item = line_item - 1
            sale_line_lenth = sale_line_lenth - 1
            sheet1['F'+str(excel_row + sale_line_lenth)] = round(float(str(temp_line_value[len(temp_line_value)-1])),2)
            
            
            if sale_line_lenth <= 0:
                break

            #drag_copy(sheet1, 'G5', 'G5:G'+str(excel_row + len(sale_line) - 1 ))
             

    text_insta.config(state=NORMAL)
    text_insta.delete(0, END)
    text_insta.insert(0, "{:.2f}".format(installation_total))
    #text_insta.config(state=DISABLED)      

    # Convert Treeview data to a dictionary



root = tk.Tk()
root.title("ORwins Fenesta")
root.geometry("690x575")

treeview = ttk.Treeview(root, show="headings", columns=("Sales Line", "Window Code:", "Qty", "Rate", "Amount", "Installation"))
v_scroll = ttk.Scrollbar(root, orient='vertical', command=treeview.yview)
treeview.configure(yscroll=v_scroll.set)

treeview.bind('<Double-1>', on_double_click)
#treeview.pack(fill='both', expand=True)

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

lable_insta = Label(root, text="Installation Chg.")
lable_insta.grid(row=2, column=2,sticky=E) 
lable_Trans = Label(root, text="Transport Cost")
lable_Trans.grid(row=3, column=2,sticky=E) 
lable_Lead = Label(root, text="Lead and Lift")
lable_Lead.grid(row=4, column=2,sticky=E) 

text_insta = Entry(root, width=20 )
text_insta.config(state=DISABLED)
text_insta.grid(row=2, column=3,sticky=W)
text_Transport_Cost = Entry(root, width=20 )
text_Transport_Cost.config(state=DISABLED)
text_Transport_Cost.grid(row=3, column=3,sticky=W)
text_Lead_and_Lift = Entry(root, width=20 )
text_Lead_and_Lift.config(state=DISABLED)
text_Lead_and_Lift.grid(row=4, column=3,sticky=W)

btn_insert_above = Button(root, text ="Insert Row Above", command = insert_row_above, width=15)
btn_insert_above.grid(row="2",column="0", padx=1)
btn_insert_below = Button(root, text ="Insert Row Below", command = insert_row_below, width=15)
btn_insert_below .grid(row="3",column="0", padx=1)
btn_delete = Button(root, text ="Delete Row", command = delete_row, width=10)
btn_delete.grid(row="3",column="1")

B_open = Button(root, text ="Select PDF File", command = file_selector,width=20)
B_open.grid(row="16",column="2", pady= 25)
B_Save = Button(root, text ="Save Excel File", command = file_saver,width=20)
B_Save.grid(row="16",column="3",pady= 25)

items = load_custom_items() 

tk.Label(root, text=f"Order Type :").grid(row="6",column="0")
Order_Type = ttk.Combobox(root, values=items["OrderType"])
Order_Type.grid(row="6",column="1")
tk.Label(root, text=f"Client Name :").grid(row="7",column="0")
Client_Name = tk.Entry(root,  width=50)
Client_Name.grid(row="7",column="1",columnspan=2)
tk.Label(root, text=f"Address :").grid(row="8",column="0")
Address = tk.Entry(root,  width=50)
Address.grid(row="8",column="1",columnspan=2)
tk.Label(root, text=f"Email ID :").grid(row="9",column="0")
Email_ID = tk.Entry(root,  width=50)
Email_ID.grid(row="9",column="1",columnspan=2)
tk.Label(root, text=f"Contact No :").grid(row="10",column="0")
Contact_No = tk.Entry(root, width=50)
Contact_No.grid(row="10",column="1",columnspan=2)
tk.Label(root, text=f"CSE (Sales Executive) :").grid(row="11",column="0")
Sales_Executive = ttk.Combobox(root, values=items["SalesExecutive"])
Sales_Executive.grid(row="11",column="1")
tk.Label(root, text=f"Client Final Price :").grid(row="12",column="0")
Client_Final_Price = tk.Entry(root, width=50)
Client_Final_Price.grid(row="12",column="1",columnspan=2)
tk.Label(root, text=f"Order Qty :").grid(row="13",column="0")
Order_Qty = tk.Spinbox(root, from_=0, to=500, increment=1)
Order_Qty.grid(row="13",column="1")
tk.Label(root, text=f"Client Discount % :").grid(row="14",column="0")
Client_Discount = tk.Spinbox(root, from_=0, to=100, increment=1)
Client_Discount.grid(row="14",column="1")
tk.Label(root, text=f"DP X % :").grid(row="15",column="0")
Dealer_Fenesta_Price_Basic = tk.Entry(root, width=50)
Dealer_Fenesta_Price_Basic.grid(row="15",column="1",columnspan=2)


# Run the application
root.mainloop()
#mainloop()