import tkinter as tk
from tkinter import *
import pandas as pd
from tkinter import ttk
import math
from pandastable import Table, TableModel
import xlsxwriter

data = pd.read_excel('Data.xlsx', dtype={'NDC11': str})
mb = pd.read_excel('MB.xlsx', dtype={'NDC11': str})
name_alias = pd.read_excel('NameAlias.xlsx', dtype={'NDC11': str})
pdl_master = pd.read_excel('PDLMaster.xlsx')

agents = mb['ProductName2'].unique().tolist()
agent = ""
states = sorted(pdl_master['ST'].unique().tolist())
state = ""
writer = pd.ExcelWriter('export.xlsx', engine='xlsxwriter')


class TestApp(Frame):
    def __init__(self, parent=None):
        global df
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('2000x2000+200+100')
        self.main.title('Table app')
        f = Frame(self.main)
        f.grid(column=0, row=4, columnspan=3, )
        self.table = pt = Table(f, dataframe=df,
                                showtoolbar=True, showstatusbar=True)
        pt.show()
        return


def export_pdl():
    global df
    writer = pd.ExcelWriter('export.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, engine='xlsxwriter')
    writer.save()


def export_bid():
    global df
    global writer
    df.to_excel(writer, sheet_name='Utilization Summary', index=False, engine='xlsxwriter')
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column)) + 1.5
        col_idx = df.columns.get_loc(column)
        writer.sheets['Utilization Summary'].set_column(col_idx, col_idx, column_width)
    dollar_format = writer.book.add_format({'num_format': "$#,##0.00"})
    comma_format = writer.book.add_format({"num_format": 37})
    util_summary_sheet = writer.sheets["Utilization Summary"]
    util_summary_sheet.set_column('H:H', None, comma_format)
    util_summary_sheet.set_column('I:I', None, comma_format)
    util_summary_sheet.set_column('L:L', 12.5, dollar_format)

    wb = xlsxwriter.Workbook(filename="export.xlsx")
    ws = wb.add_worksheet(name='Util Sum')

    def ignore_nan(_ws, _row, col, number, _format=None):
        if math.isnan(number):
            return _ws.write_blank(_row, col, None, _format)
        else:
             return None
        
    ws.add_write_handler(float, ignore_nan)

    row_num = 1
    rows = df.values.tolist()
    sum_of_scripts = 0
    sum_of_units = 0
    sum_of_total_amount = 0
    for row in rows:
        if 'Total' not in row[4]:
            row_num += 1
        else:
            scripts = float(row[8])
            units = float(row[7])
            total_amount = float(row[11])
            sum_of_scripts += scripts
            sum_of_units += units
            sum_of_total_amount += total_amount
            row_num += 1
    row_num = 1
    for row in rows:
        if 'Total' not in row[4] and 'Total' not in row[1]:
            ws.set_row(row_num, None, None, {'level': 2, "hidden": True})
            ws.write_row(row_num, 0, row)
            row_num += 1
        else:
            ws.set_row(row_num, None, None, {'level': 1})
            ws.write_row(row_num, 0, row)
            units = float(row[7])
            scripts = float(row[8])
            try:            
                units_per_rx = units / scripts        
            except ZeroDivisionError:            
                units_per_rx = 0        
            if 'Total' in row[1]:            
                ws.write(row_num, 9, '')            
                row_num += 1        
            else:            
                ws.write(row_num, 9, units_per_rx)            
                row_num += 1

    header = wb.add_format({"bg_color": "00314C", "font_color": 'white', 'bold': True})
    dollar_column = wb.add_format({"num_format": "$#,##0.00"})
    number_column = wb.add_format({"num_format": "#,##0"})
    decimal_column = wb.add_format({"num_format": "#,##0.0"})
    percentage_column = wb.add_format({"num_format": "0.00%"})
    ws.write('A1', "ID", header)
    ws.write('B1', "St", header)
    ws.write('C1', "NDC11", header)
    ws.write('D1', "ProductNameLong", header)
    ws.write('E1', "ProductName2", header)
    ws.write('F1', "Quarter", header)
    ws.write('G1', "Year", header)
    ws.write('H1', "Units", header)
    ws.write('I1', "Scripts", header)
    ws.write('J1', 'Units/Rx', header)
    ws.write('K1', 'Market Share', header)
    ws.write('L1', 'Total Amount', header)
    ws.write('M1', 'PDL Status', header)
    bold = wb.add_format({"bold": True})
    ws.set_column('E:E', 30, bold)
    ws.set_column('H:H', 10, number_column)
    ws.set_column('I:I', 10, number_column)
    ws.set_column('J:J', 10, decimal_column)
    ws.set_column('K:K', 12, percentage_column)
    ws.set_column('L:L', 15, dollar_column)
    ws.write(row_num, 4, 'Grand Total')
    ws.write(row_num, 7, sum_of_units)
    ws.write(row_num, 8, sum_of_scripts)
    ws.write(row_num, 11, sum_of_total_amount)


    ws.activate()
    wb.close()
    writer.save()


def select_agent():
    global agent
    for i in agent_listbox.curselection():
        agent_label.config(text=agent_listbox.get(i))
        agent = agent_listbox.get(i)


def select_state():
    global state
    for i in state_listbox.curselection():
        state_label.config(text=state_listbox.get(i))
        state = state_listbox.get(i)


def get_pdl_status():
    ndcs = []
    productname2s = []
    sm_states = []
    statuses = []
    clients = []
    ct_names = []
    global df
    global agent
    global state
    df = pd.DataFrame()
    for i in range(len(mb['ProductName2'])):
        if mb['ProductName2'][i] == agent:
            ndcs.append(mb['NDC11'][i])
            productname2s.append(mb['ProductName2'][i])
            sm_states.append(state)
    for i in ndcs:
        for j in range(len(name_alias['NDC11'])):
            if i == name_alias['NDC11'][j]:
                ct_names.append(name_alias['Drug (generic)'][j])
                clients.append(name_alias['Client'][j])
    for i in range(len(ct_names)):
        counter = 0
        for j in range(len(pdl_master['State'])):
            if ct_names[i] == pdl_master['Drug (generic)'][j] and clients[i] == pdl_master['Client'][j] and state == pdl_master['ST'][j] and counter < 1:
                counter += 1
                statuses.append(pdl_master['PDL Status'][j])
    df['ST'] = sm_states
    df['NDC11'] = ndcs
    df['ProductName2'] = productname2s
    df['PDL Status'] = statuses
    TestApp()


def bid():
    ndcs = []
    ids = []
    new_ndcs = []
    productnamelongs = []
    productname2s = []
    sm_states = []
    units = []
    scripts = []
    total_amount = []
    year = []
    quarter = []
    statuses = []
    ct_names = []
    clients = []
    global df
    global agent
    global state
    global writer
    df = pd.DataFrame()
    for i in range(len(mb['ProductName2'])):
        ndcs.append(mb['NDC11'][i])
    for i in ndcs:
        for j in range(len(data['NDC11'])):
            if i == data['NDC11'][j] and state == data['St'][j]:
                ids.append(data['ID'][j])
                new_ndcs.append(data['NDC11'][j])
                sm_states.append(data['St'][j])
                productnamelongs.append(data['ProductNameLong'][j])
                productname2s.append(data['ProductName2'][j])
                year.append(data['Year'][j])
                quarter.append(data['Quarter'][j])
                units.append(data['Units'][j])
                scripts.append(data['Scripts'][j])
                total_amount.append(data['Total Amount'][j])
        
    df['ID'] = ids            
    df['ST'] = sm_states
    df['NDC11'] = new_ndcs
    df['ProductNameLong'] = productnamelongs
    df['ProductName2'] = productname2s
    df['Quarter'] = quarter
    df['Year'] = year
    df['Units'] = units
    df['Scripts'] = scripts
    df['Total Amount'] = total_amount
    result_df = pd.DataFrame(columns=df.columns)
    for st in df['ST'].unique():
        bystate_rows = df[df['ST']==st]
        l = []
        for pn2 in bystate_rows['ProductName2'].unique():
            pr = bystate_rows[bystate_rows['ProductName2'] == pn2]
            l.append(pr)
            units_sub = pr['Units'].sum()
            scripts_sub = pr['Scripts'].sum()
            ta_sub = pr['Total Amount'].sum()
            sub_row = pd.DataFrame([['', st, '', '', f'{pn2} Total', '', '', units_sub, scripts_sub, ta_sub]], columns=df.columns)
            l.append(sub_row)
        units_sub = bystate_rows['Units'].sum()
        script_sub = bystate_rows['Scripts'].sum()
        ta_sub = bystate_rows['Total Amount'].sum()
        sub_row = pd.DataFrame([['', f'{st} Total', '', '', '', '', '', units_sub, script_sub, ta_sub]], columns=df.columns)
        l.append(sub_row)
        result_df = pd.concat([result_df] + l)
    result_df.reset_index(drop=True, inplace=True)
    df = result_df
    df.insert(9, 'Units/Rx', None)
    df.insert(10, 'Market Share', None)
    df.insert(12, 'PDL Status', None)
    df['Market Share'] = df['Scripts'].div(df['Scripts'].where(df['ST'].str.contains('Total')).bfill())
    for i in range(len(df['Market Share'])):
        if 'Total' in df['ST'][i]:
            df['Market Share'][i] = ''
    for i in range(len(df['ProductName2'])):
        for j in range(len(name_alias['ProductName2'])):
            if name_alias['ProductName2'][j] in df['ProductName2'][i] and 'Total' in df['ProductName2'][i]:
                ct_names.append(name_alias['Drug (generic)'][j])
                clients.append(name_alias['Client'][j])
        for i in range(len(ct_names)):
            counter = 0
            for j in range(len(pdl_master['State'])):
                if ct_names[i] == pdl_master['Drug (generic)'][j] and clients[i] == pdl_master['Client'][j] and state == pdl_master['ST'][j] and counter < 1:
                    counter += 1
                    statuses.append(pdl_master['PDL Status'][j])
    counter = 0
    print(statuses)            
    for i in range(len(df['ProductName2'])):
        if 'Total' in df['ProductName2'][i]:
            df['PDL Status'][i] = statuses[counter]
            counter += 1
        else:
            df['PDL Status'][i] = ''
    TestApp()


root = tk.Tk()
root.title('PDL Database Search')
root.geometry('1700x800+550+550')

agent_list_items = tk.Variable(value=agents)

agent_listbox = tk.Listbox(root, listvariable=agent_list_items, height=5)
agent_listbox.grid(column=0, row=0)

state_list_items = tk.Variable(value=states)

state_listbox = tk.Listbox(root, listvariable=state_list_items, height=5)
state_listbox.grid(column=2, row=0)

agent_button = tk.Button(text='Select', command=select_agent)
agent_button.grid(column=0, row=1)

state_button = tk.Button(text='Select', command=select_state)
state_button.grid(column=2, row=1)

pdl_button = tk.Button(text='Get PDL Status', command=get_pdl_status)
pdl_button.grid(column=0, row=3)

bid_button = tk.Button(text='Create Bid Analysis', command=bid)
bid_button.grid(column=2, row=3)

export_button = tk.Button(text='Export PDL Status', command=export_pdl)
export_button.grid(column=1, row=5)
export_button = tk.Button(text='Export Bid Analysis', command=export_bid)
export_button.grid(column=1, row=6)

agent_label = tk.Label(root, text=agent)
agent_label.grid(column=0, row=2)

state_label = tk.Label(root, text=state)
state_label.grid(column=2, row=2)

df = pd.DataFrame()

dataframe_label = tk.Label(root, text=df, width=200)
dataframe_label.grid(column=0, row=4, columnspan=3)

app = TestApp()

root.mainloop()
app.mainloop()