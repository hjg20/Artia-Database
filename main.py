import tkinter as tk
import pandas as pd
from tkinter import ttk
import xlsxwriter

data = pd.read_excel('Data.xlsx', dtype={'NDC11': str})
mb = pd.read_excel('MB.xlsx', dtype={'NDC11': str})
name_alias = pd.read_excel('NameAlias.xlsx', dtype={'NDC11': str})
pdl_master = pd.read_excel('PDLMaster.xlsx')

agents = mb['ProductName2'].unique().tolist()
agent = ""
states = pdl_master['ST'].unique().tolist()
state = ""


def export():
    global df
    writer = pd.ExcelWriter('export.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='export from code', index=False)
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
    ct_names = []
    ba_names = []
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
                ba_names.append(name_alias['Bid Analysis Drug Name'][j])
    for i in range(len(ct_names)):
        counter = 0
        for j in range(len(pdl_master['State'])):
            if ct_names[i] == pdl_master['Drug (generic)'][j] and ba_names[i] == pdl_master['Bid Analysis Drug Name'][j] and state == pdl_master['ST'][j] and counter < 1:
                counter += 1
                statuses.append(pdl_master['PDL Status'][j])
    df['ST'] = sm_states
    df['NDC11'] = ndcs
    df['ProductName2'] = productname2s
    df['PDL Status'] = statuses
    dataframe_label.config(text=df)


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
    ba_names = []
    global df
    global agent
    global state
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
    #df['PDL Status'] = statuses
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
    df['Market Share'] = df['Scripts'].div(df['Scripts'].where(df['ST'].str.contains('Total')).bfill())
    for i in range(len(df['Market Share'])):
        if 'Total' in df['ST'][i]:
            df['Market Share'][i] = ''
    dataframe_label.config(text=df)

































































root = tk.Tk()
root.title('PDL Database Search')
root.geometry('1700x800+550+550')

agent_list_items = tk.Variable(value=agents)

agent_listbox = tk.Listbox(root, listvariable=agent_list_items, height=5)
agent_listbox.grid(column=0, row=0)

state_list_items = tk.Variable(value=states)

state_listbox = tk.Listbox(root, listvariable=state_list_items, height=5)
state_listbox.grid(column=4, row=0)

agent_button = tk.Button(text='Select', command=select_agent)
agent_button.grid(column=0, row=1)

state_button = tk.Button(text='Select', command=select_state)
state_button.grid(column=4, row=1)

pdl_button = tk.Button(text='Get PDL Status', command=get_pdl_status)
pdl_button.grid(column=1, row=3)

bid_button = tk.Button(text='Create Bid Analysis', command=bid)
bid_button.grid(column=3, row=3)

export_button = tk.Button(text='Export', command=export)
export_button.grid(column=2, row=5)

agent_label = tk.Label(root, text=agent)
agent_label.grid(column=0, row=2)

state_label = tk.Label(root, text=state)
state_label.grid(column=4, row=2)

df = pd.DataFrame()

dataframe_label = tk.Label(root, text=df, width=200)
dataframe_label.grid(column=0, row=4, columnspan=5)

root.mainloop()