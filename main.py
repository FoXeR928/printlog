from evtx import PyEvtxParser
import xml.etree.ElementTree as ET
from xml.dom.minidom import parseString
import json
from tkinter import Tk, Label, ttk, Button, Entry,LEFT,RIGHT, BOTH,SE,NE, END, CENTER, Scrollbar

path="C:/Users/akimov.n.r/Desktop/Microsoft-Windows-PrintService%4Operational.evtx"
prop=['Param2','Param3','Param4','Param5','Param6']
event_id=307
width=800
height=400
columns=('date','docs', 'users', 'computers', 'printers')
columns_head={'date':'Дата и время','docs':'Документы', 'users':'Пользователь', 'computers':'Компьютер', 'printers':'Принтер'}

def parse():
    tree.delete(*tree.get_children())
    all_users=[]
    parser = PyEvtxParser(path)
    for record in parser.records_json():
        all_users_data=[]
        json_date=json.loads(record['data'])
        event_id_date=json_date['Event']['System']['EventID']
        if event_id_date==event_id:
            userdata=json_date['Event']['UserData']['DocumentPrinted']
            timeprint=record['timestamp']
            all_users_data.append(timeprint)
            for param in prop:
                all_users_data.append(userdata[param])
            all_users.append(all_users_data)
    for person in all_users:
        tree.insert("", END, values=person)

def find():
    tree.delete(*tree.get_children())
    all_users=[]
    parser = PyEvtxParser(path)
    for record in parser.records_json():
        all_users_data=[]
        json_date=json.loads(record['data'])
        event_id_date=json_date['Event']['System']['EventID']
        if event_id_date==event_id:
            user_name=json_date['Event']['UserData']['DocumentPrinted']['Param3']
            if entry.get() in user_name:
                userdata=json_date['Event']['UserData']['DocumentPrinted']
                timeprint=record['timestamp']
                all_users_data.append(timeprint)
                for param in prop:
                    all_users_data.append(userdata[param])
                all_users.append(all_users_data)
    for person in all_users:
        tree.insert("", END, values=person)

def conver_csv():
    print(*tree.get_children())

window=Tk()
window.title('События печати')

update_button=Button(window,text="Обновить", width=40, command=parse)
find_button=Button(window,text="Найти", width=40, command=find)
entry = Entry(window, width=50)
tree = ttk.Treeview(columns=columns, show="headings")
parse()
csv_button=Button(window,text="Конвертировать в csv", width=40, command=conver_csv)

for key,name in columns_head.items():
    tree.heading(key, text=name)

update_button.grid(row=0, column=0,padx=5, pady=5)
entry.grid(row=0, column=1,padx=5, pady=5)
find_button.grid(row=0, column=2,padx=5, pady=5)
tree.grid(row=1, column=0, columnspan=3, padx=5, pady=5)
find_button.grid(row=3, column=2,padx=5, pady=5)
window.mainloop()

