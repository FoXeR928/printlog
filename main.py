from evtx import PyEvtxParser
import xml.etree.ElementTree as ET
from tkinter import Tk, Label, ttk, Button, Entry,LEFT,RIGHT, BOTH,SE,NE

path="C:/Users/akimov.n.r/Desktop/Microsoft-Windows-PrintService%4Operational.evtx"
prop=['Param2','Param2','Param2','Param2',]
width=800
height=400
columns=('date','docs', 'users', 'computers', 'printers')
columns_head={'date':'Дата и время','docs':'Документы', 'users':'Пользователь', 'computers':'Компьютер', 'printers':'Принтер'}

def parse():
    parser = PyEvtxParser(path)
    for record in parser.records():
        tree = ET.ElementTree(ET.fromstring(record['data']))
        print(tree.getroot())
        print(f'------------------------------------------')

window=Tk()
window.geometry(f'{width}x{height}')
window.title('События печати')

find=Button(window,text="Найти", width=40)
entry = Entry(window)

tree = ttk.Treeview(columns=columns, show="headings")

for key,name in columns_head.items():
    tree.heading(key, text=name)


entry.pack(anchor=NE,padx=5, pady=5)
find.pack(anchor=NE,padx=5, pady=5)
tree.pack(fill=BOTH, padx=5, pady=5)
parse()
window.mainloop()

