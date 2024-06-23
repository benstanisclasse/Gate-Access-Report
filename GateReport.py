import tkinter as tk
from tkinter import ttk
import openpyxl

# Initialize gate counts and RFID counts
gate_counts = {'Burgundy': {'gate': 0, 'rfid': 0}, 
               'Flanders': {'gate': 0, 'rfid': 0}, 
               'Atlantic': {'gate': 0, 'rfid': 0}, 
               'Monaco': {'gate': 0, 'rfid': 0}, 
               'Normandy': {'gate': 0, 'rfid': 0}, 
               'Saxony': {'gate': 0, 'rfid': 0}}

# Read data from Excel file
wb = openpyxl.load_workbook("C:\Gate\Book1.xlsx")
sh = wb["Sheet1"]
lastRow = len(sh['A'])

for i in range(2, lastRow):
    cellValue = sh['D'+ str(i)].value
    RFIDNumber = sh['B'+ str(i)].value
    
    if cellValue == 'BurgGate (In)':
        gate_counts['Burgundy']['gate'] += 1
    elif cellValue == 'BurgGate (Out)':
        gate_counts['Burgundy']['rfid'] += 1  
    elif cellValue == 'FlanGate (In)':
        gate_counts['Flanders']['gate'] += 1
    elif cellValue == 'FlanGate (Out)':
        gate_counts['Flanders']['rfid'] += 1
    elif cellValue == 'MonaGate (In)':
        gate_counts['Monaco']['gate'] += 1
    elif cellValue == 'MonaGate (Out)':
        gate_counts['Monaco']['rfid'] += 1
    elif cellValue == 'MainGate (In)':
        gate_counts['Atlantic']['gate'] += 1
    elif cellValue == 'MainGate (Out)':
        gate_counts['Atlantic']['rfid'] += 1
    elif cellValue == 'NormGate (In)':
        gate_counts['Normandy']['gate'] += 1
    elif cellValue == 'NormGate (Out)':
        gate_counts['Normandy']['rfid'] += 1
    elif cellValue == 'SaxoGate (In)':
        gate_counts['Saxony']['rfid'] += 1
    elif cellValue == 'SaxoGate (Out)':
        gate_counts['Saxony']['gate'] += 1

# Create main application window
ws = tk.Tk()
ws.title('Kings Point Gate Report')
ws.geometry('500x500')
ws['bg'] = '#AC99F2'

game_frame = ttk.Frame(ws)
game_frame.pack()

# Create Treeview widget
my_game = ttk.Treeview(game_frame)
my_game['columns'] = ('Gate', 'Total', 'Daily Avg', 'RFID', 'Barcode')

# Define column headings
my_game.heading('#0', text='', anchor=tk.CENTER)
my_game.heading('Gate', text='Gate', anchor=tk.CENTER)
my_game.heading('Total', text='Total', anchor=tk.CENTER)
my_game.heading('Daily Avg', text='*Daily Avg', anchor=tk.CENTER)
my_game.heading('RFID', text='RFID', anchor=tk.CENTER)
my_game.heading('Barcode', text='Barcode', anchor=tk.CENTER)

# Set column widths
my_game.column("#0", width=0, stretch=tk.NO)
my_game.column("Gate", anchor=tk.CENTER, width=80)
my_game.column("Total", anchor=tk.CENTER, width=80)
my_game.column("Daily Avg", anchor=tk.CENTER, width=80)
my_game.column("RFID", anchor=tk.CENTER, width=80)
my_game.column("Barcode", anchor=tk.CENTER, width=80)

# Populate Treeview with data
for idx, (gate, counts) in enumerate(gate_counts.items()):
    gate_total = counts['gate'] + counts['rfid']
    daily_avg = round(counts['gate'] / 30.5)
    my_game.insert(parent='', index='end', iid=idx, text='', 
                   values=(gate, gate_total, daily_avg, counts['rfid'], counts['gate']))

# Pack Treeview widget
my_game.pack()

ws.mainloop()
