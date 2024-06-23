# Gate Entry Report

## Overview
This Python script processes gate entry data from an Excel file, calculates various statistics, and displays the results in a graphical user interface (GUI) using Tkinter. The script is useful for summarizing gate activity and RFID usage over a period of time.

## Dependencies
The script requires the following Python libraries:
- `tkinter`
- `ttk`
- `openpyxl`

## Installation
Before running the script, ensure you have the necessary libraries installed. You can install `openpyxl` using pip:

```bash
pip install openpyxl
```

## How It Works
1. **Import Libraries**:
   The script imports required libraries for creating a GUI (`tkinter`, `ttk`) and for reading Excel files (`openpyxl`).

2. **Initialize Variables**:
   The script initializes several counters in a dictionary to track the number of gate entries and RFID usage at different gates.

3. **Load Excel Workbook**:
   The script loads an Excel workbook from a specified path and accesses the first sheet to process the data.

4. **Process Rows**:
   The script iterates over the rows in the Excel sheet, starting from the second row. It reads values from specific columns to determine the type of gate event and increments the appropriate counters.

5. **Create GUI**:
   The script creates a Tkinter window and sets up a Treeview widget to display the gate statistics.

6. **Populate Treeview**:
   The script populates the Treeview widget with the gate statistics calculated from the Excel data.

## Example Code
```python
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
wb = openpyxl.load_workbook("C:\\Gate\\Book1.xlsx")
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
```

### Explanation of Variables
- **Gate Counts**: A dictionary to track the number of entries and RFID usage at different gates.
- **Excel Data**: The script reads gate event data from an Excel file and updates the gate counts accordingly.

### Usage
To run the script, simply execute it with Python:

The script processes the gate entry data, calculates statistics, and displays the results in a graphical user interface.
