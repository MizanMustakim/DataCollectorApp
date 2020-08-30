import datetime as dt
import time as tm
import gspread
import tkinter as tk
from tkinter import ttk
from config import api_key

# Load the google spreadsheet file with API key
gc = gspread.service_account(filename='creds.json')
spread_sheet = gc.open_by_key(api_key)
worksheet = spread_sheet.sheet1

# Make the GUI windows
win = tk.Tk()
win.title("Data Collector")
win.configure(background='Green')
win.resizable(500, 500)

# 1st Labeling
ttk.Label(win, text='------Welcome to our e-service------',
          font=50).grid(column=0, columnspan=3)

# Getting Customer Name
ttk.Label(win, text='Name:').grid(column=1, row=5, sticky=tk.W)
name = tk.StringVar()
nameEntered = ttk.Entry(win, width=20, textvariable=name)
nameEntered.focus()
nameEntered.grid(column=2, row=5, sticky=tk.W)

# Getting Customer ID
ttk.Label(win, text='Customer ID:').grid(column=1, row=7, sticky=tk.W)
id = tk.StringVar()
idEntered = ttk.Entry(win, width=10, textvariable=id)
idEntered.grid(column=2, row=7, sticky=tk.W)

# Getting Product Details
ttk.Label(win, text='Product Size: ').grid(column=1, row=8, sticky=tk.W)
measure = tk.IntVar()
measureChoosen = ttk.Combobox(
    win, width=6, textvariable=measure)
measureChoosen['values'] = [1, 2, 3, 4, 5, 10, 20, 40, 50, 70, 90, 100]
measureChoosen.grid(column=2, row=8, sticky=tk.W)
measureChoosen.current(0)

# Converter function between (inch) and (cm)


def convert(measure):
    # UNIT = input("(inch) or (cm): ")
    if unit.get() == 'inch':
        return "{:.2f} cm".format(abs(float(measure * 2.54)))
    elif unit.get() == 'cm':
        return '{:.2f} inch'.format(float(measure / 2.54))


# Getting measure converting input
ttk.Label(win, text='inch or cm').grid(column=1, row=9, sticky=tk.W)
unit = tk.StringVar()
unitRad = tk.Entry(win, width=10, textvariable=unit)
unitRad.grid(column=2, row=9, sticky=tk.W)

# Automatic stored the current date and time by using datetime module
dtnow = dt.datetime.fromtimestamp(tm.time())
date = '{}/ {}/ {}'.format(dtnow.year, dtnow.month, dtnow.day)
time = '{}:{}:{}'.format(dtnow.hour, dtnow.minute, dtnow.second)

# function for save button command


def workSave():
    dictionary = {'Customer Name': name.get(),
                  'Customer ID': id.get(),
                  'Product Size': convert(measure.get()),
                  'Order Date': date,
                  'Order Time': time}
    # files = [name.get(), id.get(), unit.get(), date, time]
    action.configure(worksheet.append_row(list(dictionary.values())))


# Save button
action = ttk.Button(win, text='Save', width=5, command=workSave)
action.grid(column=2, row=15, sticky=tk.W)

# close the loop
win.mainloop()
