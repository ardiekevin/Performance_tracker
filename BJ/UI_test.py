import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import ctypes

def run_script():
    selected_currency = currency_combobox.get()
    selected_frequency = frequency_var.get()

    if selected_frequency == "Monthly":
        script_name = "dynamic_gsheets_monthly.py"
    elif selected_frequency == "Weekly":
        script_name = "dynamic_gsheets_weekly.py"
    elif selected_frequency == "Daily":
        script_name = "dynamic_gsheets_daily.py"
    else:
        return

    if selected_currency == "BAJI USD":
        spreadsheet_name = "Baji 2024 Biz Performance Tracker - USD"
    elif selected_currency == "BAJI PHP":
        spreadsheet_name = "Baji 2024 Biz Performance Tracker - PHP"
    elif selected_currency == "BAJI VND":
        spreadsheet_name = "Baji 2024 Biz Performance Tracker - VND"
    elif selected_currency == "BAJI PKR":
        spreadsheet_name = "Baji 2024 Biz Performance Tracker - PKR"
    elif selected_currency == "BAJI INR":
        spreadsheet_name = "Baji 2024 Biz Performance Tracker - INR"
    elif selected_currency == "BAJI BDT":
        spreadsheet_name = "Baji 2024 Biz Performance Tracker - BDT"
    else:
        return

    confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to run the script?")
    if confirmation:
        subprocess.Popen(["python", rf"C:\Users\ardie.asilo\Desktop\Performance_Tracker\BJ\{script_name}", "--spreadsheet", spreadsheet_name])

root = tk.Tk()
root.title("Biz Performance Filler")
root.geometry("400x200")
root.resizable(False, False)

font_style = "Century Gothic"
style = ttk.Style()

style.configure("TCombobox", font=(font_style, 14))
style.configure("TRadiobutton", font=(font_style, 14))

label = tk.Label(root, text="Choose Brand and Currency:", font=(font_style, 12))
label.grid(row=0, column=0, columnspan=3, padx=0, pady=0)
currency_combobox = ttk.Combobox(root, values=["BAJI USD", "BAJI PHP", "BAJI VND", "BAJI PKR", "BAJI INR", 
                                               "BAJI BDT"], width=30, height=20, style = "TCombobox")

currency_combobox.grid(row=1, column=0, columnspan=3, padx=0, pady=10)

frequency_var = tk.StringVar()
frequency_var.set("Monthly")

monthly_radio = ttk.Radiobutton(root, text="Monthly", variable=frequency_var, value="Monthly", style = "TRadiobutton")
monthly_radio.grid(row=2, column=0, padx=(5,0), pady=5)

weekly_radio = ttk.Radiobutton(root, text="Weekly", variable=frequency_var, value="Weekly", style = "TRadiobutton")
weekly_radio.grid(row=2, column=1, padx=(47,83), pady=5)

daily_radio = ttk.Radiobutton(root, text="Daily", variable=frequency_var, value="Daily", style = "TRadiobutton")
daily_radio.grid(row=2, column=2, padx=(0,5),pady=5)

button = tk.Button(root, text="Run Script", command=run_script, height=2, width=20, font=(font_style, 14))
button.grid(row=3, column=0, columnspan=3, padx=0, pady=10)

root.eval('tk::PlaceWindow %s center' % root.winfo_pathname(root.winfo_id()))

root.deiconify()
root.mainloop()
