import csv
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os

def calculate_time(file_path):
    user_data = {}
    first_date = None
    file_ext = os.path.splitext(file_path)[1].lower()

    if file_ext == '.csv':
        df = pd.read_csv(file_path)
    elif file_ext in ['.xls', '.xlsx']:
        df = pd.read_excel(file_path)
    else:
        messagebox.showerror("Error", "Unsupported file format.")
        return

    for index, row in df.iterrows():
        user_id, user_name, check_time_str, status = row
        check_time = datetime.strptime(check_time_str, '%Y/%m/%d %H:%M:%S')
        if first_date is None or check_time < first_date:
            first_date = check_time
        if user_id not in user_data:
            user_data[user_id] = {
                'user_name': user_name,
                'times': []
            }
        user_data[user_id]['times'].append((check_time, status))

    result = []
    for user_id, data in user_data.items():
        sorted_times = sorted(data['times'], key=lambda x: x[0])
        total_time = 0
        clock_in_time = None
        abnormalities = []
        for check_time, status in sorted_times:
            if status == 'IN':
                if clock_in_time is not None:
                    abnormalities.append(f"Unmatched 'IN' at {clock_in_time}")
                clock_in_time = check_time
            elif status == 'OUT':
                if clock_in_time is None:
                    abnormalities.append(f"Unmatched 'OUT' at {check_time}")
                else:
                    time_diff = (check_time - clock_in_time).total_seconds()
                    total_time += time_diff
                    clock_in_time = None
        if clock_in_time is not None:
            midnight = clock_in_time.replace(hour=23, minute=59, second=59)
            time_diff = (midnight - clock_in_time).total_seconds()
            total_time += time_diff
            abnormalities.append(f"Unmatched 'IN' at {clock_in_time}")

        if sorted_times and sorted_times[0][1] == 'OUT':
            first_out_time = sorted_times[0][0]
            midnight_before = first_out_time.replace(hour=0, minute=0, second=0)
            time_diff = (first_out_time - midnight_before).total_seconds()
            total_time += time_diff
            abnormalities.append(f"Unmatched 'OUT' at {first_out_time}")

        total_hours = total_time / 3600
        hours = int(total_hours)
        minutes = int((total_hours - hours) * 60)
        total_hhmm = f"{hours:02d}:{minutes:02d}"
        result.append([user_id, data['user_name'], f"{total_hours:.2f} hours", total_hhmm, f'"{"; ".join(abnormalities)}"'])

    if first_date:
        date_str = first_date.strftime('%Y%m%d')
        save_path = filedialog.asksaveasfilename(
            initialfile=f"time_calculation_{date_str}.csv",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if save_path:
            with open(save_path, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['User ID', 'User Name', 'Total Time Clocked In', 'Total HH:MM', 'Abnormalities'])
                writer.writerows(result)
            messagebox.showinfo("Success", f"File saved successfully at {save_path}")

def open_file():
    file_path = filedialog.askopenfilename(
        title="Select CSV or Excel File",
        filetypes=(("CSV and Excel Files", "*.csv *.xls *.xlsx"), ("All Files", "*.*"))
    )
    if file_path:
        calculate_time(file_path)

def save_file(result):
    save_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if save_path:
        with open(save_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['User ID', 'User Name', 'Total Time Clocked In', 'Total HH:MM', 'Abnormalities'])
            writer.writerows(result)
        messagebox.showinfo("Success", f"File saved successfully at {save_path}")

def create_ui():
    root = tk.Tk()
    root.title("Time Clock Calculator")
    root.geometry("300x200")
    root.resizable(False, False)
    
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TButton', font=('Helvetica', 12), padding=10)
    style.configure('TLabel', font=('Helvetica', 12), padding=5)
    
    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill='both', expand=True)
    
    title_label = ttk.Label(main_frame, text="Time Clock Calculator", font=('Helvetica', 16, 'bold'))
    title_label.pack(pady=10)
    
    load_button = ttk.Button(main_frame, text="Load File", command=open_file)
    load_button.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    create_ui()