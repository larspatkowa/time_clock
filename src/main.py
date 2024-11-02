import csv
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def calculate_time(file_path):
    user_data = {}
    first_date = None

    with open(file_path, 'r') as file:
        reader = csv.reader(file)
        next(reader)
        for row in reader:
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
        result.append([user_id, data['user_name'], f"{total_hours:.2f} hours", f'"{"; ".join(abnormalities)}"'])

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
                writer.writerow(['User ID', 'User Name', 'Total Time Clocked In', 'Abnormalities'])
                writer.writerows(result)
            messagebox.showinfo("Success", f"File saved successfully at {save_path}")

def open_file():
    file_path = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*"))
    )
    if file_path:
        calculate_time(file_path)

root = tk.Tk()
root.title("Time Clock Calculator")
root.geometry("400x200")
root.resizable(False, False)

style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12), padding=10)
style.configure('TLabel', font=('Helvetica', 14))

frame = ttk.Frame(root, padding="20")
frame.pack(fill=tk.BOTH, expand=True)

label = ttk.Label(frame, text="Time Clock Calculator")
label.pack(pady=10)

open_button = ttk.Button(frame, text="Open CSV File", command=open_file)
open_button.pack(pady=20)

root.mainloop()