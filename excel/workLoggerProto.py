import tkinter
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
from datetime import datetime, timedelta

# Global
file_path = ""

def openFile():
    global file_path
    path = filedialog.askopenfilename(title="Select workbook", filetypes=[("Excel files", "*.xlsx")])

    if path:
        file_path = path
        file_label.config(text="File is selected", fg="blue")

def update_worklog(file_name, name, project_name, article_name, activity, start_time, end_time, break_time):

    time_format = '%H:%M'

    try:

        start_dt = datetime.strptime(start_time, time_format)
        end_dt = datetime.strptime(end_time, time_format)

        delta = end_dt - start_dt

        if delta.days < 0 :
            delta += timedelta(days=1)
        
        total_hours = delta.total_seconds() / 3600
        net_hours = total_hours - (break_time / 60)

        hours = round(net_hours, 2)

        today = datetime.now().strftime('%Y/%m/%d')
        
        weekday = datetime.now().strftime('%A')

        sheet_name = datetime.now().strftime('%B %Y')
        
        new_entry = {
            'Name': [name],
            'Project Name': [project_name],
            'Date': [today],
            'Weekday': [weekday],
            'Article Name': [article_name],
            'Start Time': [start_time],
            'End Time': [end_time], 
            'Break': [break_time],
            'Activity': [activity],
            'Work Hours': [hours],
            
        }

        df_new = pd.DataFrame(new_entry)


        with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            # If sheet doesn't exist, start_row is 0
            if sheet_name in writer.book.sheetnames:
                start_row = writer.book[sheet_name].max_row
                header = False
            else:
                start_row = 0
                header = True


            df_new.to_excel(writer, index=False, header=header, startrow=start_row, sheet_name=sheet_name)

        return True, f"Log updated in {sheet_name}: {hours} hours added."
    except Exception as e:
        return False, f"An error occurred: {e}"

def enter_data():
    if not file_path:
        messagebox.showwarning("File Missing", "Please select an Excel file first!")
        return
    s_h, s_m = start_hh.get(), start_mm.get()
    e_h, e_m = end_hh.get(), end_mm.get()
    name = name_entry.get()
    project_name = project_entry.get()
    article_name = article_combobox.get()
    activity = activity_entry.get()
    break_time = int(break_combobox.get())

    required_fields = [s_h, s_m, e_h, e_m, name, project_name, activity]

    if any(field == "" for field in required_fields) or break_time is None:
        messagebox.showwarning(title="Error", message="You have not filled all fields required.")
        return
    else:

        try:
            if not (0 <= int(s_h) <= 23 and 0 <= int(e_h) <= 23):
                raise ValueError("Hours must be 00-23")
            if not (0 <= int(s_m) <= 59 and 0 <= int(e_m) <= 59): # Changed e_h to e_m
                raise ValueError("Minutes must be 00-59")
        except ValueError as e:
            messagebox.showwarning("Invalid Time", f"Check your numbers: {e}")

        start_time = f"{s_h.zfill(2)}:{s_m.zfill(2)}"
        end_time = f"{e_h.zfill(2)}:{e_m.zfill(2)}"

        success, message = update_worklog(file_path, name, project_name, article_name, activity, start_time, end_time, break_time)
        # 3. Provide feedback to user
    if success:
        messagebox.showinfo("Done", message)
        # Clear fields

        widgets_to_clear = [
            start_hh, start_mm, 
            end_hh, end_mm, 
            name_entry, 
            project_entry, 
            activity_entry
        ]

        for entry in widgets_to_clear:
            entry.delete(0, tkinter.END)
        
        article_combobox.set('')
        break_combobox.set('0') 

    else:
        messagebox.showerror("Error", message)
    

# main window which holds the widgets
window = tkinter.Tk()
window.title("Work Logger")

# widget
frame = tkinter.Frame(window)
frame.pack()

# File picking

fileframe = tkinter.LabelFrame(frame, text="Choose a workbook file.")
fileframe.grid(row=0, column=0, sticky="ew", padx=20, pady=0)

btn_browse = tkinter.Button(fileframe, text="Browse Excel File", command=openFile)
btn_browse.pack(side="left")

file_label = tkinter.Label(fileframe, text="No Excel file selected", fg="red", padx=10)
file_label.pack(side="left")

# Name entry
name_frame = tkinter.LabelFrame(frame, text="User info")
name_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=0)

# Name Label
nameLabel = tkinter.Label(name_frame, text="Enter Name")
nameLabel.grid(row=0, column=0)
projectLabel = tkinter.Label(name_frame, text="Enter Project Name")
projectLabel.grid(row=0, column=1)
articleLabel = tkinter.Label(name_frame, text="Article Name")
articleLabel.grid(row=3, column=0, columnspan=8,  pady=(10, 0))

name_entry = ttk.Combobox(name_frame, values=["", "Sanjay Palad", "Tahir Palad"])
project_entry = ttk.Combobox(name_frame, values=["", "26003_CELOP26"])
name_entry.grid(row=1, column=0, padx=(10, 0))
project_entry.grid(row=1, column=1)
article_combobox = ttk.Combobox(name_frame, values=["", "ENG_Process Engineer - Sanjay Palad - Onsite", "ENG_Process Engineer - Sanjay Palad - Remote"])
article_combobox.grid(row=4, column=0, columnspan=8, sticky= "ew", padx=(10, 0))

# Save user start
user_info_frame = tkinter.LabelFrame(frame, text="Time entry (24HR format)", padx=10, pady=10)
user_info_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)

# Start
tkinter.Label(user_info_frame, text="Start:").grid(row=0, column=0)
start_hh = tkinter.Spinbox(user_info_frame, from_=0, to=24, increment=1, format="%02.0f", width=3)
start_hh.grid(row=0, column=1)
tkinter.Label(user_info_frame, text=":").grid(row=0, column=2)
start_mm = tkinter.Spinbox(user_info_frame, values=[0, 15, 30, 45], wrap=True, format="%02.0f", width=3)
start_mm.grid(row=0, column=3)

# End
tkinter.Label(user_info_frame, text="End:").grid(row=0, column=4, padx=(50, 0))
end_hh = tkinter.Spinbox(user_info_frame, from_=0, to=24, increment=1, format="%02.0f", width=3)
end_hh.grid(row=0, column=5)
tkinter.Label(user_info_frame, text=":").grid(row=0, column=6)
end_mm = tkinter.Spinbox(user_info_frame, values=[0, 15, 30, 45], wrap=True, format="%02.0f", width=3)
end_mm.grid(row=0, column=7)

tkinter.Label(user_info_frame, text="Break time:").grid(row=1, column=0, pady=(10, 0))
break_combobox = ttk.Combobox(user_info_frame, values=[0, 15, 30, 45, 60])
break_combobox.grid(row=4, column=0, columnspan=8, sticky= "ew", padx=(10, 0))


# Save user activity
activity_frame = tkinter.LabelFrame(frame, text="Activities performed", padx=10, pady=10)
activity_frame.grid(row=3, column=0, sticky="ew", padx=20)
activity_entry = tkinter.Entry(activity_frame)
activity_entry.pack(fill="x", padx=10)

# Button
button = tkinter.Button(frame, text="Log Work" , command= enter_data, bg="#F44336", fg="white")
button.grid(row=4, column=0, sticky="news", padx=20, pady=20)

window.mainloop()