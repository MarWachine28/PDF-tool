import tkinter
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
from datetime import datetime, timedelta

class WorkLoggerApp:
    def __init__(self, window):
        self.window = window
        self.window.title("Work Logger")
        self.file_path = "" 

        # Main container
        self.frame = tkinter.Frame(self.window)
        self.frame.pack()

        self.setup_ui()

    def setup_ui(self):
        # File selection
        self.fileframe = tkinter.LabelFrame(self.frame, text="Choose a workbook file.")
        self.fileframe.grid(row=0, column=0, sticky="ew", padx=20, pady=0)

        self.btn_browse = tkinter.Button(self.fileframe, text="Browse Excel File", command=self.openFile)
        self.btn_browse.pack(side="left")

        self.file_label = tkinter.Label(self.fileframe, text="No Excel file selected", fg="red", padx=10)
        self.file_label.pack(side="left")

        # User info
        self.name_frame = tkinter.LabelFrame(self.frame, text="User info")
        self.name_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=0)

        self.nameLabel = tkinter.Label(self.name_frame, text="Enter Name")
        self.nameLabel.grid(row=0, column=0)
        self.projectLabel = tkinter.Label(self.name_frame, text="Enter Project Name")
        self.projectLabel.grid(row=0, column=1)
        self.articleLabel = tkinter.Label(self.name_frame, text="Article Name")
        self.articleLabel.grid(row=3, column=0, columnspan=8, pady=(10, 0))

        self.name_entry = ttk.Combobox(self.name_frame, values=["", "Sanjay Palad", "Tahir Palad"])
        self.project_entry = ttk.Combobox(self.name_frame, values=["", "26003_CELOP26"])
        self.name_entry.grid(row=1, column=0, padx=(10, 0))
        self.project_entry.grid(row=1, column=1)
        
        self.article_combobox = ttk.Combobox(self.name_frame, values=["", "ENG_Process Engineer - Sanjay Palad - Onsite", "ENG_Process Engineer - Sanjay Palad - Remote"])
        self.article_combobox.grid(row=4, column=0, columnspan=8, sticky="ew", padx=(10, 0))

        # Time input
        self.user_info_frame = tkinter.LabelFrame(self.frame, text="Time entry (24HR format)", padx=10, pady=10)
        self.user_info_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)

        # Start
        tkinter.Label(self.user_info_frame, text="Start:").grid(row=0, column=0)
        self.start_hh = tkinter.Spinbox(self.user_info_frame, from_=0, to=24, increment=1, format="%02.0f", width=3)
        self.start_hh.grid(row=0, column=1)
        tkinter.Label(self.user_info_frame, text=":").grid(row=0, column=2)
        self.start_mm = tkinter.Spinbox(self.user_info_frame, values=[0, 15, 30, 45], wrap=True, format="%02.0f", width=3)
        self.start_mm.grid(row=0, column=3)

        # End
        tkinter.Label(self.user_info_frame, text="End:").grid(row=0, column=4, padx=(50, 0))
        self.end_hh = tkinter.Spinbox(self.user_info_frame, from_=0, to=24, increment=1, format="%02.0f", width=3)
        self.end_hh.grid(row=0, column=5)
        tkinter.Label(self.user_info_frame, text=":").grid(row=0, column=6)
        self.end_mm = tkinter.Spinbox(self.user_info_frame, values=[0, 15, 30, 45], wrap=True, format="%02.0f", width=3)
        self.end_mm.grid(row=0, column=7)

        tkinter.Label(self.user_info_frame, text="Break time:").grid(row=1, column=0, pady=(10, 0))
        self.break_combobox = ttk.Combobox(self.user_info_frame, values=[0, 15, 30, 45, 60])
        self.break_combobox.grid(row=4, column=0, columnspan=8, sticky="ew", padx=(10, 0))

        # --- Activities ---
        self.activity_frame = tkinter.LabelFrame(self.frame, text="Activities performed", padx=10, pady=10)
        self.activity_frame.grid(row=3, column=0, sticky="ew", padx=20)
        self.activity_entry = tkinter.Entry(self.activity_frame)
        self.activity_entry.pack(fill="x", padx=10)

        # --- Log Button ---
        self.button = tkinter.Button(self.frame, text="Log Work", command=self.enter_data, bg="#F44336", fg="white")
        self.button.grid(row=4, column=0, sticky="news", padx=20, pady=20)

    def openFile(self):
        path = filedialog.askopenfilename(title="Select workbook", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.file_path = path
            self.file_label.config(text="File is selected", fg="blue")

    def update_worklog(self, file_name, name, project_name, article_name, activity, start_time, end_time, break_time):
        time_format = '%H:%M'
        try:
            start_dt = datetime.strptime(start_time, time_format)
            end_dt = datetime.strptime(end_time, time_format)
            delta = end_dt - start_dt

            if delta.days < 0:
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

    def enter_data(self):
        if not self.file_path:
            messagebox.showwarning("File Missing", "Please select an Excel file first!")
            return

        s_h, s_m = self.start_hh.get(), self.start_mm.get()
        e_h, e_m = self.end_hh.get(), self.end_mm.get()
        name = self.name_entry.get()
        project_name = self.project_entry.get()
        article_name = self.article_combobox.get()
        activity = self.activity_entry.get()
        
        try:
            break_val = self.break_combobox.get()
            break_time = int(break_val) if break_val else 0
        except ValueError:
            messagebox.showwarning("Error", "Break time must be a number.")
            return

        required_fields = [s_h, s_m, e_h, e_m, name, project_name, activity]

        if any(field == "" for field in required_fields):
            messagebox.showwarning(title="Error", message="You have not filled all fields required.")
            return

        try:
            if not (0 <= int(s_h) <= 23 and 0 <= int(e_h) <= 23):
                raise ValueError("Hours must be 00-23")
            if not (0 <= int(s_m) <= 59 and 0 <= int(e_m) <= 59):
                raise ValueError("Minutes must be 00-59")
        except ValueError as e:
            messagebox.showwarning("Invalid Time", f"Check your numbers: {e}")
            return

        start_time = f"{s_h.zfill(2)}:{s_m.zfill(2)}"
        end_time = f"{e_h.zfill(2)}:{e_m.zfill(2)}"

        success, message = self.update_worklog(self.file_path, name, project_name, article_name, activity, start_time, end_time, break_time)

        if success:
            messagebox.showinfo("Done", message)
            
            # Clear fields
            widgets_to_clear = [
                self.start_hh, self.start_mm, 
                self.end_hh, self.end_mm, 
                self.name_entry, 
                self.project_entry, 
                self.activity_entry
            ]

            for entry in widgets_to_clear:
                if isinstance(entry, tkinter.Entry) or isinstance(entry, ttk.Combobox):
                    entry.delete(0, tkinter.END)
                elif isinstance(entry, tkinter.Spinbox):
                    entry.delete(0, tkinter.END)
                    entry.insert(0, "00")
            
            self.article_combobox.set('')
            self.break_combobox.set('0') 
        else:
            messagebox.showerror("Error", message)

if __name__ == "__main__":
    root = tkinter.Tk()
    app = WorkLoggerApp(root)
    root.mainloop()