from datetime import datetime
import pandas as pd
import pandasql as ps
import customtkinter
from tkinter import filedialog, messagebox
import os
import glob
import win32com.client

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()
app.title("Excel Parser Tool")

global outdir
outdir = 'time_card_output'

global outdir_with_backslash
outdir_with_backslash = r"\time_card_output"

if not os.path.exists(outdir):
    os.mkdir(outdir)
def select_excel_file():
    try:
        global excel_path
        excel_path = filedialog.askopenfilename()
        if (len(excel_path) > 1 and excel_path.split(".")[1] == "xlsx"):
            messagebox.showinfo("File Message", "Excel File Uploaded Successfully")
        else:
            messagebox.showinfo("File Message", "Selected file is not in Excel Format")
    except Exception as e:
        messagebox.showinfo("File Error", "No file Selected")

def extract_data():
    try:
        df = pd.read_excel(excel_path, sheet_name='Data')
        df.columns = df.columns.str.replace(' ', '', regex=True)
        df.columns = df.columns.str.replace('.', '', regex=True)

        output = []
        super_name_list = []
        supervisor_list = []


        q1 = """select distinct(SupervisorName) from df"""
        output = []
        output = ps.sqldf(q1, locals())

        super_list = []
        file_super = []
        temp = []
        if len(lines) > 0:
            for items in lines:
                super_list.append(items.replace(",", "").replace("\n", "").split(" "))

            for items in super_list:
                temp = items
                file_super.append(sorted(temp))

                temp = []
                df_super = []
                global super
                super = []

                for supervisor in output.SupervisorName:
                    for supervisor1 in file_super:
                        temp = supervisor.replace(",", "").split(" ")
                        if sorted(temp) in file_super:
                            if supervisor not in super:
                                super.append(supervisor)

            if len(super) > 0:
                print(super)
                current_datetime = datetime.now().strftime("%Y-%m-%d")
                global outname
                outname = "pending_timecard_details_" + current_datetime + ".xlsx"
                if os.path.exists(f"{outdir}/{outname}"):
                    os.remove(f"{outdir}/{outname}")
                writer = pd.ExcelWriter(f"{outdir}/{outname}", engine='xlsxwriter')
                for supervisor in super:
                    q2 = f"""select EmpID, EmployeeName, SupervisorName, Status, TSStatus, Period, WeekEnding, Employee_Email from df where SupervisorName = '{supervisor}'"""
                    supervisor_sql = ps.sqldf(q2, locals())
                    supervisor_sql.sort_values(by = ['EmployeeName', "WeekEnding"], ascending=True, inplace=True)
                    supervisor_sql['EmployeeName'] = supervisor_sql['EmployeeName'].str.replace('.', '',
                                                                                            regex=True)
                    supervisor_sql['EmployeeName'] = supervisor_sql['EmployeeName'].str.replace(',', '',
                                                                                            regex=True)
                    supervisor_sql.rename(columns={"TSStatus": "N+1 TS Status"}, inplace=True)
                    supervisor_sql['N+1 TS Status'] = supervisor_sql['N+1 TS Status'].str.replace('Non Submitter','', regex = True)
                    supervisor_sql['WeekEnding'] = supervisor_sql['WeekEnding'].str.replace('00:00:00.000000', '', regex=True)
                    supervisor_sql.to_excel(writer, sheet_name=supervisor[0:30], index=False)

                    # Auto-adjust columns' width
                    for column in supervisor_sql:
                        column_width = max(supervisor_sql[column].astype(str).map(len).max(), len(column))
                        col_idx = supervisor_sql.columns.get_loc(column)
                        writer.sheets[supervisor[0:30]].set_column(col_idx, col_idx, column_width)

                writer.close()
                global excel_file_path
                excel_file_path = os.path.abspath(f"{outdir}/{outname}")
                print(excel_file_path)
                messagebox.showinfo("Data Extraction Message", "Data Extracted Successfully!")
                folder_path = os.getcwd() + outdir_with_backslash
                file_type = r'\*xlsx'
                files = glob.glob(folder_path + file_type)
                max_file = max(files, key=os.path.getctime)
                os.startfile(max_file)
            else:
                messagebox.showinfo("Data Extraction Error", "File not created. Supervisor names does not match the Text file.")
        else:
            messagebox.showinfo("File Error",
                                "Text file is empty")
    except Exception as e:
        messagebox.showinfo("Data Extraction Error", f"No Data Extracted {e}")

def select_text_file():
    try:
        global text_path
        text_path = filedialog.askopenfilename()
        lst = []
        if len(text_path) > 1 and text_path.split(".")[1] == "txt":
            with open(text_path) as f:
                global lines
                lines = f.readlines()
                for items in lines:
                    lst.append(items.isspace())
                print(lst)
                if all(item is True for item in lst):
                    messagebox.showinfo("File Error", "File does not contain any data!")
                else:
                    messagebox.showinfo("File Input", "File Uploaded Successfully!")
        else:
            messagebox.showinfo("File Input", "Selected file is not in Text format")
    except Exception as e:
        messagebox.showinfo("Text File Error", "No File Uploaded")

def open_folder():
    if not os.path.exists(outdir):
        os.mkdir(outdir)
    folder_path = os.getcwd() + outdir_with_backslash
    os.startfile(folder_path)

def open_latest_file():
    try:
        folder_path = os.getcwd() + outdir_with_backslash
        file_type = r'\*xlsx'
        files = glob.glob(folder_path + file_type)
        max_file = max(files, key=os.path.getctime)
        os.startfile(max_file)
    except Exception as e:
        messagebox.showinfo("File Error", "No Extracted File Found!")

def close():
    app.destroy()

def send_mail_click():
    for superviser_names in super:
        send_mail(superviser_names)

def send_mail(name):

    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Timecard Not Submitted'
    # name = "Upadhye, Manoj"
    # newmail.To = name
    newmail.To = "Jitendra Dang, Siddharth"
    newmail.CC = 'Upadhye, Manoj'

    newmail.Body = f'Hello {name},' \
                   f'\n\nPlease refer the attached excel file for timecard not compliance.' \
                   f'\ncheck the sheet for your name and act accordingly' \
                   f'\n\n\nThanks' \

    attach = excel_file_path

    newmail.Attachments.Add(attach)
    newmail.Send()

frame_1 = customtkinter.CTkFrame(master=app)
frame_1.pack(pady=0, padx=0)

button_1 = customtkinter.CTkButton(master=frame_1, text="Select Excel File", command=select_excel_file)
button_1.grid(pady=(50,20), padx=(30,30))

button_1 = customtkinter.CTkButton(master=frame_1, text="Select Supervisor Text File", command=select_text_file)
button_1.grid(pady=(20,20), padx=(30,30))

button_1 = customtkinter.CTkButton(master=frame_1, text="Extract Data", command=extract_data)
button_1.grid(pady=(20,50), padx=(30,30))

button_1 = customtkinter.CTkButton(master=frame_1, text="Send Mail", command=send_mail_click)
button_1.grid(pady=(0,50), padx=(30,30))
#
# button_1 = customtkinter.CTkButton(master=frame_1, text="Open Latest File", command=open_latest_file)
# button_1.grid(pady=20, padx=(30,30))
#
# button_1 = customtkinter.CTkButton(master=frame_1, text="Close", command=close)
# button_1.grid(pady=(20,50), padx=(30,30))

app.mainloop()