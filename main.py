import os
import numpy as np
import pandas as pd
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.converter import TextConverter
from pdfminer3.layout import LAParams
from pdfminer3.pdfpage import PDFPage
import io
from tkinter import filedialog
import tkinter as tk
import tkinter.ttk as ttk
import time
from datetime import datetime
import concurrent.futures
from multiprocessing import freeze_support

# Pandas Print Display Options
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', None)
pd.set_option('display.width', None)


class MasterFrame:
    def __init__(self, root):
        self.root = root
        self.export_path = tk.StringVar()
        self.import_files = []

        # Root Options
        root.title('Director Report Batch Converter')
        root.geometry('900x500')
        root.minsize(width=900, height=500)
        root.rowconfigure(1, weight=1)
        root.columnconfigure(1, weight=1)

        # region ========== 1 - Directory Path Selection ==========
        self.fp_frame = tk.LabelFrame(root, text='Select Import Directory')
        self.fp_frame.grid(row=0, column=0, columnspan=3, sticky='NEW')
        self.fp_frame.columnconfigure(1, weight=1)

        # Filepath Selection
        self.fp_button = tk.Button(self.fp_frame, text='Open:', command=self.directory_path)
        self.fp_button.grid(row=0, column=0)
        self.fp_label = tk.Label(self.fp_frame, textvariable=self.export_path, anchor='w')
        self.fp_label.grid(row=0, column=1, sticky='EW')
        # endregion

        # region ========== 2 - Directory Contents List ==========
        self.contents_frame = tk.LabelFrame(root, text='File Type List')
        self.contents_frame.grid(row=1, column=0, sticky='NSEW')
        self.contents_frame.columnconfigure(0, weight=2)
        self.contents_frame.columnconfigure(1, weight=2)
        self.contents_frame.rowconfigure(1, weight=1)

        # PDF Listbox
        self.pdf_label = tk.Label(self.contents_frame, text=' . . . ')
        self.pdf_label.grid(row=0, column=0)
        self.pdf_listbox = tk.Listbox(self.contents_frame)
        self.pdf_listbox.grid(row=1, column=0, sticky='NSEW')
        # Other Listbox
        self.other_label = tk.Label(self.contents_frame, text=' . . . ')
        self.other_label.grid(row=0, column=1)
        self.other_listbox = tk.Listbox(self.contents_frame)
        self.other_listbox.grid(row=1, column=1, sticky='NSEW')
        # endregion

        # region ========== 3 - Filtering Options =========
        self.options_frame = tk.LabelFrame(root, text='Options')
        self.options_frame.grid(row=2, column=0, sticky='NSEW')
        self.options_frame.columnconfigure(1, weight=1)

        # Filter by date
        self.date_label = tk.Label(self.options_frame, text='Filter by appointment & resignation dates:')
        self.date_label.grid(row=0, column=0, sticky='W')
        self.date_button = tk.Button(self.options_frame, text='No', command=self.date_on_off)
        self.date_button.grid(row=0, column=2, sticky='EW')
        # Year Start
        self.ys_label = tk.Label(self.options_frame, text='Year Start (DD/MM/YYYY):')
        self.ys_label.grid(row=1, column=0)
        self.ys = tk.Entry(self.options_frame, state='disabled', justify='center')
        self.ys.grid(row=1, column=2)
        # Year End
        self.ye_label = tk.Label(self.options_frame, text='Year End (DD/MM/YYYY):')
        self.ye_label.grid(row=2, column=0)
        self.ye = tk.Entry(self.options_frame, state='disabled', justify='center')
        self.ye.grid(row=2, column=2)

        # Filter by role
        self.role_label = tk.Label(self.options_frame, text='Filter out non-Director/LLP key roles:')
        self.role_label.grid(row=3, column=0, sticky='W')
        self.role_button = tk.Button(self.options_frame, text='No', command=lambda: self.on_off(self.role_button))
        self.role_button.grid(row=3, column=2, sticky='EW')

        # Filter by Status
        self.status_label = tk.Label(self.options_frame, text='Filter out companies if not active:')
        self.status_label.grid(row=4, column=0, sticky='W')
        self.status_button = tk.Button(self.options_frame, text='No', command=lambda: self.on_off(self.status_button))
        self.status_button.grid(row=4, column=2, sticky='EW')

        # Rename PDF
        self.rename_label = tk.Label(self.options_frame, text='Rename PDF Files to name of director:')
        self.rename_label.grid(row=5, column=0, sticky='W')
        self.rename_button = tk.Button(self.options_frame, text='No', command=lambda: self.on_off(self.rename_button))
        self.rename_button.grid(row=5, column=2, sticky='EW')
        # endregion

        # region ========== 4 - Progress Frame ==========
        self.scrollbar = tk.Scrollbar(root, orient='vertical')
        self.scrollbar.grid(row=1, rowspan=2, column=2, sticky='NS')

        # Output Listbox
        self.output_listbox = tk.Listbox(root, yscrollcommand=self.scrollbar.set)
        self.output_listbox.grid(row=1, rowspan=2, column=1, sticky='NSEW')
        self.output_listbox.columnconfigure(0, weight=1)
        self.output_listbox.rowconfigure(0, weight=1)
        self.scrollbar['command'] = self.output_listbox.yview
        # endregion

        # region ========== 5 - Convert Button ==========
        self.convert_frame = tk.Frame(root)
        self.convert_frame.grid(row=4, column=0, columnspan=3, sticky='NSEW')
        self.convert_frame.columnconfigure(1, weight=1)

        # Convert Button & Progress Bar
        self.ep_convert_button = tk.Button(self.convert_frame, text='Begin PDF-CSV File Conversion:',
                                           command=self.convert,
                                           state='disabled')
        self.ep_convert_button.grid(row=0, column=0, sticky='NS')
        self.total_progress = ttk.Progressbar(self.convert_frame, orient='horizontal', length=200, mode='determinate')
        self.total_progress.grid(row=0, column=1, sticky='NSEW')
        # endregion

    def directory_path(self):
        # 1) Clear GUI & Import List
        self.import_files.clear()
        self.pdf_listbox.delete(0, tk.END)
        self.other_listbox.delete(0, tk.END)

        # 2) Select import directory
        directory = tk.filedialog.askdirectory(initialdir='/', title='Select a directory')

        # 3) Check FP Selection is Valid
        try:
            if directory == '':
                self.export_path.set('ERROR: No import directory selected')
                self.ep_convert_button.config(state='disabled')
            else:
                # 4) Change directory
                self.export_path.set(str(directory))
                os.chdir(directory)

                # 5) Create list of files with .pdf extension and exceptions
                for file in os.listdir(directory):
                    if file.endswith('.pdf'):
                        self.pdf_listbox.insert(tk.END, file)
                        self.pdf_listbox.itemconfig(tk.END, fg='#34eb37')
                        self.import_files.append(file)
                    else:
                        self.other_listbox.insert(tk.END, file)
                        self.other_listbox.itemconfig(tk.END, fg='red')

                # 6) Update GUI with list of files
                self.pdf_label['text'] = 'PDF file count: ' + str(self.pdf_listbox.size())
                self.other_label['text'] = 'Other file count: ' + str(self.other_listbox.size())
                self.ep_convert_button.config(state='normal')
        except Exception as error:
            self.export_path.set('ERROR:' + str(error))
            self.ep_convert_button.config(state='disabled')

    @staticmethod
    def on_off(button):
        # Selects GUI buttons on and off
        button_text = 'Yes' if button.config('text')[-1] == 'No' else 'No'
        button_relief = 'raised' if button.config('relief')[-1] == 'sunken' else 'sunken'
        button.config(text=button_text, relief=button_relief)

    def date_on_off(self):
        # Turn button off and disable entries
        if self.date_button.config('text')[-1] == 'Yes':
            self.date_button.config(text='No', relief='raised')
            self.ys.config(state='disabled')
            self.ye.config(state='disabled')
        # Turn button on and enable entries
        else:
            self.date_button.config(text='Yes', relief='sunken')
            self.ys.config(state='normal')
            self.ye.config(state='normal')

    def convert(self):
        # 1) Clear Output Listbox & Reset Progress Bar
        self.output_listbox.delete(0, tk.END)
        self.total_progress['maximum'] = len(self.import_files)
        self.total_progress['value'] = 0

        # 2) Filepath Error Handling
        try:
            os.chdir(self.export_path.get())
        except Exception as error:
            self.output_listbox.insert(tk.END, str(error))
            self.output_listbox.itemconfig(tk.END, fg='red')
            return

        # 3) Date Time Error Handling
        if self.date_button['text'] == 'Yes':
            try:
                datetime.strptime(self.ys.get(), '%d/%m/%Y')
                datetime.strptime(self.ye.get(), '%d/%m/%Y')
            except ValueError:
                self.output_listbox.insert(tk.END, 'Date format on filter does not match DD/MM/YYYY format.')
                self.output_listbox.itemconfig(tk.END, fg='red')
                return

        # 4) Reset Export Dataframe, Error List and Start Load Timer
        export_data = pd.DataFrame()
        start = time.time()
        error_list = []

        # 5) Begin Concurrency Execution
        with concurrent.futures.ProcessPoolExecutor() as executor:
            if __name__ == '__main__':
                # 6) Loop Through PDF Import list Files
                multiprocess_files = [executor.submit(pdf_text, pdf) for pdf in self.import_files]

                # 7) Reset Variables Used for GUI Output
                index = 0
                raw_lines_total = 0
                filtered_lines_total = 0

                # 8) Filter Each PDF Data File as Completed
                for raw_data in concurrent.futures.as_completed(multiprocess_files):
                    index += 1

                    # 9) Assign function returns to variables
                    file_name, data, director_name = raw_data.result()
                    raw_lines = str(len(data))
                    raw_lines_total += len(data)

                    if not data.empty:
                        # 10) Filter by date range
                        if self.date_button['text'] == 'Yes':
                            start_datetime = datetime.strptime(self.ys.get(), '%d/%m/%Y')
                            end_datetime = datetime.strptime(self.ye.get(), '%d/%m/%Y')
                            data = data[(data['Appoint. Date'].fillna('1900-01-01') <= end_datetime) &
                                        (data['Res. Date'].fillna('2100-01-01') >= start_datetime)]

                        # 11) Filter by Role
                        if self.role_button['text'] == 'Yes':
                            key_roles = ['Director', 'LLP Designated Member', 'LLP Member']
                            data = data[data['Appoint. Type'].isin(key_roles)]

                        # 12) Filter by Status
                        if self.status_button['text'] == 'Yes':
                            data = data[data['Status'] == 'Active']

                        # 13) Rename pdf
                        if self.rename_button['text'] == 'Yes':
                            try:
                                os.rename(file_name, (director_name + '.pdf'))
                            except Exception as error:
                                self.output_listbox.insert(tk.END, str(error))
                                self.output_listbox.itemconfig(tk.END, fg='red')

                        # 14) Set Variable for Companies Filtered
                        line_exports = str(len(data)) + '/' + str(raw_lines) + ' lines exported.'
                        filtered_lines_total += len(data)
                    else:
                        line_exports = ' ERROR: PDF is Non-Standard Format'
                        error_list.append(str(index) + " - " + str(file_name))

                    # 15) GUI Output for PDF File
                    self.output_listbox.insert(tk.END, str(index) + ' - ' + str(file_name) + ': ' + line_exports)
                    fg_colour = 'red' if line_exports.startswith(' ERROR:') \
                        else '#34eb37' if len(data) != 0 \
                        else 'yellow'
                    self.output_listbox.itemconfig(tk.END, fg=fg_colour)
                    self.output_listbox.insert(tk.END, '--------------------')
                    self.output_listbox.update_idletasks()
                    self.total_progress['value'] += 1

                    export_data = pd.concat([export_data, data])

        # 16) Reset Datafram Index When Loop is Completed
        export_data.reset_index(inplace=True, drop=True)

        # 17) Write to Excel Files
        with pd.ExcelWriter('Consolidated Director Report.xlsx') as writer:
            export_data.to_excel(writer, sheet_name='Director Reports', index=False)

        # 18) Update GUI with Summary of Data
        self.output_listbox.insert(tk.END, '')
        self.output_listbox.insert(tk.END, 'Total Companies:')
        self.output_listbox.insert(tk.END, str(filtered_lines_total) + "/" + str(raw_lines_total)
                                   + ' companies were converted after filtering.')
        total_colour = '#34eb37' if filtered_lines_total > 0 else 'red'
        self.output_listbox.itemconfig(tk.END, fg=total_colour)
        self.output_listbox.insert(tk.END, "")

        # GUI Update on PDF Renaming
        if self.rename_button['text'] == 'Yes':
            self.output_listbox.insert(tk.END, 'PDFs have been renamed to reflect their director:')
            self.output_listbox.insert(tk.END, 'Conversion button disabled due to original file name dependency.')
            self.output_listbox.itemconfig(tk.END, fg='yellow')
            self.output_listbox.insert(tk.END, 'Please reselect import directory to re-enable.')
            self.output_listbox.itemconfig(tk.END, fg='yellow')
            self.output_listbox.insert(tk.END, '')
            self.ep_convert_button.config(state='disabled')

        # GUI Update on Errors Encountered
        if len(error_list) > 0:
            self.output_listbox.insert(tk.END, 'The following Files had import errors:')
            for file in error_list:
                self.output_listbox.insert(tk.END, file)
                self.output_listbox.itemconfig(tk.END, fg='red')
            self.output_listbox.insert(tk.END, '')

        # GUI Update on Time Taken
        end = time.time()
        self.output_listbox.insert(tk.END, 'Conversion Time: ' + str(round((end - start), 3)) + ' Seconds')
        self.output_listbox.update_idletasks()


def pdf_text(pdf):
    # 1) Set Import Parameters
    string_io = io.StringIO()
    la_params = LAParams(line_overlap=0.1, char_margin=40, line_margin=0.1, word_margin=0.1, boxes_flow=1.0)
    device = TextConverter(PDFResourceManager(), string_io, codec='utf-8', laparams=la_params)
    interpreter = PDFPageInterpreter(PDFResourceManager(), device)

    # 2) Open PDF file
    fp = open(pdf, 'rb')

    # 3) Get Total Pages in PDF
    page_list = [page for page in PDFPage.get_pages(fp, check_extractable=True)]

    # 4) Extract all pages except first and last
    for page in page_list[1:-1]:
        interpreter.process_page(page)

    # 5) Convert pdf text to list of lines, removing line breaks
    pdf_data = string_io.getvalue()
    lines = list(filter(None, pdf_data.splitlines()))

    # 6) Check files match director report format
    try:
        if lines[0] == 'Onboard':
            # 7) Director Name should always be at this position
            director_name = lines[2]

            # 8) Organises relevant info into lists & purges substrings NOTE: Consider rewriting, seems convoluted
            company = [x.replace('Company Name : ', '') for x in lines if x.startswith('Company Name')]
            number = [x.replace('Company Number:  ', '') for x in lines if x.startswith('Company Number')]
            status = [x.replace('Company Status:  ', '') for x in lines if x.startswith('Company Status')]
            app_type = [x.replace('Appointment Type:  ', '') for x in lines if x.startswith('Appointment Type')]
            app_date = [x.replace('Appointment Date:  ', '') for x in lines if x.startswith('Appointment Date')]
            res_date = [x.replace('Resignation Date:  ', '') for x in lines if x.startswith('Resignation Date')]

            # 9) Combine lists into Dataframe
            data = pd.DataFrame(list(zip(company, number, status, app_type, app_date, res_date)),
                                columns=['Company', 'No.', 'Status', 'Appoint. Type', 'Appoint. Date', 'Res. Date'])

            # 10) Add Director Name as independent column
            data.insert(0, 'Name', director_name)

            # 11) Date Standardisation for Filtering
            data['Appoint. Date'].replace('PRE12/01/1992', '12/01/1992', inplace=True)
            data['Appoint. Date'] = pd.to_datetime(data['Appoint. Date'], format="%d/%m/%Y", errors='coerce')
            data['Res. Date'].replace('-', np.NaN, inplace=True)
            data['Res. Date'] = pd.to_datetime(data['Res. Date'], format="%d/%m/%Y", errors='coerce')
        else:
            data = pd.DataFrame()
            director_name = pdf
    except Exception as error:
        data = pd.DataFrame()
        director_name = pdf
        print(error)

    fp.close()
    return pdf, data, director_name


def colour(parent):
    # 1) loop through each child of widget to set theme
    for child in parent.winfo_children():
        widget_type = child.winfo_class()
        if widget_type == 'Frame':
            child.config(bg='#282828')
        if widget_type == 'Labelframe':
            child.config(bg='#282828', fg='#FFFFFF')
        if widget_type == 'Label':
            child.config(bg='#282828', fg='#FFFFFF')
        if widget_type == 'Button':
            child.config(bg='#404040', fg='#FFFFFF')
        if widget_type == 'Canvas':
            child.config(bg='#404040')
        if widget_type == 'Entry':
            child.config(bg='#D7D7D7', disabledbackground='#282828')
        if widget_type == 'Listbox':
            child.config(bg='#282828', fg='#FFFFFF')
        else:
            colour(child)


if __name__ == "__main__":
    freeze_support()
    master = tk.Tk()

    MasterFrame(master)
    colour(master)
    logo = tk.PhotoImage(file='bdo logo.png')
    master.iconphoto(False, logo)

    master.mainloop()
