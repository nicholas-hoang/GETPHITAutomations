# Modules: Email_Parser and PDF_Converter
from email.policy import default
from email.parser import BytesParser
import glob
import pdfkit
import os
from bs4 import BeautifulSoup
import datetime
from pathlib import Path

# Modules: Registration List and Folder Creator
import os
import pdfquery
import pandas as pd

# Modules: Folder Creator
import os
import pandas as pd
import numpy as np
import shutil
import pathlib
import xlwings as xw

import glob
import os
import datetime
import pdfkit
from email.parser import BytesParser
from email.policy import default
from bs4 import BeautifulSoup


class RegistrationFormPDFGenerator:
    def __init__(self, student_files):
        self.student_files = student_files
        self.student_dict = {}

    def extract_file_url(self):
        for student_file in self.student_files:
            with open(student_file, 'rb') as fp:
                raw_message = BytesParser(policy=default).parse(fp)
                soup = BeautifulSoup(raw_message.get_body(preferencelist=('html')).get_content(), 'html.parser')
                registration_link = None
                try:
                    registration_link = soup.find_all('a')[1].get('href')
                    self.student_dict[student_file] = [registration_link]

                    time_submitted = raw_message['Received'].split(';')[1].strip()
                    time_submitted = datetime.datetime.strptime(time_submitted, '%a, %d %b %Y %H:%M:%S %z')
                    time_submitted = time_submitted.strftime('%m-%d-%Y %H:%M:%S')
                    self.student_dict[student_file].append(time_submitted)

                except TypeError:
                    print(f"No URL found in {student_file}")
                    continue

        return self.student_dict

    def generate_pdfs(self):
        count = 0
        options = {'orientation': 'Landscape'}
        for student_file, registration_link in self.student_dict.items():
            out_file = student_file.replace('.eml', '.pdf')
            pdfkit.from_url(registration_link[0], out_file, options=options)
            count += 1

        print('Done!')
        print(f'A total of {count} pdfs were generated.')

    def move_pdfs(self):
        os.mkdir("Generated PDFs")
        for pdf_file in glob.glob("*.pdf"):
            os.rename(pdf_file, "Generated PDFs/" + pdf_file)

        print('Done!')
        print('All PDFs have been moved to the Generated PDFs folder.')

#
import os
import pandas as pd
import numpy as np
import shutil
import pathlib
import xlwings as xw
import tkinter as tk
from tkinter import filedialog

class Solution:
    def __init__ (self):
        self.file = None
        self.evaluation = None

    def select_registration_file(self):
        self.file = filedialog.askopenfilename(initialdir = ".", title = "Select Registration File",
                                               filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))

    def select_evaluation_file(self):
        self.evaluation = filedialog.askopenfilename(initialdir = ".", title = "Select Evaluation File",
                                                     filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))

    def load_data(self):
        if self.file is not None and self.evaluation is not None:
            student_registration = pd.read_excel(self.file)
            student_registration = student_registration.rename(columns={
                'Registration Date': 'Registration_Date',
                'First Name': 'First_Name',
                'Last Name': 'Last_Name'
            })
            student_registration.Registration_Date = student_registration.Registration_Date.dt.strftime('%m/%d/%Y')

            with open(self.evaluation, 'r') as f:
                evaluation_form = pd.read_excel(f)

            return student_registration, evaluation_form

        else:
            print("Please select files")
            return None, None

    def get_full_name(self, row):
        fullname = row.First_Name + "_" + row.Last_Name
        return fullname

    def load_registration_file(self):
        student_registration = pd.read_excel(self.file)
        student_registration = student_registration.rename(columns = {'Registration Date':'Registration_Date',
        'First Name':'First_Name','Last Name':'Last_Name'})
        student_registration.Registration_Date = student_registration.Registration_Date.dt.strftime('%m/%d/%Y')

        return student_registration

    def get_first_last_name(self, row):
        fullname = row.First_Name + " " + row.Last_Name
        return fullname

    def get_existing_folders(self):
        num_of_docs = [directory for directory in os.listdir('.') if os.path.isdir(os.path.join(".", directory))]
        print(f'There are a total of {len(num_of_docs)} student folders in this directory.')
        return num_of_docs

    def create_new_folders(self, df):
        count = 0
        student_names = pd.Series(df.full_name)

        for student in student_names:
            os.mkdir(student.title()) # Create Student Folder
            count +=1

        print(f'A total of {count} student folders were generated.')

    def fill_folders(self, source):
        path = pathlib.Path('.')
        for item in path.iterdir():
            if item.is_dir():
                shutil.copy2(source, item) # Place Form in Folder

    def rename_files(self):
        path = pathlib.Path('.')
        for folder in path.iterdir():
            if folder.is_dir():
                for file in folder.iterdir():
                    if file.suffix == ".xlsx":
                        os.rename(file, folder / (folder.name + ".xlsx"))
                    elif file.suffix == ".pdf":
                        os.rename(file, folder / (folder.name + ".pdf"))

    def run(self):
        root = tk.Tk()
        root.title("File Selection")

        # Registration file button
        reg_file_button = tk.Button(root, text="Select Registration File", command=self.select_registration_file)
        reg_file_button.pack(pady=10)

        # Evaluation file button
        eval_file_button = tk.Button(root, text="Select Evaluation File", command=self.select_evaluation_file)
        eval_file_button.pack(pady=

class 




# Example usage:
generator = RegistrationFormPDFGenerator(glob.glob('*.eml'))
name_dict = generator.extract_file_url()
generator.generate_pdfs()
generator.move_pdfs()