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

#GUI Modules
import os
import pandas as pd
import numpy as np
import shutil
import pathlib
import xlwings as xw
import tkinter as tk
from tkinter import filedialog

import os
import pandas as pd
import shutil
import pathlib
import tkinter as tk
from tkinter import filedialog

# Student Grading Evaluation Form
import glob
import pathlib
import os
import pandas as pd
import numpy as np


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

# GUI for RegistrationFormPDFGenerator

class Solution:
    def __init__(self):
        self.file = None
        self.evaluation = None

    def load_data(self):
        if self.file is not None and self.evaluation is not None:
            student_registration = pd.read_excel(self.file)
            student_registration = student_registration.rename(
                columns={
                    'Registration Date': 'Registration_Date',
                    'First Name': 'First_Name',
                    'Last Name': 'Last_Name'
                }
            )
            student_registration.Registration_Date = student_registration.Registration_Date.dt.strftime('%m/%d/%Y')
            evaluation_form = pd.read_excel(self.evaluation)

            return student_registration, evaluation_form
        else:
            print("Please select files")
            return None, None

    def select_registration_file(self):
        self.file = filedialog.askopenfilename(
            initialdir=".",
            title="Select Registration File",
            filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*"))
        )

    def select_evaluation_file(self):
        self.evaluation = filedialog.askopenfilename(
            initialdir=".",
            title="Select Evaluation File",
            filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*"))
        )

    def create_student_folders(self):
        student_registration, _ = self.load_data()
        if student_registration is not None:
            for _, row in student_registration.iterrows():
                folder_name = f"{row['First_Name']}_{row['Last_Name']}"
                os.mkdir(folder_name.title())
        else:
            print("Failed to load data.")

    def copy_evaluation_form(self):
        _, evaluation_form = self.load_data()
        if evaluation_form is not None:
            for folder_name in os.listdir('.'):
                if os.path.isdir(folder_name):
                    shutil.copy2(evaluation_form, folder_name)
        else:
            print("Failed to load data.")

    def rename_files(self):
        for folder_name in os.listdir('.'):
            if os.path.isdir(folder_name):
                folder_path = pathlib.Path(folder_name)
                for file in folder_path.iterdir():
                    if file.name.endswith('.xlsx'):
                        new_file_name = f"{folder_name}.xlsx"
                        file.rename(new_file_name)

    def run(self):
        root = tk.Tk()
        root.title("File Selection")

        # Registration file button
        reg_file_button = tk.Button(root, text="Select Registration File", command=self.select_registration_file)
        reg_file_button.pack(pady=10)

        # Evaluation file button
        eval_file_button = tk.Button(root, text="Select Evaluation File", command=self.select_evaluation_file)
        eval_file_button.pack(pady=10)

        # Load files button
        load_files_button = tk.Button(root, text="Load Files", command=self.load_data)
        load_files_button.pack(pady=10)

        # Create folders button
        create_folders_button = tk.Button(root, text="Create Folders", command=self.create_student_folders)
        create_folders_button.pack(pady=10)

        # Copy evaluation form button
        copy_form_button = tk.Button(root, text="Copy Evaluation Form", command=self.copy_evaluation_form)
        copy_form_button.pack(pady=10)

        # Rename files button
        rename_files_button = tk.Button(root, text="Rename Files", command=self.rename_files)
        rename_files_button.pack(pady=10)

        root.mainloop()



class StudentEvaluations:
    def __init__(self):
        self.files = []
        self.student_evals = pd.DataFrame(columns=['Student Name', 'Grader 1', 'Grader 2', 'Score 1', 'Score 2'])

    def prompt_user_for_folder(self):
        folder_path = input("Please enter the path to the folder containing the student folders: ")
        path = pathlib.Path(folder_path)

        for folder in path.iterdir():
            if folder.is_dir():
                for file in folder.iterdir():
                    if file.is_file() and file.suffix == '.xlsx':
                        self.files.append(file)

    def write_scores(self):
        for file in self.files:
            # read data from file
            data = pd.read_excel(file)

            try:
                student_name = data.iloc[1,1]
                grader_1 = data.iloc[4,0]
                grader_2 = data.iloc[5,0]

                # Check if there are enough rows and columns in the data frame
                if data.shape[0] >= 5 and data.shape[1] >= 8:
                    score_1 = data.iloc[4,7]
                    score_2 = data.iloc[5,7]
                    average = (score_1 + score_2) / 2
                    new_row = pd.DataFrame({'Student Name': [student_name], 'Grader 1': [grader_1], 'Grader 2': [grader_2], 'Score 1': [score_1], 'Score 2': [score_2], 'Average': [average]})
                    self.student_evals = pd.concat([self.student_evals, new_row], ignore_index=True)

            except IndexError:
                continue

        # Write the data frame to Excel
        self.student_evals.to_excel('student_evals.xlsx', index=False)

        # return the full Excel file
        return pd.read_excel('student_evals.xlsx')





# Example usage for Solution class
solution = Solution()
solution.run()


# Example usage:
generator = RegistrationFormPDFGenerator(glob.glob('*.eml'))
name_dict = generator.extract_file_url()
generator.generate_pdfs()
generator.move_pdfs()


# Example usage for StudentEvaluations class
evals = StudentEvaluations()
evals.prompt_user_for_folder()
evals.write_scores()
