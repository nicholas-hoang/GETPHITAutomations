# Install Requirements from requirements.txt
import os
import pandas as pd
import numpy as np
import shutil
import pathlib
import xlwings as xw
import glob
from openpyxl import load_workbook
pip install - r requirements.txt

# Load Modules from requirements.txt


class BootCamp:

    def __init__(self, file, evaluation):
        self.file = file
        self.evaluation = evaluation

    def load_data(self):
        student_registration = pd.read_excel(self.file)
        student_registration = student_registration.rename(columns={
            'Registration Date': 'Registration_Date',
            'First Name': 'First_Name',
            'Last Name': 'Last_Name'})

        student_registration.Registration_Date = student_registration.Registration_Date.dt.strftime(
            '%m/%d/%Y')

        with open(self.evaluation, 'r') as f:
            evaluation_form = pd.read_excel(f)

        return student_registration, evaluation_form


def get_first_last_name(row):
    fullname = row.First_Name + " " + row.Last_Name
    return fullname


def get_existing_folders():
    """ get_existing_folders() function checks the current directory for existing folders and returns a list of the folders found.
    """
    num_of_docs = [directory for directory in os.listdir(
        '.') if os.path.isdir(os.path.join(".", directory))]
    print(
        f'There are a total of {len(num_of_docs)} student folders in this directory.')
    return num_of_docs


def create_new_folders():
    """_summary_ create_new_folders() function creates a new folder for each student in the dataframe.
    """
    count = 0
    student_names = pd.Series(df.full_name)

    for student in student_names:
        os.mkdir(student.title())  # Create Student Folder
        count += 1

    print(f'A total of {count} student folders were generated.')


def fill_folders(evaluation_form):
    """ fill_folders(evaluation_form): takes the evaluation form and places it in each student folder.
    """
    path = pathlib.Path('.')
# Put Evaluation Form In Student Folders
    for item in path.iterdir():
        if item.is_dir():
            shutil.copy2(evaluation_form, item)  # Place Form in Folder


def rename_files():
    """ rename_files() function renames the files in the student folders to the name of the folder.
    """
    path = pathlib.Path('.')
    for folder in path.iterdir():
        if folder.is_dir():
            for file in folder.iterdir():
                if file.is_file():
                    new_file = folder.name + file.suffix
                    file.rename(path / folder.name / new_file)


def writeStudentNames():
    """ writeStudentNames() function writes the student names to the evaluation form.
    """
    # Get a list of files with extension .xlsx
    student_files = glob.glob('*.xlsx')

    # Iterate through each file in student_files
    for file in student_files:
        try:
            # Load Workbook in Binary Mode
            workbook = load_workbook(filename=file)

            # Define desired sheet names
            sheet = workbook['Evaluator 1']

            # Extract student name from filename
            student_name = os.path.splitext(file.name)[0]

            # Define desired cell value
            sheet['B3'] = student_name

            # Save and close the workbook
            workbook.save(filename=file)
            workbook.close()

        except Exception as e:
            print(f"Error writing student name to {file}: {e}")

# When Running this Script from the command line, the following code will execute.
# Replace the file names with the names of the files you are using.
# The first file is the student registration file i.e Bootcamp Applications UTEP 5-1-23.xlsx
# The second file is the evaluation form i.e Evaluation_Rubric.xlsx


def main():
    # TODO Change File Names Here.
    BootCamp = BootCamp(
        'Bootcamp Applications UTEP 5-1-23.xlsx', 'Evaluation_Rubric.xlsx')
    df = BootCamp.load_data()[0]
    get_existing_folders()
    df['full_name'] = df.apply(get_first_last_name, axis=1)
    create_new_folders()
    fill_folders(BootCamp.load_data()[1])
    rename_files()
    writeStudentNames()


if __name__ == "__main__":
    main()

# TODO: Create a function that will write the registration date to the evaluation form.
