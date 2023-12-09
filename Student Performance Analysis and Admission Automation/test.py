import json
import subprocess
import os
import zipfile
import pandas as pd
from faker import Faker
import datetime
from docxtpl import DocxTemplate

# Constants
MINIMUM_MATH_SCORE = 80
MINIMUM_READING_SCORE = 85
MINIMUM_WRITING_SCORE = 85
VACANCIES = 100
DATASET_NAME = 'rkiattisak/student-performance-in-mathematics'
RESOURCES_FOLDER = "project_assets"  
KAGGLE_CREDENTIAL_PATH = os.path.join(RESOURCES_FOLDER, "kaggle.json")
ADMISSION_TEMPLATE_PATH = os.path.join(RESOURCES_FOLDER, "admission_template.docx")

def read_kaggle_credentials():
    try:
        print("... loading credentials\n")
        with open(KAGGLE_CREDENTIAL_PATH, "r") as json_file:
            return json.load(json_file)
    except Exception as e:
        print(f"An Error has ocurred while reading the kaggle credentials: \n {e}")

def set_kaggle_credentials(credentials):
    os.environ['KAGGLE_USERNAME'] = credentials["username"]
    os.environ['KAGGLE_KEY'] = credentials["key"]

def download_dataset(download_path):
    try:
        print("... downloading dataset\n")
        subprocess.run(f'kaggle datasets download -d "{DATASET_NAME}" -p "{download_path}"', shell=True)
    except Exception as e:
        print(f"An Error has ocurred while downloading the dataset: \n {e}")

def extract_dataset(download_path, extraction_path):
    try:
        print("... extracting dataset\n")
        zip_file_path = os.path.join(download_path, 'student-performance-in-mathematics.zip')
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extraction_path)
    except Exception as e:
        print(f"An Error has ocurred while extracting the dataset: \n {e}")

def load_student_data(extraction_path):
    csv_path = os.path.join(extraction_path, 'exams.csv')
    return pd.read_csv(csv_path)

def process_student_data(df):
    try:
        print("... processing Student data\n")
        df['total score'] = df['math score'] + df['reading score'] + df['writing score']
        df = df[df['math score'] >= MINIMUM_MATH_SCORE]
        df = df[df['reading score'] >= MINIMUM_READING_SCORE]
        df = df[df['writing score'] >= MINIMUM_WRITING_SCORE]
        return df.sort_values(by='total score', ascending=False).head(VACANCIES)
    except Exception as e:
        print(f"An Error has ocurred while extracting the dataset: \n {e}")

def add_fake_names(df):
    print("... adding fake names to preserve student's privacy\n")
    fake = Faker()
    df['name'] = [fake.name() for _ in range(len(df))]
    return df

def generate_admission_letters(admited_students, output_admited_folder):
    print("... generating letters to admited students\n")
    date_now = datetime.datetime.now().strftime("%m/%d/%Y")
    doc = DocxTemplate(ADMISSION_TEMPLATE_PATH)

    for _, row in admited_students.iterrows():
        context = {"Student_Name": row['name'], "date": date_now}
        doc.render(context)
        output_file = os.path.join(output_admited_folder, f"{row['name']}_Welcome_Letter.docx")
        doc.save(output_file)
    
    return output_admited_folder
    

def main():
    kaggle_credentials = read_kaggle_credentials()
    set_kaggle_credentials(kaggle_credentials)

    project_folder = os.path.dirname(__file__)
    download_path = os.path.join(project_folder, 'download')
    extraction_path = os.path.join(project_folder, 'extracted_files')
    output_admited_folder = os.path.join(project_folder, 'admited_students')

    for path in [download_path, extraction_path, output_admited_folder]:
        if not os.path.exists(path):
            os.makedirs(path)

    download_dataset(download_path)
    extract_dataset(download_path, extraction_path)
    
    student_data = load_student_data(extraction_path)
    processed_data = process_student_data(student_data)
    admited_students = add_fake_names(processed_data)

    generate_admission_letters(admited_students, output_admited_folder)
    
    print(f"Student letters successfully created at path: \n {output_admited_folder}")

if __name__ == "__main__":
    main()
