import pandas as pd
import re


class Annotation:

    def __init__(self, file_path):
        self.file_path = file_path

    def clean_data(self):
        df = pd.read_excel(self.file_path)

        # Replace NaN values with empty strings
        df = df.fillna('')

        # Cleaning logic within the single method
        if 'name' in df.columns:
            df['name'] = df['name'].apply(lambda x: re.sub(r'[^a-zA-Z\s]', '', str(x)))

        if 'tel_no' in df.columns:
            df['tel_no'] = df['tel_no'].apply(lambda x: re.sub(r'[^0-9+]', '', str(x)))

        if 'email' in df.columns:
            df['email'] = df['email'].apply(lambda x: re.sub(r'[^a-zA-Z0-9@._-]', '', str(x)))

        if 'occupation' in df.columns:
            df['occupation'] = df['occupation'].apply(lambda x: re.sub(r'[^a-zA-Z\s]', '', str(x)))

        if 'post_code' in df.columns:
            df['post_code'] = df['post_code'].apply(lambda x: re.sub(r'[^0-9]', '', str(x)))

        if 'edu_grad_year' in df.columns:
            df['edu_grad_year'] = df['edu_grad_year'].apply(lambda x: re.sub(r'[^0-9]', '', str(x)))

        if 'education' in df.columns:
            df['education'] = df['education'].apply(lambda x: re.sub(r'[^a-zA-Z\s]', '', str(x)))

        if 'job_responsibilities' in df.columns:
            df['job_responsibilities'] = df['job_responsibilities'].apply(lambda x: re.sub(r'[^a-zA-Z\s]', '', str(x)))

        if 'dept_worked' in df.columns:
            df['dept_worked'] = df['dept_worked'].apply(lambda x: re.sub(r'[^a-zA-Z\s]', '', str(x)))


        df.to_excel(self.file_path, index=False)

        print(f"Cleaned data saved to {self.file_path}")


file_path = 'Merged_BOOK1.xlsx'
annotation = Annotation(file_path)
annotation.clean_data()
