import cv2
import pytesseract
import pandas as pd
import os
from tqdm import tqdm
import configparser
import openpyxl

config = configparser.RawConfigParser()
config.read('keys.properties')
source = config.get('key', 'CSV_FILE_FROM_ANNOTATION')
target = config.get('key', 'AUTO_CORRECTED_CSV_FILE')
clean = config.get('key', 'clean')

class Annotation:
    def __init__(self, csv_file_path, images_folder):
        self.csv_file_path = csv_file_path
        self.images_folder = images_folder
        self.df = pd.read_csv(csv_file_path, delimiter='\t')  # Specify tab delimiter
        self.out_put_field = clean

    @staticmethod   # used for edit csv file and return New CSV file location
    def CSV_editor(source, target):
        print('Source File Name : ', source)
        df = pd.read_csv(source, delimiter='\t')  # Specify tab delimiter
        df['bbox_height'] = df['bbox_height'].apply(lambda x: x if x >= 25 else 26)
        df['bbox_width'] = df['bbox_width'].apply(lambda x: x if x >= 20 else 100)
        df.to_csv(target, index=False, sep='\t')  # Specify tab delimiter
        print("Auto correction Completed !!")
        return target

    # extract all values from a CSV and feed all to extract_text_from_image method
    # in addition, it is used to print values for a given key
    def annotate_images(self, source, print_extracted_text=False):
        print("Columns in the source DataFrame:", source.columns)  # Debugging line
        output_data = []
        for index, row in tqdm(source.iterrows(), total=len(source), desc="Annotating Images"):
            try:
                image_name = row['image_name']
                base_name = '_'.join(image_name.split('_')[:-1])
                image_path = os.path.join(self.images_folder, image_name)
                label_name = row['label_name']
                extracted_text = self.extract_text_from_image(image_path, row['bbox_x'], row['bbox_y'], row['bbox_width'], row['bbox_height'], row)
                output_data.append({
                    'base_name': base_name,
                    'label_name': label_name,
                    'extracted_text': extracted_text
                })
                if print_extracted_text:
                    print(extracted_text)
            except KeyError as e:
                print(f"Missing column: {e}")  # Debugging line

        # Grouping data by 'base_name'
        grouped_data = {}
        for item in output_data:
            base_name = item['base_name']
            if base_name not in grouped_data:
                grouped_data[base_name] = {}
            if item['label_name'] not in grouped_data[base_name]:
                grouped_data[base_name][item['label_name']] = []
            grouped_data[base_name][item['label_name']].append(item['extracted_text'])

        # Converting grouped data to a list of dictionaries
        final_output = []
        for base_name, data in grouped_data.items():
            row_data = {'image_name': base_name}
            for label_name, texts in data.items():
                # Filter out None values before joining texts
                cleaned_texts = [text for text in texts if text is not None]
                row_data[label_name] = '|'.join(cleaned_texts).strip()
            final_output.append(row_data)

        # Converting list of dictionaries to DataFrame
        ocr_output_df = pd.DataFrame(final_output)
        return ocr_output_df

    # get text from image with respect to the key and values and stored in a text format
    def extract_text_from_image(self, image_path, bbox_x, bbox_y, bbox_width, bbox_height, row):
        try:
            image = cv2.imread(image_path)
            cropped_image = image[bbox_y:bbox_y + bbox_height, bbox_x:bbox_x + bbox_width]
            text = pytesseract.image_to_string(cropped_image)
            return text
        except Exception as e:
            print(f"Error processing image: {image_path}")
            print(f"Error details: {e}")
            print(f"Row data: {row}")
            print("-" * 50)
            return None

    # used to store all extracted values into the Excel sheet
    def save_to_excel(self, excel_file_path, df):
        ocr_output_df = self.annotate_images(df)
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='w') as writer:
            ocr_output_df.to_excel(writer, index=False, sheet_name='Extracted_Data')
        print("\nExtraction complete. Data saved to Excel file.")

    # used to review values by key label name
    def key_value(self, df):
        print(config.get('key', 'message'))
        target_value = input("Enter the Key Name: ")
        if target_value == '1':
            print(df[config.get('key','column_name')].unique())
            target_value = input("Enter the Key Name: ")
            key_df = df[df[config.get('key', 'column_name')] == target_value]
            print(key_df)
            self.annotate_images(key_df, print_extracted_text=True)
        elif target_value in df[config.get('key', 'column_name')].values:
            key_df = df[df[config.get('key', 'column_name')] == target_value]
            print(key_df)
            self.annotate_images(key_df, print_extracted_text=True)


    # used for various cleaning purposes
    def remove_duplicates(self,cell_value):
        # Split the cell value by '|' and strip spaces from each item
        items = [item.strip() for item in cell_value.split('|')]
        # Use a dictionary to maintain order and remove duplicates (case insensitive)
        seen = {}
        for item in items:
            lowercase_item = item.lower()
            if lowercase_item not in seen:
                seen[lowercase_item] = item
        # Join the unique items with '|' symbol
        unique_items = list(seen.values())
        return ' | '.join(unique_items)





if __name__ == "__main__":
    print(config.get('key', 'sms'))
    user = input('ENTER THE INPUT: ')
    if user == '1':
        print('Annotating File Name :', source)
        annotation = Annotation(source, images_folder='Images')
        annotation.save_to_excel(excel_file_path='extracted_data.xlsx', df=annotation.df)
    elif user == '2':
        out_put = Annotation.CSV_editor(source, target)
        print('Annotating File Name :', out_put)
        annotation = Annotation(out_put, images_folder='Images')
        annotation.save_to_excel(excel_file_path='extracted_data.xlsx', df=annotation.df)

    elif user == '3':
        annotation = Annotation(source, images_folder='Images')
        annotation.key_value(annotation.df)