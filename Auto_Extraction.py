from pyresparser import ResumeParser
import warnings
import nltk
import pytesseract
from PIL import Image
import cv2
import os
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)

# Ensure stopwords are downloaded
nltk.download('stopwords')


def ocr_image_to_text(image_path):
    # Print the image path for debugging
    print(f"Reading image from path: {image_path}")

    # Load the image using OpenCV
    img = cv2.imread(image_path)

    # Check if the image is loaded successfully
    if img is None:
        raise FileNotFoundError(f"Image file not found at path: {image_path}")

    # Convert the image to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Use Tesseract to extract text
    text = pytesseract.image_to_string(gray)

    # Save the extracted text to a temporary file
    temp_txt_file = 'temp_resume.txt'
    with open(temp_txt_file, 'w', encoding='utf-8') as f:
        f.write(text)

    return temp_txt_file


# Define the labels
labels = ["image_name", "name", "country", "state", "tel_no", "email", "occupation",
          "edu_grad_year", "education", "job_responsibilities", "dept_worked", "photo",
          "grad_institution", "nationality", "med_devices_equip", "skills", "training",
          "nurse_regn", "post_code", "city", "exp_orgn", "experience_yrs", "addl_certif",
          "reference"]

# Initialize an empty DataFrame
df = pd.DataFrame(columns=labels)

# Path to the resume image (Ensure the file extension is included)
resume_image_path = "Images/Ivy Tran CV_page-0001.jpg"  # Correct the file path

# Convert the image to text
temp_txt_file = ocr_image_to_text(resume_image_path)

# Parse the text file using pyresparser
data = ResumeParser(temp_txt_file).get_extracted_data()

# Clean up the temporary text file
os.remove(temp_txt_file)

# Prepare the data row
row = {
    "image_name": resume_image_path,
    "name": data.get("name", ""),
    "country": "",  # You can add custom extraction logic here
    "state": "",  # You can add custom extraction logic here
    "tel_no": data.get("mobile_number", ""),
    "email": data.get("email", ""),
    "occupation": "",  # You can add custom extraction logic here
    "edu_grad_year": "",  # You can add custom extraction logic here
    "education": data.get("degree", ""),
    "job_responsibilities": "",  # You can add custom extraction logic here
    "dept_worked": "",  # You can add custom extraction logic here
    "photo": "",  # You can add custom extraction logic here
    "grad_institution": "",  # You can add custom extraction logic here
    "nationality": "",  # You can add custom extraction logic here
    "med_devices_equip": "",  # You can add custom extraction logic here
    "skills": ', '.join(data.get("skills", [])) if data.get("skills") else "",
    "training": "",  # You can add custom extraction logic here
    "nurse_regn": "",  # You can add custom extraction logic here
    "post_code": "",  # You can add custom extraction logic here
    "city": "",  # You can add custom extraction logic here
    "exp_orgn": ', '.join(data.get("company_names", [])) if data.get("company_names") else "",
    "experience_yrs": "",  # You can add custom extraction logic here
    "addl_certif": "",  # You can add custom extraction logic here
    "reference": ""  # You can add custom extraction logic here
}

# Convert the row to a DataFrame and concatenate it
row_df = pd.DataFrame([row])
df = pd.concat([df, row_df], ignore_index=True)

# Save the DataFrame to an Excel file
df.to_excel("resume_data.xlsx", index=False)

print("Data extracted and saved to resume_data.xlsx")
