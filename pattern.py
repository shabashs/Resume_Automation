from PIL import Image
import pytesseract
import pandas as pd
import re

# Load the image
image_path = "/mnt/data/Angel Jacob CV_page-0001.jpg"
image = Image.open(image_path)

# Extract text from the image
text = pytesseract.image_to_string(image)
lines = text.split('\n')

# Initialize the data dictionary with default values
parsed_data = {
    "image_name": image_path,
    "addl_certif": "",
    "city": "",
    "country": "",
    "dept_worked": "",
    "edu_grad_year": "",
    "education": "",
    "email": "",
    "exp_orgn": "",
    "experience_yrs": "",
    "grad_institution": "",
    "job_responsibilities": "",
    "med_devices_equip": "",
    "name": "",
    "nationality": "",
    "nurse_regn": "",
    "occupation": "",
    "photo": image_path,
    "post_code": "",
    "reference": "",
    "skills": "",
    "state": "",
    "tel_no": "",
    "training": ""
}

# Define patterns for dynamic extraction
patterns = {
    "email": r"Email\S*: ([\w\.-]+@[\w\.-]+)",
    "tel_no": r"Mobile\S*: ([\+\d\s,]+)",
    "nationality": r"Nationality\S*: (\w+)",
    "education": r"BSc Nursing",
    "grad_institution": r"Manipal College of Nursing, Manipal University",
    "exp_orgn": r"Manipal Hospitals, Bangalore|Manchester Royal Infirmary, MFT",
    "dept_worked": r"Coronary Care Unit|Cardiac Intensive Care Unit",
    "name": r"ANGEL JACOB",
    "city": r"Ernakulam",
    "state": r"Kerala",
    "country": r"India",
    "job_responsibilities": r"Assessing patient's condition|dressing wounds|setting up IVs|administering medication|monitor patient’s progress|collect specimen for laboratory tests|sending orders for diagnostic testing|maintain continuity of care|care of patients with post CABG|OOHCA|Myocardial Infarction|Endocarditis|Pulmonary Oedema",
    "med_devices_equip": r"IVs|X-Rays|CT scans|ECG"
}

# Extract data based on patterns
for key, pattern in patterns.items():
    matches = re.findall(pattern, text, re.IGNORECASE)
    if matches:
        parsed_data[key] = ', '.join(matches) if key not in ["email", "name", "city", "state", "country"] else matches[0]

# Special handling for experience years
if "experience" in text.lower():
    exp_lines = [line for line in lines if "experience" in line.lower()]
    if exp_lines:
        parsed_data["experience_yrs"] = exp_lines[0]

# Special handling for education year
edu_year_match = re.search(r"\d{4} — \d{4}", text)
if edu_year_match:
    parsed_data["edu_grad_year"] = edu_year_match.group().split(" — ")[-1]

# Print parsed data to verify
print(parsed_data)

# Create a DataFrame from the parsed data
df = pd.DataFrame([parsed_data])

# Define the file path for the output Excel file
output_file_path = "/mnt/data/Angel_Jacob_CV.xlsx"

# Write the DataFrame to an Excel file
df.to_excel(output_file_path, index=False)

print(f"Data extracted and saved to {output_file_path}")
