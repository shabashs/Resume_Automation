from pyresparser import ResumeParser
import warnings
import nltk

warnings.filterwarnings("ignore", category=UserWarning)

# Ensure stopwords are downloaded
nltk.download('stopwords')

data = ResumeParser("Angel Jacob CV.pdf").get_extracted_data()
print("Name:", data["name"])
print("Email:", data["email"])
print("Mobile Number:", data["mobile_number"])
print("Skills:", data["skills"])
print("College Name:", data["college_name"])
print("Degree:", data["degree"])
print("Designation:", data["designation"])
print("Company Names:", data["company_names"])
