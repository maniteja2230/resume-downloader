"""
Generate a sample Excel file for testing the Bulk Resume Downloader.
Run once:  python create_sample_excel.py
"""
import pandas as pd

data = {
    "name": [
        "Ravi Kumar",
        "Priya Sharma",
        "Arjun Mehta",
        "Sneha Patel",
        "Kiran Reddy",
    ],
    "resume_link": [
        "https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE1uAY/view?usp=sharing",
        "https://drive.google.com/open?id=1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE1uAY",
        "https://drive.google.com/uc?id=1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE1uAY",
        "https://drive.google.com/file/d/PRIVATE_FILE_ID/view",   # will be skipped
        "not-a-valid-link",                                         # invalid
    ],
}

df = pd.DataFrame(data)
df.to_excel("sample_students.xlsx", index=False)
print("✅ sample_students.xlsx created successfully!")
print(df.to_string(index=False))
