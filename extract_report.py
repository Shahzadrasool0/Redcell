from pathlib import Path
from docling.document_converter import DocumentConverter
import pandas as pd
import re

def extract_report(file_path):
    converter = DocumentConverter()
    result = converter.convert(file_path)
    text = result.document.export_to_markdown()

    data = {
        "Patient Name": "",
        "Age": "",
        "Gender": "",
        "Date": "",
        "Hemoglobin": "",
        "Hematocrit (PCV)": "",
        "RBC": "",
        "MCV": "",
        "MCH": "",
        "MCHC": "",
        "WBC": "",
        "Neutrophils": "",
        "Lymphocytes": "",
        "Eosinophils": "",
        "Monocytes": "",
        "Basophils": "",
        "Blast Cells": "",
        "Platelets": ""
    }

    # ----------------- Patient Info -----------------
    name_match = re.search(r"Patient Name\s*:?\s*\n([A-Z\s/]+)", text, re.IGNORECASE)
    if name_match:
        data["Patient Name"] = name_match.group(1).strip()

    age_gender_match = re.search(r"Age\s*\|\s*Gender\s*:\s*\n([\d\sA-Za-z]+)\|?\s*([A-Za-z]*)", text, re.IGNORECASE)
    if age_gender_match:
        data["Age"] = age_gender_match.group(1).strip()
        data["Gender"] = age_gender_match.group(2).strip()

    date_match = re.search(r"\n(\d{2}/\d{2}/\d{4})", text)
    if date_match:
        data["Date"] = date_match.group(1)

    # ----------------- Lab Results -----------------
    lab_section = text.split("**TEST NAME**")[-1]
    bold_values = re.findall(r"\*\*(.*?)\*\*", lab_section)

    # ⚠️ Drop the first 3 items: RESULT, UNITS, REFERENCE VALUE
    if len(bold_values) > 3:
        bold_values = bold_values[3:]

    ordered_tests = [
        "Hemoglobin",
        "Hematocrit (PCV)",
        "RBC",
        "MCV",
        "MCH",
        "MCHC",
        "WBC",
        "Neutrophils",
        "Lymphocytes",
        "Eosinophils",
        "Monocytes",
        "Basophils",
        "Blast Cells",
        "Platelets",
    ]

    for i, test in enumerate(ordered_tests):
        if i < len(bold_values):
            data[test] = bold_values[i].strip()

    return data


if __name__ == "__main__":
    file_path = Path("Abdullah.docx")
    extracted = extract_report(file_path)
    df = pd.DataFrame(list(extracted.items()), columns=["Field", "Value"])
    print(df)
