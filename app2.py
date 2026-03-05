import streamlit as st
import pandas as pd
import tempfile
from pathlib import Path
from docling.document_converter import DocumentConverter

st.set_page_config(page_title="Docling Report Extractor", layout="wide")
st.title("🧪 Extract Lab Report Data")

# Upload DOCX file
uploaded_file = st.file_uploader("Upload a Word Report (.docx)", type=["docx"])

def extract_structured_data(doc_path):
    """Extract structured info from Word report using Docling"""
    converter = DocumentConverter()
    result = converter.convert(doc_path)

    text = result.document.export_to_text()

    # Define key fields to capture
    fields = {
        "Patient Name": None,
        "Age": None,
        "Gender": None,
        "Date": None,
        "Hemoglobin": None,
        "Hematocrit (PCV)": None,
        "RBC": None,
        "WBC": None,
        "Platelets": None
    }

    # Simple parsing (line by line search)
    for line in text.splitlines():
        line_clean = line.strip()
        for key in fields.keys():
            if key.lower() in line_clean.lower():
                parts = line_clean.split(":")
                if len(parts) > 1:
                    fields[key] = parts[1].strip()

    return fields


if uploaded_file is not None:
    if st.button("Extract Data"):
        try:
            # Save uploaded file to temp
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                tmp.write(uploaded_file.getbuffer())
                tmp_path = Path(tmp.name)

            # Extract structured data
            data = extract_structured_data(tmp_path)

            # Show in table
            df = pd.DataFrame(list(data.items()), columns=["Field", "Value"])
            st.success("✅ Data Extracted Successfully!")
            st.dataframe(df, use_container_width=True)

        except Exception as e:
            st.error(f"Error extracting data: {e}")
