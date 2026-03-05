import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
import streamlit as st
import torch
import pandas as pd
import torch.nn as nn
import base64
import io
import matplotlib.pyplot as plt
import docx
import re

st.set_page_config(page_title="RedCell AI", layout="wide")

# Function to get base64 image string
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

img_base64 = get_base64_image("images/logo.png")  # Adjust path if needed
logo1 = get_base64_image("images/logo1.png")
logo2 = get_base64_image("images/logo2.png")
logo3 = get_base64_image("images/logo3.png")

collab1_base64 = get_base64_image("images/collab1.png")      # Footer logo 1
collab2_base64 = get_base64_image("images/collab2.png")      # Footer logo 2
collab3_base64 = get_base64_image("images/collab3.png")      # Footer logo 3



# Custom HTML and CSS for layout
st.markdown(
    f"""
    <style>
    .header-container {{
        display: flex;
        align-items: center;
        justify-content: space-between; /* push extra logos to the right */
        padding: 10px 0 30px 0;
    }}
    .logo {{
        width: 130px;
        margin-right: 20px;
    }}
    .titles {{
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .titles h1 {{
        font-size: 2.2rem;
        margin: 0;
    }}
    .titles h3 {{
        font-size: 1.2rem;
        font-weight: normal;
        color: #555;
        margin: 0;
    }}
    .right-logos {{
        display: flex;
        gap: 15px;
    }}
    .right-logos img {{
        width: 60px;
        
    }}
    
    /* Footer style (center aligned) */
    .footer-container {{
        display: flex;
        justify-content: center;
        align-items: center;
        padding: 20px 0 10px 0;
        gap: 50px;
        border-top: 1px solid #ddd;
        margin-top: 50px;
    }}
    .footer-container img {{
        width: 120px;
    }}
    
    /* Responsive styling for mobile */
    @media (max-width: 768px) {{
        .header-container {{
            flex-direction: column;
            text-align: center;
        }}
        .header-container > div {{
            flex-direction: column;
            margin-bottom: 20px;
        }}
        .logo {{
            margin-right: 0;
            margin-bottom: 10px;
        }}
        .right-logos {{
            justify-content: center;
        }}
        .footer-container {{
            flex-direction: column;
            gap: 20px;
        }}
    }}
    </style>

    <div class="header-container">
        <div style="display:flex;align-items:center;">
            <img src="data:image/png;base64,{img_base64}" class="logo">
            <div class="titles">
                <h1> Precision in Every Drop</h1>
                <h3>AI Based Clinical Diagnostic System for Decision Support</h3>
            </div>
        </div>
        <div class="right-logos">
            <img src="data:image/png;base64,{logo1}">
            <img src="data:image/png;base64,{logo2}">            
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


st.markdown(
    "<h1 style='font-size:24px;'>Upload Reports</h1>",
    unsafe_allow_html=True
)


# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"], key="excel_upload")

st.set_page_config(page_title="Excel to CSV Converter", layout="wide")



# Font size for table
font_size = 16
st.markdown(
    f"""
    <style>
    .dataframe-wrapper tbody td {{
        font-size: {font_size}px !important;
    }}
    .dataframe-wrapper thead th {{
        font-size: {font_size + 2}px !important;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# Process file on button click
if uploaded_file is not None:
    if st.button("Convert to CSV"):
        try:
            # Read Excel file
            df = pd.read_excel(uploaded_file)
            st.success("File converted successfully!")
            st.dataframe(df, use_container_width=True)
            df = pd.read_excel(uploaded_file)
            df = df.astype(str)  # Prevent ArrowTypeError

            # Convert to CSV in memory
            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False)
            csv_bytes = csv_buffer.getvalue().encode()
            b64 = base64.b64encode(csv_bytes).decode()
            file_name = uploaded_file.name.rsplit('.', 1)[0] + '.csv'
            download_link = f'<a href="data:file/csv;base64,{b64}" download="{file_name}">📥 Download Converted CSV</a>'
            st.markdown(download_link, unsafe_allow_html=True)

            # Plot Histogram (for numeric columns)
            numeric_cols = df.select_dtypes(include='number').columns.tolist()
            if numeric_cols:
                st.subheader("📊 Histogram")
                selected_hist_col = st.selectbox("Select a column for histogram", numeric_cols)
                fig, ax = plt.subplots()
                df[selected_hist_col].plot(kind='hist', bins=20, ax=ax, color='skyblue', edgecolor='black')
                ax.set_title(f"Histogram of {selected_hist_col}")
                st.pyplot(fig)
            else:
                st.info("No numeric columns available for histogram.")

            # Plot Pie Chart (for categorical columns)
            cat_cols = df.select_dtypes(include='object').columns.tolist()
            if cat_cols:
                st.subheader("🥧 Pie Chart")
                selected_pie_col = st.selectbox("Select a column for pie chart", cat_cols)
                pie_data = df[selected_pie_col].value_counts()
                fig2, ax2 = plt.subplots()
                ax2.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%', startangle=90)
                ax2.set_title(f"Pie Chart of {selected_pie_col}")
                ax2.axis('equal')
                st.pyplot(fig2)
            else:
                st.info("No categorical columns available for pie chart.")

        except Exception as e:
            st.error(f"Error during processing: {e}")


st.markdown(
    f"""
    <div class="footer-container">
        <img src="data:image/png;base64,{collab1_base64}">
        <img src="data:image/png;base64,{collab2_base64}">
        <img src="data:image/png;base64,{collab3_base64}">
    </div>
    """,
    unsafe_allow_html=True
)


st.markdown("<h1 style='font-size:24px;'>Extract Data from Word Reports</h1>", unsafe_allow_html=True)

uploaded_word = st.file_uploader("Upload a Word Report (.docx)", type=["doc","docx"], key="word_upload")

def extract_report_data(docx_file):
    """Extract key values from lab report"""
    doc = docx.Document(docx_file)
    text = "\n".join([para.text for para in doc.paragraphs])
    data = {}

    # Patient Info
    name_match = re.search(r"Patient Name\s*:\s*(.+)", text)
    age_match = re.search(r"Age\s*\|\s*Gender\s*:\s*([\d]+.*)", text)
    date_match = re.search(r"Date\s*:\s*(.+)", text)

    if name_match:
        data["Patient Name"] = name_match.group(1).strip()
    if age_match:
        parts = age_match.group(1).split("|")
        if len(parts) >= 2:
            data["Age"] = parts[0].strip()
            data["Gender"] = parts[1].strip()
    if date_match:
        data["Date"] = date_match.group(1).strip()

    # Lab Test Values
    tests = {
        "Hemoglobin": r"Hb: Haemoglobin.*?\n?([0-9.]+)",
        "Hematocrit (PCV)": r"P\.C\.V.*?\n?([0-9.]+)",
        "RBC": r"Red Blood Cells.*?\n?([0-9.]+)",
        "MCV": r"M\.C\.V.*?\n?([0-9.]+)",
        "MCH": r"M\.C\.H\s*\n?([0-9.]+)",
        "MCHC": r"M\.C\.H\.C\s*\n?([0-9.]+)",
        "WBC": r"Total WBC Count\s*\n?([0-9.]+)",
        "Neutrophils": r"Neutrophils\s*\n?([0-9.]+)",
        "Lymphocytes": r"Lymphocytes\s*\n?([0-9.]+)",
        "Eosinophils": r"Eosinophils\s*\n?([0-9.]+)",
        "Monocytes": r"Monocytes\s*\n?([0-9.]+)",
        "Platelets": r"Platelets\s*\n?([0-9.]+)"
    }
    for test, pattern in tests.items():
        match = re.search(pattern, text)
        if match:
            data[test] = match.group(1).strip()

    return data


if uploaded_word is not None:
    if st.button("Extract Data from Word Report"):
        try:
            extracted_data = extract_report_data(uploaded_word)
            st.success("✅ Data Extracted Successfully!")
            st.json(extracted_data)  # shows all extracted info
        except Exception as e:
            st.error(f"Error extracting data: {e}")
st.markdown("<h1 style='font-size:24px;'>Extract Data from Word Reports</h1>", unsafe_allow_html=True)

uploaded_word = st.file_uploader("Upload a Word Report (.docx)", type=["doc","docx"])

def extract_report_data(docx_file):
    ...
    return data

if uploaded_word is not None:
    if st.button("Extract Data from Word Report"):
        try:
            extracted_data = extract_report_data(uploaded_word)
            st.success("✅ Data Extracted Successfully!")
            st.json(extracted_data)
        except Exception as e:
            st.error(f"Error extracting data: {e}")
