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
            display: none; /* Hide logo 1 and logo 2 on phones */
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


st.markdown("<br>", unsafe_allow_html=True) # Add some spacing

# --- Create Tabs for Cleaner UI ---
tab1, tab2 = st.tabs(["Excel to CSV Converter", "Word Report Extractor"])

with tab1:
    st.markdown("<h2 style='font-size:20px;'>Upload Excel Reports</h2>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"], key="excel_upload")

    # Font size for table
    font_size = 14
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
        button_col1, button_col2 = st.columns([1, 1])
        with button_col1:
            convert_clicked = st.button("Convert to CSV", type="primary", use_container_width=True)
        with button_col2:
            drop_nulls_clicked = st.button("🗑️ Remove Null Values & Convert", use_container_width=True)

        if convert_clicked or drop_nulls_clicked:
            try:
                # Read Excel file
                df = pd.read_excel(uploaded_file)
                
                # Apply drop null logic if that specific button was clicked
                if drop_nulls_clicked:
                    original_rows = len(df)
                    # We drop rows where EVERY column is NaN (completely empty rows)
                    # Alternatively, use dropna() to drop any row with any missing value
                    df.dropna(how='all', inplace=True) 
                    rows_dropped = original_rows - len(df)
                    st.success(f"✅ Removed {rows_dropped} empty rows and converted successfully!")
                else:
                    st.success("✅ File converted successfully!")
                
                # Use an expander to optionally hide the large dataframe
                with st.expander("Preview Data", expanded=True):
                    # Convert only overly complex object types to string to prevent Arrow serialization errors,
                    # but LEAVE numeric data as float/int so histograms work.
                    display_df = df.copy()
                    for col in display_df.select_dtypes(include=['object', 'datetime']).columns:
                        display_df[col] = display_df[col].astype(str)
                    
                    st.dataframe(display_df, use_container_width=True)

                # Convert to CSV in memory (using original df so data types remain intact)
                csv_buffer = io.StringIO()
                df.to_csv(csv_buffer, index=False)
                csv_bytes = csv_buffer.getvalue().encode()
                b64 = base64.b64encode(csv_bytes).decode()
                file_name = uploaded_file.name.rsplit('.', 1)[0] + '.csv'
                
                # Styled download button
                st.markdown(f'<a href="data:file/csv;base64,{b64}" download="{file_name}" class="download-btn" style="display: inline-block; padding: 10px 20px; background-color: #ff4b4b; color: white; border-radius: 5px; text-decoration: none; font-weight: bold; margin-bottom: 20px;">📥 Download Converted CSV</a>', unsafe_allow_html=True)
                
                st.divider()
                st.markdown("### 📈 Data Visualizations")

                # Put charts side-by-side using columns
                col1, col2 = st.columns(2)
                
                # Plot Histogram (for numeric columns)
                with col1:
                    numeric_cols = df.select_dtypes(include='number').columns.tolist()
                    if numeric_cols:
                        st.markdown("**Histogram**")
                        selected_hist_col = st.selectbox("Select a column", numeric_cols, key="hist_select")
                        fig, ax = plt.subplots(figsize=(5,3))
                        df[selected_hist_col].plot(kind='hist', bins=20, ax=ax, color='#ff4b4b', edgecolor='black', alpha=0.7)
                        ax.set_title(f"Histogram of {selected_hist_col}")
                        st.pyplot(fig)
                    else:
                        st.info("No numeric columns available.")

                # Plot Pie Chart (for categorical columns)
                with col2:
                    cat_cols = df.select_dtypes(include='object').columns.tolist()
                    if cat_cols:
                        st.markdown("**Pie Chart**")
                        selected_pie_col = st.selectbox("Select a column", cat_cols, key="pie_select")
                        pie_data = df[selected_pie_col].value_counts().head(10) # Limit to top 10 for cleaner chart
                        fig2, ax2 = plt.subplots(figsize=(5,3))
                        ax2.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%', startangle=90, colors=plt.cm.Pastel1.colors)
                        ax2.set_title(f"Top 10: {selected_pie_col}")
                        ax2.axis('equal')
                        st.pyplot(fig2)
                    else:
                        st.info("No categorical columns available.")

            except Exception as e:
                st.error(f"Error during processing: {e}")

with tab2:
    st.markdown("<h2 style='font-size:20px;'>Extract Data from Word Reports</h2>", unsafe_allow_html=True)
    uploaded_word = st.file_uploader("Upload a Word Report (.docx)", type=["doc","docx"], key="word_upload_2")

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
        if st.button("Extract Data from Word Report", type="primary"):
            try:
                with st.spinner("Analyzing document..."):
                    extracted_data = extract_report_data(uploaded_word)
                st.success("✅ Extracted Successfully!")
                
                # Display extracted data nicely in columns instead of raw JSON
                if extracted_data:
                    st.markdown("### Patient Information")
                    info_col1, info_col2, info_col3 = st.columns(3)
                    info_col1.metric("Patient Name", extracted_data.get("Patient Name", "N/A"))
                    info_col2.metric("Age/Gender", f"{extracted_data.get('Age', 'N/A')} / {extracted_data.get('Gender', 'N/A')}")
                    info_col3.metric("Date", extracted_data.get("Date", "N/A"))
                    
                    st.markdown("### Lab Results")
                    # Create a dataframe for a much cleaner display than JSON
                    results = {k: v for k, v in extracted_data.items() if k not in ["Patient Name", "Age", "Gender", "Date"]}
                    if results:
                        df_results = pd.DataFrame(list(results.items()), columns=["Test", "Value"])
                        st.dataframe(df_results, use_container_width=True, hide_index=True)
                else:
                    st.warning("Could not find structured data using the current logic.")
            except Exception as e:
                st.error(f"Error extracting data: {e}")

# --- Footer ---
