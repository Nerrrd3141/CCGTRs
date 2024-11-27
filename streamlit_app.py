import streamlit as st
import openpyxl
import language_tool_python
import re
from datetime import datetime, timedelta
from io import BytesIO
import os


def process_excel(file: BytesIO, original_file_name: str):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.active

    tool = language_tool_python.LanguageTool('en-US')

    start_row, end_row = 13, 19
    col = 'L'

    for row in range(start_row, end_row + 1):
        cell = sheet[f"{col}{row}"]
        original_text = cell.value

        if original_text:
            matches = tool.check(original_text)
            corrected_text = language_tool_python.utils.correct(original_text, matches)
            corrected_text = re.sub(r'\b(CCG|CL|CM)\b', lambda m: m.group(0), corrected_text)
            cell.value = corrected_text
            cell.font = openpyxl.styles.Font(name="Calibri", size=14)

    for row in range(start_row, end_row + 1):
        sheet.row_dimensions[row].height = None

    today = datetime.now()
    last_week_start = (today - timedelta(days=today.weekday() + 7)).strftime("%m.%d.%y")
    last_week_end = (today - timedelta(days=today.weekday() + 1)).strftime("%m.%d.%y")

    base_name, extension = os.path.splitext(original_file_name)
    if not extension:
        extension = ".xlsx"
    new_file_name = f"{base_name} CCG Activity Tracking Form {last_week_start} - {last_week_end}{extension}"

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output, new_file_name


# Streamlit App
st.set_page_config(page_title="CCG Activity Tracking Processor", layout="wide", page_icon="📊")

# Sidebar
st.sidebar.header("📂 File Upload")
uploaded_files = st.sidebar.file_uploader(
    "Upload one or more Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# Main Page
st.title("📊 CCG Activity Tracking Form Processor")
st.markdown("""
This tool processes **Excel files** to:
- Correct spelling and grammar in specific cells.
- Rename the files dynamically with the current date range.
- Maintain their original structure and formatting.

**Instructions**:
1. Upload one or more `.xlsx` files using the sidebar.
2. Processed files will appear below for download.
""")

if uploaded_files:
    st.subheader(f"Processing {len(uploaded_files)} File(s):")

    processed_files = {}
    for uploaded_file in uploaded_files:
        original_file_name = uploaded_file.name
        st.write(f"🟢 **Processing**: `{original_file_name}`...")

        processed_file, new_file_name = process_excel(uploaded_file, original_file_name)
        processed_files[new_file_name] = processed_file

    st.success(f"✅ Successfully processed {len(processed_files)} file(s)!")
    st.balloons()  # Add celebratory confetti animation
    st.image("https://media.tenor.com/iRL2vkMKuJMAAAAi/yay-happy.gif", width=300)  # Add celebratory GIF with smaller size
    st.markdown("---")

    # Display processed files in a neat format
    for new_file_name, processed_file in processed_files.items():
        col1, col2 = st.columns([3, 1])
        with col1:
            st.write(f"📁 **File**: `{new_file_name}`")
        with col2:
            st.download_button(
                label="⬇️ Download",
                data=processed_file,
                file_name=new_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("📂 Upload files in the sidebar to begin processing.")
