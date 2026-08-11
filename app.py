import streamlit as st
import pandas as pd
import re
import os
from PyPDF2 import PdfReader
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment

# Configure the web page
st.set_page_config(page_title="SRM Schedule Consolidator", page_icon="🤖", layout="centered")
st.title("SRM Schedule PDF Consolidator 🤖")
st.write("Drag and drop your PDF schedules below to instantly generate your formatted Excel sheet.")

# File uploader allows multiple files 
uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("Process Files"):
        all_rows = []
        
        with st.spinner("Extracting and formatting data..."):
            for file in uploaded_files:
                file_name = file.name
                try:
                    # 1. Read the PDF from memory
                    reader = PdfReader(file)
                    text = ""
                    for page in reader.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + " "
                    
                    clean_text = re.sub(r'\s+', ' ', text)
                    
                    # 2. Extract Programme & Subject from filename
                    name_without_ext = os.path.splitext(file_name)[0]
                    if '_' in name_without_ext:
                        parts = name_without_ext.split('_', 1) 
                        programme = parts[0].strip()
                        subject = parts[1].strip()
                    else:
                        programme = name_without_ext
                        subject = "See Filename"
                    
                    # 3. Regex Extraction
                    dates = re.findall(r'\d{2}/\d{2}/\d{4}', clean_text)
                    link_match = re.search(r'https?://[^\s]+', clean_text)
                    zoom_link = link_match.group(0) if link_match else ""
                    day_match = re.search(r'\b(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b', clean_text)
                    day = day_match.group(0) if day_match else ""
                    time_match = re.search(r'\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M', clean_text, re.IGNORECASE)
                    time_str = time_match.group(0) if time_match else ""
                    
                    # 4. Build the Rows
                    if dates:
                        for i, date_val in enumerate(dates):
                            all_rows.append({
                                'Programme & Semester': programme,
                                'Subject Name': subject,
                                'Class': f"Class {i+1}",
                                'Day': day,
                                'Date': date_val,
                                'Time (PM)': time_str,
                                'Zoom Link': zoom_link
                            })
                except Exception as e:
                    st.error(f"Error processing {file_name}: {e}")
        
        # 5. Format and Output
        if all_rows:
            df = pd.DataFrame(all_rows)
            st.success(f"✅ Successfully processed {len(uploaded_files)} files!")
            
            # Create Excel file in memory
            output = BytesIO()
            df.to_excel(output, index=False, sheet_name="Schedule")
            output.seek(0)
            
            # Apply OpenPyXL formatting (merging and centering)
            wb = openpyxl.load_workbook(output)
            ws = wb.active
            
            merge_cols = [1, 2, 4, 6, 7]
            for col_idx in merge_cols:
                start_row = 2
                for row_idx in range(3, ws.max_row + 2):
                    val_current = ws.cell(row=row_idx, column=col_idx).value if row_idx <= ws.max_row else None
                    val_start = ws.cell(row=start_row, column=col_idx).value
                    
                    if val_current != val_start:
                        if row_idx - 1 > start_row:
                            ws.merge_cells(start_row=start_row, start_column=col_idx, end_row=row_idx - 1, end_column=col_idx)
                            ws.cell(row=start_row, column=col_idx).alignment = Alignment(vertical='center', horizontal='center')
                        start_row = row_idx
            
            final_output = BytesIO()
            wb.save(final_output)
            final_output.seek(0)
            
            st.download_button(
                label="⬇️ Download Formatted Excel Sheet",
                data=final_output,
                file_name='Consolidated_Live_Session_Time_Table.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
        else:
            st.warning("No valid data could be extracted.")
