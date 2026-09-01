import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment
import pdf2image
import pytesseract

# Configure the web page
st.set_page_config(page_title="SRM Schedule Consolidator", page_icon="🤖", layout="centered")
st.title("SRM Schedule PDF Consolidator 🤖 (v7.0 - AI Vision)") 
st.write("Drag and drop your PDF schedules below to instantly generate your formatted Excel sheet.")

uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("Process Files"):
        all_rows = []
        
        with st.spinner("Taking snapshots and running AI Vision (OCR)... this may take a moment!"):
            for file in uploaded_files:
                file_name = file.name
                try:
                    file_bytes = file.getvalue()
                    
                    # Convert the PDF pages directly into high-res images!
                    images = pdf2image.convert_from_bytes(file_bytes, dpi=300)
                    
                    for page_num, img in enumerate(images):
                        try:
                            # Use AI to read the text visually from the image
                            page_text = pytesseract.image_to_string(img)
                            
                            if not page_text or page_text.strip() == "":
                                continue
                                
                            clean_text = re.sub(r'\s+', ' ', page_text)
                            no_space_text = re.sub(r'\s+', '', page_text)
                            
                            # 1. Extract Programme Name
                            name_without_ext = os.path.splitext(file_name)[0]
                            if '_' in name_without_ext:
                                programme = name_without_ext.split('_', 1)[0].strip()
                            else:
                                programme = name_without_ext.strip()
                                
                            # 2. Smart Subject Extraction
                            subject_match = re.search(r'Class\s*6\s+(.*?)\s+Class\s*7', clean_text, re.IGNORECASE)
                            if subject_match:
                                subject = subject_match.group(1).replace('|', '').strip()
                            elif '_' in name_without_ext:
                                subject = name_without_ext.split('_', 1)[1].strip()
                            else:
                                subject = "Unknown Subject"
                                
                            # 3. Nuclear Date Extraction
                            # Finds dates even if OCR adds weird artifacts between numbers
                            raw_dates = re.findall(r'(\d{1,2})[^a-zA-Z0-9]*(\d{1,2})[^a-zA-Z0-9]*(202\d)', no_space_text)
                            dates = []
                            for d, m, y in raw_dates:
                                formatted_date = f"{int(d):02d}/{int(m):02d}/{y}"
                                if formatted_date not in dates:
                                    dates.append(formatted_date)
                            
                            link_match = re.search(r'https?://[^\s]+', clean_text)
                            zoom_link = link_match.group(0) if link_match else ""
                            
                            day_match = re.search(r'\b(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b', clean_text)
                            day = day_match.group(0) if day_match else ""
                            
                            # 4. Fix Scrambled Times
                            time_match = re.search(r'\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M', clean_text, re.IGNORECASE)
                            if time_match:
                                time_str = time_match.group(0)
                            else:
                                times = re.findall(r'\d{1,2}:\d{2}', clean_text)
                                if len(times) >= 2:
                                    t_ints = [int(t.split(':')[0])*60 + int(t.split(':')[1]) for t in times[:2]]
                                    sorted_times = [x for _, x in sorted(zip(t_ints, times[:2]))]
                                    time_str = f"{sorted_times[0]} - {sorted_times[1]} PM"
                                else:
                                    time_str = ""
                                    
                            # 5. Build the Rows
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
                            else:
                                st.warning(f"⚠️ Skipped Page {page_num + 1}. No dates found via AI Vision.")
                                with st.expander(f"🔍 Click to view OCR TEXT for Page {page_num + 1}"):
                                    st.code(page_text)
                                
                        except Exception as e:
                            st.warning(f"⚠️ Error on Page {page_num + 1}: {e}")
                            
                except Exception as e:
                    st.error(f"❌ Critical error opening {file_name}: {e}")
        
        # 6. Format and Output
        if all_rows:
            df = pd.DataFrame(all_rows)
            st.success(f"✅ Successfully processed {len(all_rows)} total classes using AI Vision!")
            
            output = BytesIO()
            df.to_excel(output, index=False, sheet_name="Schedule")
            output.seek(0)
            
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
            st.error("No valid data could be extracted.")
