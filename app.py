import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from io import BytesIO

# --- PASSWORD PROTECTION ---
st.title("🧪 RespIND T2 Sample Collection Summary")
password = st.text_input("Enter Password:", type="password")
if password != "RIND123":  # change password if needed
    st.warning("Please enter the correct password to access data.")
    st.stop()

# --- Load Google Sheet ---
def load_sheet(sheet_id):
    # Corrected URL for direct CSV export
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return pd.read_csv(url, on_bad_lines="skip")

# Replace this with your actual sheet ID
df = load_sheet("1Nj-jx92SdX6TOnUXT2QidIXKbdiISFL20euWgrfZPdo")

# --- Data Cleaning ---
df.columns = df.columns.str.strip().str.lower()
# Access column with lowercase name after conversion
df["submissiondate"] = pd.to_datetime(df["submissiondate"], errors="coerce").dt.tz_localize(None)

# --- Filter Today’s Data ---
today_str = pd.Timestamp.today().strftime("%Y-%m-%d")
df_today = df[df["submissiondate"].dt.strftime("%Y-%m-%d") == today_str].copy()

# --- Create Sample ID ---
# Using the 'sample_id' column directly from the provided sheet columns
df_today["sample_id"] = df_today["sample_id"]

# --- Map Type Cohort and Sample Type ---
sample_type_map = {
    "1": "NP swab",
    "2": "OP swab",
    "3": "Nasal swab",
    "4": "ET aspirate",
    "5": "Bronchial lavage /aspirate"
}

# Using 'type_swab' from the provided sheet columns for sample type mapping
df_today["sample_type"] = df_today["type_swab"].astype(str).map(sample_type_map).fillna(df_today["type_swab"])

# --- Calculate Age from p_dob ---
df_today["p_dob"] = pd.to_datetime(df_today["p_dob"], errors="coerce") # Ensure p_dob is datetime

def calculate_age_string(dob):
    if pd.isna(dob):
        return None
    today = datetime.today()
    years = today.year - dob.year
    months = today.month - dob.month
    days = today.day - dob.day

    if days < 0:
        months -= 1
    if months < 0:
        years -= 1
        months += 12

    return f"{years} yr {months} m"

df_today["calculated_age"] = df_today["p_dob"].apply(calculate_age_string)

# --- Map Gender ---
gender_map = {
    "1": "Male",
    "2": "Female"
}
df_today["mapped_gender"] = df_today["p_gender"].astype(str).map(gender_map).fillna("Other") # Handle other values as 'Other'

# --- Select Final Columns ---
table = df_today[["submissiondate", "sample_id", "sample_type", "prev_screen_no", "p_participant_id", "p_uhid", "p_child_name", "calculated_age", "mapped_gender"]].copy()
table.columns = [
    "Date & time of collection",
    "Sample ID",
    "Sample type",
    "Screening ID",
    "Participant ID",
    "UHID",
    "Name",
    "Age",
    "Sex"
]

# Add 'Fields to be filled by Virology Lab' columns
table["Received by"] = ""
table["Volume"] = ""
table["Remarks"] = ""

# Add 'S.No' column as the first column
table.insert(0, 'S.No', range(1, 1 + len(table)))

# --- Display Table ---
st.subheader("📋 Today's Sample Collection Details")
if table.empty:
    st.warning("No sample collections found for today.")
else:
    st.dataframe(table, width='stretch') # Changed use_container_width=True to width='stretch'

    # --- Download as Excel ---
    excel_filename = f"{datetime.today().strftime('%d-%m-%Y')}_RespIND T2_SampleCollection.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Today_Samples"

    # Formatting styles
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))
    bold_center = Font(bold=True)
    align_center = Alignment(horizontal="center", vertical="center")

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=table.shape[1])
    title_cell = ws.cell(row=1, column=1, value="RespIND Daily Sample Collection Summary")
    title_cell.font, title_cell.alignment = Font(bold=True, size=12), align_center

    # Header Row
    for col_num, column_title in enumerate(table.columns, 1):
        cell = ws.cell(row=2, column=col_num, value=column_title)
        cell.font = bold_center
        cell.alignment = align_center
        cell.border = border

    # Data Rows
    for row_num, row_data in enumerate(table.values, 3):
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border

    # Save Excel to buffer
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="⬇️ Download Excel",
        data=buffer,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )