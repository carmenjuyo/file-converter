import streamlit as st
import pandas as pd
import os
import re

st.set_page_config(page_title="Multi-File RN & REV Extractor", layout="wide")
st.title("📊 Multi-File RN & REV Extractor")

uploaded_files = st.file_uploader("Step 1: Upload one or more .xlsx files", type="xlsx", accept_multiple_files=True)

def cell_to_indices(cell):
    match = re.match(r"([A-Za-z]+)([0-9]+)", cell)
    if not match:
        return None, None
    col_letters, row_number = match.groups()
    col_idx = sum((ord(char.upper()) - ord('A') + 1) * (26 ** i) for i, char in enumerate(reversed(col_letters))) - 1
    row_idx = int(row_number) - 1
    return row_idx, col_idx

if uploaded_files:
    date_mode = st.radio("Step 2: Is your data time-based?", ["Yes – monthly/yearly", "No – static data"])

    spread_type = None
    date_source = None
    selected_years = []
    date_cell_input = None

    all_sheet_options = {}
    for file in uploaded_files:
        excel = pd.ExcelFile(file)
        all_sheet_options[file.name] = excel.sheet_names

    sheet_selection = {}
    for file in uploaded_files:
        sheet_selection[file.name] = st.multiselect(f"Select sheets to process from {file.name}", options=all_sheet_options[file.name])

    if date_mode == "Yes – monthly/yearly":
        spread_type = st.radio("Step 2a: What kind of data spread is this?", ["Monthly (one sheet per month)", "Yearly (one sheet for full year)"])

        if spread_type == "Yearly (one sheet for full year)":
            date_col_letter = st.text_input("Enter the Excel column letter where the month dates are listed (e.g., A)", value="A")
            date_row_start = st.number_input("Start row for date column", value=26, min_value=1, step=1)
        elif spread_type == "Monthly (one sheet per month)":
            date_source = st.radio("How should we extract the date from monthly sheets?", ["From sheet name", "From a specific cell in each sheet"])
            if date_source == "From a specific cell in each sheet":
                date_cell_input = st.text_input("Enter the Excel-style cell that contains the date (e.g., B2)", value="B2")

        years = [str(y) for y in range(2023, 2031)]
        selected_years = st.multiselect("Step 3: Select year(s) to extract", options=years, default=[years[0]])

    st.markdown("#### Step 4: Define the data fields you want to extract")
    num_fields = st.number_input("How many fields do you want to extract?", min_value=1, max_value=10, value=2, step=1)

    user_fields = []
    for i in range(num_fields):
        with st.expander(f"Field {i+1}"):
            label = st.text_input(f"Field name {i+1}", key=f"label_{i}")
            field_mode = st.selectbox(f"Mode for {label}", ["Single Cell", "Column Range"], key=f"mode_{i}")
            field_scope = st.selectbox(f"Is {label} present in all files or only some?", ["All files", "Only specific files"], key=f"scope_{i}")

            if field_scope == "Only specific files":
                field_files = st.multiselect(f"Select files that contain {label}", options=[f.name for f in uploaded_files], key=f"files_{i}")
            else:
                field_files = [f.name for f in uploaded_files]

            if field_mode == "Single Cell":
                cell_ref = st.text_input(f"Excel-style cell (e.g., E25) for {label}", key=f"cell_{i}")
                row_start, row_end, until_end = None, None, False
            else:
                column_letter = st.text_input(f"Column letter for range (e.g., E) for {label}", key=f"col_{i}")
                row_start = st.number_input(f"Start row", value=26, min_value=1, step=1, key=f"row_start_{i}")
                until_end = st.checkbox(f"Until end of rows?", key=f"until_end_{i}")
                row_end = None if until_end else st.number_input(f"End row", value=37, min_value=1, step=1, key=f"row_end_{i}")
                cell_ref = column_letter

            dtype = st.selectbox(f"Data type for {label}", ["number", "text", "date"], key=f"dtype_{i}")
            user_fields.append((label, field_mode, cell_ref, dtype, row_start, row_end, until_end, field_scope, field_files))

    aggregation_enabled = st.checkbox("Step 5: Do you want to aggregate the extracted data?")
    if aggregation_enabled:
        group_field = st.selectbox("Select a field to group by", [f[0] for f in user_fields], key="agg_field")
        agg_func = st.selectbox("Aggregation function for numeric fields", ["sum", "mean", "first", "last"], key="agg_func")

    parsed_fields = []
    for label, mode, ref, dtype, row_start, row_end, until_end, scope, field_files in user_fields:
        if mode == "Single Cell":
            row_idx, col_idx = cell_to_indices(ref)
            parsed_fields.append((label, mode, row_idx, col_idx, dtype, None, None, until_end, scope, field_files))
        else:
            col_idx = cell_to_indices(ref + "1")[1] if ref else None
            parsed_fields.append((label, mode, None, col_idx, dtype, row_start, row_end, until_end, scope, field_files))

    month_mapping = {
        'Janvier': '01/01', 'Fevrier': '01/02', 'Mars': '01/03', 'Avril': '01/04',
        'Mai': '01/05', 'Juin': '01/06', 'Juillet': '01/07', 'Aout': '01/08',
        'Septembre': '01/09', 'Octobre': '01/10', 'Novembre': '01/11', 'Decembre': '01/12',
        'January': '01/01', 'February': '01/02', 'March': '01/03', 'April': '01/04',
        'May': '01/05', 'June': '01/06', 'July': '01/07', 'August': '01/08',
        'September': '01/09', 'October': '01/10', 'November': '01/11', 'December': '01/12'
    }

    if st.button("🛠️ Start Extraction"):
        compiled_data = []

        for file in uploaded_files:
            excel = pd.ExcelFile(file)
            file_name = os.path.splitext(file.name)[0]
            sheet_names = sheet_selection[file.name]

            for sheet_name in sheet_names:
                try:
                    df = pd.read_excel(excel, sheet_name=sheet_name, header=None)

                    max_start = max([pf[5] for pf in parsed_fields if pf[1] == "Column Range" and pf[5] is not None] + [1])
                    max_end_candidates = [pf[6] for pf in parsed_fields if pf[1] == "Column Range" and pf[6] is not None]
                    max_end = max(max_end_candidates) if max_end_candidates else df.shape[0]

                    for row in range(max_start - 1, max_end):
                        row_data = {'filename': file_name, 'sheet': sheet_name}

                        if date_mode == "Yes – monthly/yearly":
                            if spread_type == "Monthly (one sheet per month)":
                                if date_source == "From sheet name":
                                    for month_label, month_day in month_mapping.items():
                                        if month_label.lower() in sheet_name.lower():
                                            if selected_years:
                                                row_data['date'] = f"{month_day}/{selected_years[0]}"
                                            break
                                elif date_source == "From a specific cell in each sheet":
                                    row_idx, col_idx = cell_to_indices(date_cell_input)
                                    row_data['date'] = str(df.iat[row_idx, col_idx])
                            elif spread_type == "Yearly (one sheet for full year)":
                                date_col = cell_to_indices(date_col_letter + "1")[1]
                                row_data['date'] = str(df.iat[row, date_col])

                        for label, mode, r_idx, c_idx, dtype, r_start, r_end, until_end, scope, files_applied in parsed_fields:
                            try:
                                if file.name not in files_applied:
                                    row_data[label] = None
                                    continue
                                if mode == "Single Cell":
                                    row_data[label] = df.iat[r_idx, c_idx]
                                else:
                                    row_data[label] = df.iat[row, c_idx]
                            except:
                                row_data[label] = None

                        compiled_data.append(row_data)

                except Exception as e:
                    st.warning(f"Could not process sheet {sheet_name} in {file_name}: {e}")

        if compiled_data:
            df_out = pd.DataFrame(compiled_data)

            if aggregation_enabled and group_field in df_out.columns:
                numeric_cols = df_out.select_dtypes(include='number').columns.tolist()
                group_fields = ['filename', 'sheet']
                if 'date' in df_out.columns:
                    group_fields.append('date')
                group_fields.append(group_field)
                df_out = df_out.groupby(group_fields, as_index=False).agg({col: agg_func for col in numeric_cols if col not in group_fields})

            st.dataframe(df_out.head(100))
            st.download_button("⬇️ Download Combined CSV", df_out.to_csv(index=False).encode("utf-8"), file_name="compiled_output.csv", mime="text/csv")
        else:
            st.error("No data was extracted. Please review your settings and files.")
