import streamlit as st
import pandas as pd
import os
import re

st.set_page_config(page_title="Multi-File RN & REV Extractor", layout="wide")
st.title("📊 Multi-File RN & REV Extractor")

uploaded_files = st.file_uploader("Step 1: Upload one or more .xlsx files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    date_mode = st.radio("Step 2: Is your data time-based?", ["Yes – monthly/yearly", "No – static data"])
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
        selected_year = st.selectbox("Step 3: Select year to extract", options=years)
    else:
        spread_type = None
        date_source = None
        selected_year = None

    if date_mode == "Yes – monthly/yearly" and spread_type == "Monthly (one sheet per month)":
        date_source = st.radio("Step 2c: How should we extract the date from monthly sheets?", ["From sheet name", "From a specific cell in each sheet"], key="date_src")
        if date_source == "From a specific cell in each sheet":
            date_cell_input = st.text_input("Enter the Excel-style cell that contains the date (e.g., B2)", value="B2", key="date_cell")

    if date_mode == "Yes – monthly/yearly":
        years = [str(y) for y in range(2023, 2031)]
        selected_year = st.selectbox("Step 3: Select year to extract", options=years)

    st.markdown("#### Step 4: Define the data fields you want to extract")
    num_fields = st.number_input("How many fields do you want to extract?", min_value=1, max_value=10, value=2, step=1)

    user_fields = []
    for i in range(num_fields):
        with st.expander(f"Field {i+1}"):
            label = st.text_input(f"Field name {i+1}", key=f"label_{i}")
            field_mode = st.selectbox(f"Mode for {label}", ["Single Cell", "Column Range"], key=f"mode_{i}")

            if field_mode == "Single Cell":
                cell_ref = st.text_input(f"Excel-style cell (e.g., E25) for {label}", key=f"cell_{i}")
                row_start, row_end = None, None
            else:
                column_letter = st.text_input(f"Column letter for range (e.g., E) for {label}", key=f"col_{i}")
                row_start = st.number_input(f"Start row", value=26, min_value=1, step=1, key=f"row_start_{i}")
                row_end = st.number_input(f"End row", value=37, min_value=1, step=1, key=f"row_end_{i}")
                cell_ref = column_letter

            dtype = st.selectbox(f"Data type for {label}", ["number", "text", "date"], key=f"dtype_{i}")
            user_fields.append((label, field_mode, cell_ref, dtype, row_start, row_end))

    def cell_to_indices(cell):
        match = re.match(r"([A-Za-z]+)([0-9]+)", cell)
        if not match:
            return None, None
        col_letters, row_number = match.groups()
        col_idx = sum((ord(char.upper()) - ord('A') + 1) * (26 ** i) for i, char in enumerate(reversed(col_letters))) - 1
        row_idx = int(row_number) - 1
        return row_idx, col_idx

    parsed_fields = []
    for label, mode, ref, dtype, row_start, row_end in user_fields:
        if mode == "Single Cell":
            row_idx, col_idx = cell_to_indices(ref)
            parsed_fields.append((label, mode, row_idx, col_idx, dtype, None, None))
        else:
            col_idx = cell_to_indices(ref + "1")[1] if ref else None
            parsed_fields.append((label, mode, None, col_idx, dtype, row_start, row_end))

    month_mapping = {
        'Janvier': '01/01', 'Fevrier': '01/02', 'Mars': '01/03', 'Avril': '01/04',
        'Mai': '01/05', 'Juin': '01/06', 'Juillet': '01/07', 'Aout': '01/08',
        'Septembre': '01/09', 'Octobre': '01/10', 'Novembre': '01/11', 'Decembre': '01/12',
        'January': '01/01', 'February': '01/02', 'March': '01/03', 'April': '01/04',
        'May': '01/05', 'June': '01/06', 'July': '01/07', 'August': '01/08',
        'September': '01/09', 'October': '01/10', 'November': '01/11', 'December': '01/12'
    }

    compiled_data = []
    segment_order = []

    for uploaded_file in uploaded_files:
        xls = pd.ExcelFile(uploaded_file)
        file_name = os.path.splitext(uploaded_file.name)[0]

        st.markdown(f"### Sheets in {uploaded_file.name}")
        selected_sheets = []
        for sheet in xls.sheet_names:
            if st.checkbox(f"{sheet} (from {uploaded_file.name})", value=True, key=f"{file_name}_{sheet}"):
                selected_sheets.append(sheet)

        for sheet_name in selected_sheets:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

                if spread_type == "Yearly (one sheet for full year)":
                    date_col_idx = sum((ord(char.upper()) - ord('A') + 1) * (26 ** i) for i, char in enumerate(reversed(date_col_letter))) - 1
                    extracted_date = df.iloc[date_row_start - 1:, date_col_idx].dropna().astype(str).tolist()
                    for idx, month_str in enumerate(extracted_date):
                        month_day = f"01/{month_str.zfill(2)}" if month_str.isdigit() else f"01/01"
                        base_row = {'filename': file_name, 'sheet': sheet_name, 'date': f"{month_day}/{selected_year}" if date_mode == "Yes – monthly/yearly" else ''}
                        segment_col = df.iloc[25:, 0].dropna()
                        new_segments = [
                            str(s).strip()
                            for s in segment_col
                            if isinstance(s, str) and s.strip().upper() not in ['TOTAL', 'VS BUD 25']
                        ]
                        for segment in new_segments:
                            if segment not in segment_order:
                                segment_order.append(segment)
                            row = base_row.copy()
                            try:
                                seg_row_idx = df[df.iloc[:, 0].astype(str).str.strip() == segment].index[0]
                                for label, mode, row_idx, col_idx, dtype, r_start, r_end in parsed_fields:
                                    try:
                                        val = df.iloc[seg_row_idx, col_idx] if mode == "Single Cell" else df.iloc[r_start - 1 + idx, col_idx]
                                        if dtype == "number":
                                            row[f'{segment}_{label}'] = float(val)
                                        elif dtype == "text":
                                            row[f'{segment}_{label}'] = str(val)
                                        elif dtype == "date":
                                            row[f'{segment}_{label}'] = pd.to_datetime(val).strftime("%d/%m/%Y")
                                    except:
                                        row[f'{segment}_{label}'] = 0.0 if dtype == "number" else ""
                                compiled_data.append(row)
                            except:
                                continue
                else:
                    if date_mode == "No – static data":
                        month_day = ""
                    elif date_source == "From sheet name":{'filename': file_name, 'date': f"{month_day}/{selected_year}"}
                    segment_col = df.iloc[25:, 0].dropna()
                    new_segments = [
                        str(s).strip()
                        for s in segment_col
                        if isinstance(s, str) and s.strip().upper() not in ['TOTAL', 'VS BUD 25']
                    ]
                    for segment in new_segments:
                        if segment not in segment_order:
                            segment_order.append(segment)
                        row = base_row.copy()
                        try:
                            seg_row_idx = df[df.iloc[:, 0].astype(str).str.strip() == segment].index[0]
                            for label, mode, row_idx, col_idx, dtype, r_start, r_end in parsed_fields:
                                try:
                                    val = df.iloc[seg_row_idx, col_idx] if mode == "Single Cell" else df.iloc[r_start - 1, col_idx]
                                    if dtype == "number":
                                        row[f'{segment}_{label}'] = float(val)
                                    elif dtype == "text":
                                        row[f'{segment}_{label}'] = str(val)
                                    elif dtype == "date":
                                        row[f'{segment}_{label}'] = pd.to_datetime(val).strftime("%d/%m/%Y")
                                except:
                                    row[f'{segment}_{label}'] = 0.0 if dtype == "number" else ""
                            compiled_data.append(row)
                        except:
                            continue
            except Exception as e:
                st.warning(f"Could not process sheet {sheet_name} in {file_name}: {e}")

    if compiled_data:
        final_df = pd.DataFrame(compiled_data)
        if date_mode == "Yes – monthly/yearly" and 'date' in final_df.columns:
            base_cols = ['filename', 'sheet', 'date']
            data_cols = [col for col in final_df.columns if col not in base_cols]
            final_df = final_df[base_cols + data_cols]

            final_df['date'] = pd.to_datetime(final_df['date'], format="%d/%m/%Y")
            final_df = final_df.sort_values(by=['filename', 'date']).reset_index(drop=True)
            final_df['date'] = final_df['date'].dt.strftime("%d/%m/%Y")
        else:
            final_df.drop(columns=['date'], errors='ignore', inplace=True)
            base_cols = ['filename', 'sheet']
            data_cols = [col for col in final_df.columns if col != 'filename']
            final_df = final_df[base_cols + data_cols]
        

        st.success("✅ Data extracted successfully!")
        st.dataframe(final_df)

        st.markdown("### 📋 Preview: Grouped Field Summary")
        with st.expander("See summary of all extracted fields"):
            for parsed in parsed_fields:
                label = parsed[0]
                preview_cols = [col for col in final_df.columns if col.endswith(f"_{label}")]
                if preview_cols:
                    st.markdown(f"**{label} fields**")
                    if 'date' in final_df.columns:
                        st.dataframe(final_df[preview_cols + ['date']].groupby('date').sum().reset_index())
                    else:
                        st.dataframe(final_df[preview_cols].sum().to_frame(name='Total'))

        csv = final_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Combined CSV",
            data=csv,
            file_name="combined_rn_rev_data.csv",
            mime="text/csv"
        )
else:
    st.info("Step 1: Please upload one or more XLSX files to begin.")
