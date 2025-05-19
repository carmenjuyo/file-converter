import streamlit as st
import pandas as pd
import os
import re

st.set_page_config(page_title="Multi-File RN & REV Extractor", layout="wide")
st.title("📊 Multi-File RN & REV Extractor")

# Upload first
uploaded_files = st.file_uploader("Step 1: Upload one or more .xlsx files", type="xlsx", accept_multiple_files=True)

# Only show year and column selection after upload
if uploaded_files:
    spread_type = st.radio("Step 2: What kind of data spread is this?", ["Monthly (one sheet per month)", "Yearly (one sheet for full year)"])

    if spread_type == "Yearly (one sheet for full year)":
        date_col_letter = st.text_input("Enter the Excel column letter where the month dates are listed (e.g., A)", value="A")
        date_row_start = st.number_input("Start row for date column", value=26, min_value=1, step=1)

    # Step 3: Select year
    years = [str(y) for y in range(2023, 2031)]
    selected_year = st.selectbox("Step 2: Select year to extract", options=years)

    # Step 3: Input RN and REV Excel-style cell references
    rn_cell = st.text_input(f"Step 3: Enter RN cell (e.g., E25) for {selected_year}", value="")
    rev_cell = st.text_input(f"Step 3: Enter REV cell (e.g., M25) for {selected_year}", value="")

    def cell_to_indices(cell):
        match = re.match(r"([A-Za-z]+)([0-9]+)", cell)
        if not match:
            return None, None
        col_letters, row_number = match.groups()
        col_idx = sum((ord(char.upper()) - ord('A') + 1) * (26 ** i) for i, char in enumerate(reversed(col_letters))) - 1
        row_idx = int(row_number) - 1
        return row_idx, col_idx

    rn_row_idx, rn_col_idx = cell_to_indices(rn_cell)
    rev_row_idx, rev_col_idx = cell_to_indices(rev_cell)

    if None in (rn_row_idx, rn_col_idx, rev_row_idx, rev_col_idx):
        st.warning("Please enter valid Excel-style cell references like E25 and M25.")
        st.stop()

    # Month mapping for sheet names to dates
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
            if sheet_name in month_mapping:
                month_day = month_mapping[sheet_name]
            else:
                # Use first day of year if no month keyword found
                month_day = f"01/01"
                month_day = month_mapping[sheet_name]
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

                    segment_col = df.iloc[25:, 0].dropna()
                    new_segments = [
                        str(s).strip()
                        for s in segment_col
                        if isinstance(s, str) and s.strip().upper() not in ['TOTAL', 'VS BUD 25']
                    ]
                    for seg in new_segments:
                        if seg not in segment_order:
                            segment_order.append(seg)

                    existing_keys = set().union(*[row.keys() for row in compiled_data]) if compiled_data else set()
                    for seg in new_segments:
                        if f"{seg}_RN" not in existing_keys or f"{seg}_REV" not in existing_keys:
                            for row in compiled_data:
                                row.setdefault(f"{seg}_RN", 0.0)
                                row.setdefault(f"{seg}_REV", 0.0)
                    segments = list(new_segments)

                    row = {'filename': file_name, 'date': f"{month_day}/{selected_year}"}
                    for segment in segments:
                        try:
                            seg_row_idx = df[df.iloc[:, 0].astype(str).str.strip() == segment].index[0]
                            row[f'{segment}_RN'] = float(df.iloc[seg_row_idx, rn_col_idx])
                            row[f'{segment}_REV'] = float(df.iloc[seg_row_idx, rev_col_idx])
                        except:
                            row[f'{segment}_RN'] = 0.0
                            row[f'{segment}_REV'] = 0.0
                    compiled_data.append(row)
                except Exception as e:
                    st.warning(f"Could not process sheet {sheet_name} in {file_name}: {e}")

    if compiled_data:
        final_df = pd.DataFrame(compiled_data)
        base_cols = ['filename', 'date']
        segment_cols = [f"{seg}_{suffix}" for seg in segment_order for suffix in ('RN', 'REV') if f"{seg}_{suffix}" in final_df.columns]
        extra_cols = [col for col in final_df.columns if col not in base_cols + segment_cols]
        final_df = final_df[base_cols + segment_cols + extra_cols]

        final_df['date'] = pd.to_datetime(final_df['date'], format="%d/%m/%Y")
        final_df = final_df.sort_values(by=['filename', 'date']).reset_index(drop=True)
        final_df['date'] = final_df['date'].dt.strftime("%d/%m/%Y")

        st.success("✅ Data extracted successfully!")
        st.dataframe(final_df)

        csv = final_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Combined CSV",
            data=csv,
            file_name="combined_rn_rev_data.csv",
            mime="text/csv"
        )
else:
    st.info("Step 1: Please upload one or more XLSX files to begin.")
