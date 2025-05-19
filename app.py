import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Multi-File RN & REV Extractor", layout="wide")
st.title("📊 Multi-File RN & REV Extractor")

# Upload first
uploaded_files = st.file_uploader("Step 1: Upload one or more .xlsx files", type="xlsx", accept_multiple_files=True)

# Only show year and column selection after upload
if uploaded_files:
    # Step 2: Select year
    years = [str(y) for y in range(2023, 2031)]
    selected_year = st.selectbox("Step 2: Select year to extract", options=years)

    # Step 3: Input RN and REV columns for selected year
    rn_col_input = st.text_input(f"Step 3: Enter RN column number for {selected_year} (1-based)", value="")
    rev_col_input = st.text_input(f"Step 3: Enter REV column number for {selected_year} (1-based)", value="")
rn_col_input = st.text_input(f"Step 2: Enter RN column number for {selected_year} (1-based)", value="")
rev_col_input = st.text_input(f"Step 2: Enter REV column number for {selected_year} (1-based)", value="")

def to_col_idx(value):
    try:
        idx = int(value) - 1
        if idx < 0:
            raise ValueError
        return idx
    except:
        return None

rn_col_idx = to_col_idx(rn_col_input)
rev_col_idx = to_col_idx(rev_col_input)

    rn_col_idx = to_col_idx(rn_col_input)
    rev_col_idx = to_col_idx(rev_col_input)

    if rn_col_idx is None or rev_col_idx is None:
        st.warning("Please enter valid RN and REV column numbers to continue.")
        st.stop()

    # Month mapping for sheet names to dates
month_mapping = {
    'Janvier': '01/01', 'Fevrier': '01/02', 'Mars': '01/03', 'Avril': '01/04',
    'Mai': '01/05', 'Juin': '01/06', 'Juillet': '01/07', 'Aout': '01/08',
    'Septembre': '01/09', 'Octobre': '01/10', 'Novembre': '01/11', 'Decembre': '01/12'
}

uploaded_files = st.file_uploader("Step 3: Upload one or more .xlsx files", type="xlsx", accept_multiple_files=True)

compiled_data = []
segment_order = []

if uploaded_files:
    if rn_col_idx is None or rev_col_idx is None:
        st.warning("Please enter valid RN and REV column numbers to continue.")
        st.stop()
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
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=24)
                    segment_col = df.iloc[:, 0].dropna()
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
                            seg_row = df[df.iloc[:, 0].astype(str).str.strip() == segment]
                            row[f'{segment}_RN'] = float(seg_row.iloc[0, rn_col_idx])
                            row[f'{segment}_REV'] = float(seg_row.iloc[0, rev_col_idx])
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
    st.info("Please upload one or more XLSX files to begin extraction.")
