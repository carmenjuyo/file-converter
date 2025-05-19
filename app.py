import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Multi-File RN & REV Extractor", layout="wide")
st.title("📊 Extract RN & REV from Multiple XLSX Files")

# Year range
years = list(range(2023, 2031))
selected_year = st.sidebar.selectbox("Select year to extract", options=[str(y) for y in years])

# RN and REV column inputs for all years (can be empty)
rn_cols = {}
rev_cols = {}
for y in years:
    rn_cols[y] = st.sidebar.text_input(f"{y} RN column (1-based)", value="", key=f"rn_{y}")
    rev_cols[y] = st.sidebar.text_input(f"{y} REV column (1-based)", value="", key=f"rev_{y}")

def to_col_idx(value):
    try:
        idx = int(value) - 1
        if idx < 0:
            raise ValueError
        return idx
    except:
        return None  # invalid or empty input

rn_idx = to_col_idx(rn_cols[int(selected_year)])
rev_idx = to_col_idx(rev_cols[int(selected_year)])

uploaded_files = st.file_uploader("Upload one or more .xlsx files", type="xlsx", accept_multiple_files=True)

month_mapping = {
    'Janvier': '01/01', 'Fevrier': '01/02', 'Mars': '01/03', 'Avril': '01/04',
    'Mai': '01/05', 'Juin': '01/06', 'Juillet': '01/07', 'Aout': '01/08',
    'Septembre': '01/09', 'Octobre': '01/10', 'Novembre': '01/11', 'Decembre': '01/12'
}

compiled_data = []
segment_order = []

if uploaded_files:
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

                    existing_keys = set().union(*[row.keys() for row in compiled_data])
                    for seg in new_segments:
                        if f"{seg}_RN" not in existing_keys or f"{seg}_REV" not in existing_keys:
                            for row in compiled_data:
                                row.setdefault(f"{seg}_RN", 0.0)
                                row.setdefault(f"{seg}_REV", 0.0)
                    segments = list(new_segments)

                    if rn_idx is None or rev_idx is None:
                        st.warning(f"RN or REV column not defined for year {selected_year}, skipping extraction.")
                        continue

                    row = {'filename': file_name, 'date': f"{month_day}/{selected_year}"}
                    for segment in segments:
                        try:
                            seg_row = df[df.iloc[:, 0].astype(str).str.strip() == segment]
                            row[f'{segment}_RN'] = float(seg_row.iloc[0, rn_idx])
                            row[f'{segment}_REV'] = float(seg_row.iloc[0, rev_idx])
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
        st.warning("⚠️ No data extracted — please check your inputs and selections.")
