import streamlit as st
import pandas as pd
import os
from io import BytesIO

# UI setup
st.set_page_config(page_title="Multi-File RN & REV Extractor", layout="wide")
st.title("📊 Extract RN & REV from Multiple XLSX Files")

uploaded_files = st.file_uploader("Upload one or more .xlsx files", type="xlsx", accept_multiple_files=True)

# Month mapping
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

        for sheet_name, month_day in month_mapping.items():
            if sheet_name in xls.sheet_names:
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
                    # Dynamically track any new segments
                    for seg in new_segments:
                        if f"{seg}_RN" not in existing_keys or f"{seg}_REV" not in existing_keys:
                            for row in compiled_data:
                                row.setdefault(f"{seg}_RN", 0.0)
                                row.setdefault(f"{seg}_REV", 0.0)
                    segments = list(new_segments)

                    # 2023 data (B25 = col 1, J25 = col 9)
                    row_2023 = {'filename': file_name, 'date': f"{month_day}/2023"}
                    for segment in segments:
                        try:
                            seg_row = df[df.iloc[:, 0].astype(str).str.strip() == segment]
                            row_2023[f'{segment}_RN'] = float(seg_row.iloc[0, 1])
                            row_2023[f'{segment}_REV'] = float(seg_row.iloc[0, 9])
                        except:
                            row_2023[f'{segment}_RN'] = 0.0
                            row_2023[f'{segment}_REV'] = 0.0
                    compiled_data.append(row_2023)

                    # 2024 data (C25 = col 2, K25 = col 10)
                    row_2024 = {'filename': file_name, 'date': f"{month_day}/2024"}
                    for segment in segments:
                        try:
                            seg_row = df[df.iloc[:, 0].astype(str).str.strip() == segment]
                            row_2024[f'{segment}_RN'] = float(seg_row.iloc[0, 2])
                            row_2024[f'{segment}_REV'] = float(seg_row.iloc[0, 10])
                        except:
                            row_2024[f'{segment}_RN'] = 0.0
                            row_2024[f'{segment}_REV'] = 0.0
                    compiled_data.append(row_2024)

                    # 2025 data (E25 = col 4, M25 = col 12)
                    row_2025 = {'filename': file_name, 'date': f"{month_day}/2025"}
                    for segment in segments:
                        try:
                            seg_row = df[df.iloc[:, 0].astype(str).str.strip() == segment]
                            row_2025[f'{segment}_RN'] = float(seg_row.iloc[0, 4])
                            row_2025[f'{segment}_REV'] = float(seg_row.iloc[0, 12])
                        except:
                            row_2025[f'{segment}_RN'] = 0.0
                            row_2025[f'{segment}_REV'] = 0.0
                    compiled_data.append(row_2025)
                except Exception as e:
                    st.warning(f"Could not process sheet {sheet_name} in {file_name}: {e}")

    final_df = pd.DataFrame(compiled_data)
    # Reorder columns to match segment_order followed by any extras
    base_cols = ['filename', 'date']
    segment_cols = [f"{seg}_{suffix}" for seg in segment_order for suffix in ('RN', 'REV') if f"{seg}_{suffix}" in final_df.columns]
    extra_cols = [col for col in final_df.columns if col not in base_cols + segment_cols]
    final_df = final_df[base_cols + segment_cols + extra_cols]
    if final_df.empty:
        st.warning("⚠️ No data extracted — please check your file formatting and sheet structure.")
    else:
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
