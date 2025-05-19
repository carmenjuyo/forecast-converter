import streamlit as st
import pandas as pd
import os
from io import BytesIO

# UI setup
st.set_page_config(page_title="Multi-File RN & REV Extractor", layout="wide")
st.title("📊 Extract RN & REV from Multiple XLSX Files")

uploaded_files = st.file_uploader("Upload one or more .xlsx files", type="xlsx", accept_multiple_files=True)

# Segment definitions
segments = ['BAR', 'BARIDS', 'DSC', 'PRO', 'PROIDS', 'PRF', 'COR', 'FIT',
            'EVE', 'GEV', 'GSE', 'CRE', 'GCO', 'GIT', 'GSR', 'GPL', 'DIV']

# Month mapping
month_mapping = {
    'Janvier': '01/01', 'Fevrier': '01/02', 'Mars': '01/03', 'Avril': '01/04',
    'Mai': '01/05', 'Juin': '01/06', 'Juillet': '01/07', 'Aout': '01/08',
    'Septembre': '01/09', 'Octobre': '01/10', 'Novembre': '01/11', 'Decembre': '01/12'
}

compiled_data = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        xls = pd.ExcelFile(uploaded_file)
        file_name = os.path.splitext(uploaded_file.name)[0]

        for sheet_name, month_day in month_mapping.items():
            if sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=24, nrows=17)

                    # 2023 data (B25 = col 1, J25 = col 9)
                    row_2023 = {'filename': file_name, 'date': f"{month_day}/2023"}
                    for segment in segments:
                        try:
                            seg_row = df[df.iloc[:, 0] == segment]
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
                            seg_row = df[df.iloc[:, 0] == segment]
                            row_2024[f'{segment}_RN'] = float(seg_row.iloc[0, 2])
                            row_2024[f'{segment}_REV'] = float(seg_row.iloc[0, 10])
                        except:
                            row_2024[f'{segment}_RN'] = 0.0
                            row_2024[f'{segment}_REV'] = 0.0
                    compiled_data.append(row_2024)
                except Exception as e:
                    st.warning(f"Could not process sheet {sheet_name} in {file_name}: {e}")

    final_df = pd.DataFrame(compiled_data)
    final_df['date'] = pd.to_datetime(final_df['date'], format="%d/%m/%Y")
    final_df = final_df.sort_values(by=['filename', 'date']).reset_index(drop=True)
    final_df['date'] = final_df['date'].dt.strftime("%d/%m/%Y")

    st.success("✅ Data extracted successfully!")
    st.dataframe(final_df.head())

    csv = final_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Combined CSV",
        data=csv,
        file_name="combined_rn_rev_data.csv",
        mime="text/csv"
    )


            with l_column:
                st.subheader("🏨 Total discrepancies Rn's:")
                st.metric("RN's",f'{total_rn:n}')

            with m_column:
                st.subheader("💲 Total discrepancies REV:")
                st.metric("REV",f"€ {total_rev:n}") 

            with r_column:
                st.markdown('### 📝 Average accuracy:')
                st.metric("Revenue accuracy",f"{mean_rev:n} %") 

                st.metric("OTB accuracy",f"{mean_OTB:n} %")

            source = pd.DataFrame({
                'difference OTB': output['OTB_DIFF'],
                'date': output[f'{type_path[0]}'],
                'mean': output['OTB_%']
                })
        
            c = alt.Chart(source).mark_bar().encode(
                x='date',
                y='difference OTB',
                tooltip=['difference OTB', 'date', 'mean']
                ) 

            rule = alt.Chart(source).mark_rule(color='red').encode(
                alt.Y('mean(mean)',
                    scale=alt.Scale(domain=(1, 100)))
                )

            f = alt.layer(c, rule).resolve_scale(y='independent')

            st.subheader('Difference by day')
            st.altair_chart(f, use_container_width=True)

            output = output.drop('REV_%', axis=1)
            output = output.drop('OTB_%', axis=1)

            output.loc[:, "REV_DIFF"] = output["REV_DIFF"].map('{:.2f}'.format)
            output.loc[:, "REV_XML"] = output["REV_XML"].map('{:.2f}'.format)

            if type_juyo:
                output.rename(columns={f'{type_path[0]}': 'date'}, inplace=True)

            output

            output.to_excel("output.xlsx",index=False)

        except Exception as e:
        
            traceback.print_exc()

            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]

            st.error(f'Err4: {exc_type}; {exc_obj}; ({str(e)}), line: {exc_tb.tb_lineno}, in {fname}')
            return

    st.success('Data accuracy check done', icon='✅')
    
    my_expander = st.expander(label='Expand me for text explanation for client')
    with my_expander:
        st.write(f'''
        The data accuracy of the period {date_JUYO} till {date_last} has been reviewed. 
        
        A total of {total_rn:n} room night discrepancies and €{total_rev:n} in revenue were identified. 
        
        The average revenue accuracy percentage for this period was {mean_rev:n}%, 
        and the average OTB accuracy percentage was {mean_OTB:n}%.
        ''')
    
    with open("output.xlsx", "rb") as file:
        st.download_button(
            label="click me to download excel",
            data=file,
            file_name=f'output data accuracy.xlsx',
            mime="application/octet-stream"
            )

# Header of the page
with st.container():
    
    l_column, m_column, r_column = st.columns([3,5,1])
    
    with l_column:
        st.write("")
    
    with m_column:
        st.write(
            """
        # 📊 Data accuracy check
        """
        )
    
    with r_column:
        imagejuyo = Image.open('images/JUYO3413_Logo_Gris1.png')
        st.image(imagejuyo)

# Here will start the step-by-step process for data input.
with st.container():

    st.write("---")
    disabled = 1

    left_column, right_column = st.columns(2)

    with left_column:
        # //TODO upload hotel segment performance:

        st.header("JUYO File")
        type_juyo = st.checkbox('Upload Hotel Segment Performance by day file?')

        if type_juyo:
            type_name = 'Hotel Segment Performance by day'
            uploaded_file_JUYO = st.file_uploader("Upload Hotel Segment Performance by day", type=".xlsx")
        else:
            type_name = 'Exploration by Day'
            uploaded_file_JUYO = st.file_uploader("Upload Exploration by day", type=".xlsx")

        if uploaded_file_JUYO:
            
            JUYO_DF = pd.read_excel(uploaded_file_JUYO)
            
            st.markdown(f"### Data preview; {type_name}")
            
            st.dataframe(JUYO_DF.head())
        
        with right_column:
            st.header("XML file database")

            type = st.checkbox('Upload a Excel file instead of XML file?')
            
            if type:
                uploaded_file_XML = st.file_uploader("Upload XML file", type=".xlsx")
            else:
                uploaded_file_XML = st.file_uploader("Upload XML file", type=".XML")

            if uploaded_file_XML:
                
                if type:
                    XML_DF = pd.read_excel(uploaded_file_XML)
                else:    
                    XML_DF = run_XML_transer()

                st.markdown("### Data preview; XML file")
                st.dataframe(XML_DF.head())

    if uploaded_file_JUYO and uploaded_file_XML:

        st.write("---")

        if type_juyo:
            date_time_obj1 = JUYO_DF['category'][0]
            date_time_obj2 = JUYO_DF['category'].iloc[-1]
        else:
            obj_string = JUYO_DF['date'].iloc[-1]
            date_time_obj1 = datetime.datetime.strptime(JUYO_DF['date'][0], '%Y-%m-%d')
            date_time_obj2 = datetime.datetime.strptime(obj_string, '%Y-%m-%d')

        date_time_obj = datetime.datetime.strptime(XML_DF['CONSIDERED_DATE'][0], '%d-%b-%y')    
        
        date_JUYO = date_time_obj1.date()
        date_XML = date_time_obj.date()
        date_last = date_time_obj2.date()

        if date_JUYO == date_XML:
            if st.button('Start check'): run_check()
        else:
            st.warning(f'''
                ##### The dates the first row of both the files needs to be the same.
                ###### JUYO file = {date_JUYO} | XML file = {date_XML}\\
                Make sure that on the first row, both the dates are the same.
                ''', 
                icon='⚠️')
