import streamlit as st
import pandas as pd
import base64, io
from data import series_tree, jm_codes
#import plotly.express as px

st.set_page_config(layout="wide")
# Function to check if DataFrame contains all required columns
required_columns = [
    'Equipment Code', 'Equipment Name', 'Job Title', 'Job Type']
    # , 'Maintenance Type',
    # 'Primary Frequency', 'Last Done Date', 'Next Due Date', 'Secondary Frequency',
    # 'Last Done Hrs.', 'Next Due Hrs.', 'Discipline', 'Present Reading',
    # 'Rhrs/Days Since the Last Entry', 'Remaining RHrs per Days', 'Safety Level',
    # 'Sub Group', 'Critical to Safety', 'Risk Assessment Required', 'Forms Attached',
    # 'Procedures', 'P', 'Remarks', 'Job Assigned To', 'Class Reference',
    # 'Maintenance Cause', 'Job Priority'

today = pd.to_datetime('today').date()
jobmaster=pd.read_excel('JM.xlsx')
jobcodes=jobmaster['Job Code'].tolist()
# st.text(jobcodes)

def check_columns(df):
    return all(col in df.columns for col in required_columns)


def download_excel(df, filename):
    # Create a BytesIO object to write the Excel data to
    output = io.BytesIO()
    # Use the ExcelWriter class to write DataFrame to Excel file
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    # Set the BytesIO object's position to the start
    output.seek(0)
    # Convert the BytesIO object to base64 encoding
    excel_data = output.read()
    b64 = base64.b64encode(excel_data).decode()
    # Generate the download link with base64 encoded Excel data
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx">Export {filename} to Excel</a>'
    return href


# Streamlit app
def main():
    # st.sidebar.title("Theme Selector")
    st.title("Maintenance jobs analysis")

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Read Excel file into DataFrame
        try:
            df = pd.read_excel(uploaded_file)

        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            return

        # Check if DataFrame contains required columns
        if check_columns(df):
            st.success(f"Processing....{df.shape[0]} jobs found")
            df = df[df['Primary Frequency'] != '0 EVENT']
            df['Triggered'] = df['Last Done Date'].apply(lambda x: 'Y' if pd.notnull(x) else 'N')
            df = df[['Triggered'] + [col for col in df.columns if col not in ['Vessel', 'Triggered']]]
            df['Last Done Date'] = pd.to_datetime(df['Last Done Date'])
            df['Next Due Date'] = pd.to_datetime(df['Next Due Date'], errors='coerce')
            # df = df[['Triggered', 'Equipment Code', 'Equipment Name', 'Job Title',
            #          'Primary Frequency', 'Last Done Date', 'Next Due Date',
            #          'Secondary Frequency', 'Last Done Hrs.', 'Next Due Hrs.',
            #          'Discipline', 'Present Reading', 'Rhrs/Days Since the Last Entry',
            #          'Remaining RHrs per Days', 'Safety Level', 'Critical to Safety',
            #          'Risk Assessment Required', 'Forms Attached', 'Procedures', 'Remarks']]

            treelist = df[['Triggered', 'Equipment Code', 'Equipment Name', 'Job Title','Primary Frequency']]
            treelist['Series'] = treelist['Equipment Code'].str[0] # Extract the first digit
            treelist['Sub_Series'] = treelist['Equipment Code'].str[:2]
            treelist['Series_name'] = treelist['Series'].map(series_tree)
            treelist['Sub_Series_name'] = treelist['Sub_Series'].map(series_tree)
            treelist.insert(treelist.columns.get_loc('Equipment Code') + 1, 'Series', treelist.pop('Series'))
            treelist.insert(treelist.columns.get_loc('Equipment Code') + 2, 'Sub_Series', treelist.pop('Sub_Series'))
            # st.write(treelist)
            overdue_jobs = df[df['Next Due Date'].dt.date < today]
            duplicate_rows = df[df.duplicated(subset=['Equipment Code', 'Equipment Name', 'Job Title'], keep=False)]
            duplicate_count = duplicate_rows.shape[0]
            triggered_jobs = df[df['Triggered'] == 'Y']
            triggered_jobs_count = triggered_jobs.shape[0]
            non_triggered_jobs = df[df['Triggered'] == 'N']
            non_triggered_jobs_count = non_triggered_jobs.shape[0]
            critical_equipment_df = df[df['Safety Level'] == 'CRITICAL']
            critical_equipment_list = critical_equipment_df['Equipment Name'].unique()
            critical_jobs_df = df[(df['Safety Level'] == 'CRITICAL') & (df['Critical to Safety'] == 'YES')]
            critical_jobs_list = critical_jobs_df[['Equipment Name', 'Job Title']]

            if 'Job Code' in df.columns:
                # New code to filter dataframe based on job codes not in jm_codes
                jobs_not_in_jobmaster = df[~df['Job Code'].isin(jm_codes)]
                jobs_not_in_jobmaster_count=jobs_not_in_jobmaster.shape[0]
                with st.expander(f'{jobs_not_in_jobmaster_count} Jobs not linked to Central Library'):
                    st.dataframe(jobs_not_in_jobmaster)

            else:
                st.error("Please enable job codes and download excel")

            with st.expander('Raw Data'):
                # Multiselect widget to select columns to display
                columns_to_display = st.multiselect("Select columns to display", df.columns.tolist(),
                                                    default=df.columns.tolist())

                # Display the DataFrame with selected columns
                st.dataframe(df[columns_to_display])
            with st.expander('Series analysis'):
                series_analysis = treelist.groupby(['Series','Series_name']).size().reset_index(name='Job Count')
                series_analysis.index += 1  # Starting row number from 1
                #Create pie chart
                #fig = px.pie(series_analysis, values='Job Count', names='Series_name', title='Series Job Counts')

                # Display DataFrame and pie chart in two columns
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.dataframe(series_analysis)
                with col2:
                    # st.plotly_chart(fig)

            with st.expander('Critical Items'):
                col1, col2 = st.columns(2)
                with col1:
                    st.header('Critical Equipment List')

                    st.table(critical_equipment_list)
                with col2:
                    st.header('Critical Jobs')
                    download_link_critical_jobs = download_excel(critical_jobs_list, 'Critical_Jobs')
                    st.markdown(download_link_critical_jobs, unsafe_allow_html=True)
                    st.dataframe(critical_jobs_list)

            with st.expander(f"{overdue_jobs.shape[0]} Overdue jobs"):
                download_link_overdue_jobs = download_excel(overdue_jobs, 'Overdue_Jobs')
                st.markdown(download_link_overdue_jobs, unsafe_allow_html=True)
                st.dataframe(overdue_jobs, width=None)

            with st.expander(f"{triggered_jobs_count} Triggered jobs"):
                download_link_triggered_jobs = download_excel(triggered_jobs, 'Triggered_Jobs')
                st.markdown(download_link_triggered_jobs, unsafe_allow_html=True)
                st.dataframe(triggered_jobs, width=None)

            with st.expander(f"{non_triggered_jobs_count} Non Triggered jobs"):
                download_link_non_triggered_jobs = download_excel(non_triggered_jobs, 'Non_Triggered_Jobs')
                st.markdown(download_link_non_triggered_jobs, unsafe_allow_html=True)
                st.dataframe(non_triggered_jobs, width=None)

            with st.expander(f"{duplicate_count} Duplicate Jobs"):
                download_link_duplicate_jobs = download_excel(duplicate_rows, 'Duplicate_Jobs')
                st.markdown(download_link_duplicate_jobs, unsafe_allow_html=True)
                # Display duplicate rows, if any
                if not duplicate_rows.empty:
                    st.dataframe(duplicate_rows, width=None)
                else:
                    st.write("No duplicate Jobs found.")

            # st.dataframe(df)
        else:
            st.error("Uploaded file does not contain all required columns.")
            st.write("Please upload a valid file")


if __name__ == "__main__":
    main()
