import streamlit as st
import pandas as pd
import base64
import io
# import plotly.express as px

st.set_page_config(layout="wide")
# Function to check if DataFrame contains all required columns
required_columns = [
    'Equipment Code', 'Equipment Name', 'Job Title', 'Job Type', 'Maintenance Type',
    'Primary Frequency', 'Last Done Date', 'Next Due Date', 'Secondary Frequency',
    'Last Done Hrs.', 'Next Due Hrs.', 'Discipline', 'Present Reading',
    'Rhrs/Days Since the Last Entry', 'Remaining RHrs per Days', 'Safety Level',
    'Sub Group', 'Critical to Safety', 'Risk Assessment Required', 'Forms Attached',
    'Procedures', 'P', 'Remarks', 'Job Assigned To', 'Class Reference',
    'Maintenance Cause', 'Job Priority'
]
today = pd.to_datetime('today').date()


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


series_tree = {"1": "SHIP GENERAL", "2": "HULL", "3": "EQUIPMENT FOR CARGO", "4": "SHIP EQUIPMENT",
               "5": "EQUIPMENT FOR CREW AND PASSENGERS", "6": "MAIN MACHINERY",
               "7": "SYSTEMS FOR MACHINERY MAIN COMPONENTS", "8": "SHIP COMMON SYSTEMS", "9": "VARIOUS ANALYSIS",
               "21": "HULL AFT", "23": "TANKS SPACES AND STRUCTURES", "24": "SHELL PLATES TRUNKS ETC",
               "25": "DECK HOUSES AND SUPERSTRUCTURES", "26": "HULL OUTFITTING", "27": "MATERIAL PROTECTION EXTERNAL",
               "28": "CARGO AREA", "30": "HATCHES PORTS", "31": "EQUIPMENT FOR CARGO IN HOLDS/ON DECK",
               "32": "SPECIAL CARGO HANDLING EQUIPMENT", "33": "DECK CRANES FOR CARGO",
               "35": "LOADING/DISCHARGING SYSTEMS FOR LIQUID CARGO",
               "36": "FREEZING REFRIGERATING & HEATING SYSTEMS FOR CARGO",
               "37": "GAS/VENTILATION SYSTEMS FOR CARGO HOLDS/TANKS", "38": "AUXILIARY SYSTEMS & EQUIPMENT FOR CARGO",
               "39": "OIL DISCHARGE MONITORING", "40": "MANOEUVRING MACHINERY & EQUIPMENT",
               "41": "NAVIGATION & SEARCHING EQUIPMENT", "42": "COMMUNICATION EQUIPMENTS",
               "43": "ANCHORING MOORING & TOWING EQUIPMENT",
               "44": "REP./MAINT./CLEAN. EQUIP. WORKSHOP/STORE OUTFIT NAME PLATES",
               "45": "LIFTING & TRANSPORT EQUIPMENT FOR MACHINERY COMPONENTS", "48": "SPECIAL EQUIPMENT",
               "50": "LIFESAVING PROTECTION & MEDICAL EQUIPMENT",
               "51": "INSULATION PANELS BULKHEADS DOORS SIDESCUTTLES SKYLIGHTS",
               "54": "FURNITURE INVENTORY ENTERTAINMENT EQUIPMENT",
               "55": "GALLEY/PANTRY EQUIP. PROVISION PLANTS LAUNDRY/IRONING EQU.",
               "56": "TRANSPORT EQUIPMENT FOR CREW PASSENGERS & PROVISIONS",
               "57": "VENTILATION AIR-CONDITIONING & HEATING SYSTEMS",
               "58": "SANITARY SYST. W/DISCHARGES ACCOMMODATION DRAIN SYSTEMS", "60": "DIESEL ENGINES FOR PROPULSION",
               "62": "OTHER TYPES OF PROPULSION MACHINERY", "63": "PROPELLERS TRANSMISSIONS FOILS",
               "64": "BOILERS STEAM & GAS GENERATORS", "65": "MOTOR AGGREGATES FOR MAIN ELECTRIC POWER PRODUCTION",
               "66": "OTHER AGGR. & GEN. FOR MAIN & EMERGENCY EL. POWER PRODUCTION", "70": "FUEL SYSTEMS",
               "71": "LUBE OIL SYSTEMS", "72": "COOLING SYSTEMS", "73": "COMPRESSED AIR SYSTEMS",
               "74": "EXHAUST SYSTEMS & AIR INTAKES", "75": "STEAM CONDENSATE & FEED WATER SYSTEMS",
               "76": "DISTILLED & MAKE-UP WATER SYSTEMS", "79": "AUTOMATION SYSTEMS FOR MACHINERY",
               "80": "BALLAST & BILGE SYSTEMS GUTTER PIPES OUTSIDE ACCOMMOD.",
               "81": "FIRE & LIFEBOAT ALARM FIRE FIGHTING & WASH DOWN SYSTEMS",
               "83": "SPECIAL COMMON HYDRAULIC OIL SYSTEMS", "85": "General Purpose Equipments",
               "88": "COMMON ELECTRICAL SYSTEMS", "89": "LIGHTING SYSTEM", "101": "BDP", "102": "BALLAST WATER REPORTS",
               "103": "CARGO DOUCMENTATION", "104": "CHPC – CARGO HANDLING PROCEDURES – CHEMICAL TANKERS",
               "105": "CARGO HANDLING PROCEDURES – GAS CARRIERS", "106": "CARGO HANDLING PROCEDURES – LNG CARRIERS",
               "107": "CARGO HANDLING PROCEDURES - PCC", "108": "CARGO HANDLING PROCEDURES - TANKERS",
               "109": "CARGO HANDLING PROCEDURES - WCC", "110": "COVID-19 MANAGEMENT PLAN",
               "111": "CYBER SECURITY PROCEDURES CYSM", "112": "EMERGENCY CONTINGENCY PROCEDURE",
               "113": "ENVIRONMENTAL MANAGEMENT SYSTEM", "114": "ENGINE ROOM PROCEDURES",
               "115": "EMERGENCY TOWING PROCEDURE MANUAL", "116": "GARBAGE MANAGEMENT PLAN",
               "117": "ICE CLASS VESSEL PROCEDURES", "118": "INVENTORY OF HAZARDOUS MATERIAL MANAGEMENT PLAN",
               "119": "OFFICE MANAGEMENT PROCEDURES", "120": "SHIP ENERGY EFFICIENCY MANAGEMENT PLAN",
               "121": "SHIP MANAGEMENT PROCEDURES", "122": "SHIPBOARD MARINE POLLUTION EMERGENCY PLAN",
               "123": "SHIPBOARD OIL POLLUTION EMERGENCY PLAN", "124": "SHIP TO SHIP TRANSFER OPERATION PLAN",
               "125": "VESSEL GENERAL PERMIT", "126": "VOC MANAGEMENT PLAN", "900": "LO ANALYSIS", "902": "FO ANALYSIS",
               "904": "WATER ANALYSIS", "AP": "ALL PUMPS", "CEN": "CENTRIFUGAL PUMPS", "GRP": "GEAR PUMP",
               "PDP": "PNEUMATIC DIAPHRAGM PUMP", "PIS": "PISTON PUMPS", "SCR": "SCREW PUMPS", "VAN": "VANE PUMP"}


# Streamlit app
def main():
    # st.sidebar.title("Theme Selector")
    st.title("MariApps maintenance jobs analysis")

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
            df = df[['Triggered', 'Equipment Code', 'Equipment Name', 'Job Title',
                     'Primary Frequency', 'Last Done Date', 'Next Due Date',
                     'Secondary Frequency', 'Last Done Hrs.', 'Next Due Hrs.',
                     'Discipline', 'Present Reading', 'Rhrs/Days Since the Last Entry',
                     'Remaining RHrs per Days', 'Safety Level', 'Critical to Safety',
                     'Risk Assessment Required', 'Forms Attached', 'Procedures', 'Remarks']]

            treelist = df[['Triggered', 'Equipment Code', 'Equipment Name', 'Job Title',
                           'Primary Frequency']]
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

            with st.expander('Series analysis'):
                series_analysis = treelist.groupby(['Series','Series_name']).size().reset_index(name='Job Count')
                series_analysis.index += 1  # Starting row number from 1
                # Create pie chart
                # fig = px.pie(series_analysis, values='Job Count', names='Series_name', title='Series Job Counts')

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
