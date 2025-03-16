import streamlit as st
import pandas as pd
import os
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, ColumnsAutoSizeMode
st.set_page_config(layout='wide')
st.title("Compare Multiple Job Overview Files (Showing Equipment Name)")

# 1) File Uploader for multiple Job Overview files
job_overview_files = st.file_uploader(
    "Upload Job Overview files (xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if job_overview_files:

    for job_file in job_overview_files:
        df = pd.read_excel(job_file)

        # Define the columns that must match to consider a row a duplicate
        duplicate_columns = ["Equipment Code", "Equipment Name", "Job Code", "Job Title"]

        # Identify rows where all four columns match
        duplicates_df = df[df.duplicated(subset=duplicate_columns, keep=False)].copy()

        if not duplicates_df.empty:
            # Create an expander for the file only if duplicates are found
            with st.expander(f"Duplicates in {job_file.name}"):
                st.write(
                    "Below are the rows that have duplicate Equipment Code, "
                    "Equipment Name, Job Code, and Job Title "
                    f"in **{job_file.name}**."
                )
                st.dataframe(duplicates_df)
        else:
            # If no duplicates, just display a message
            st.write(f"No duplicates found in {job_file.name}")

    # STEP A: Collect a Union of All Equipment Codes & Build File Dictionaries

    all_eq_codes = set()
    # For each file, store a dict: eq_code -> list of (eq_name, job_title)
    file_dicts = {}  # file_dicts[file_name] = (eq_code_to_pairs, sorted_list_of_codes)

    for job_file in job_overview_files:
        df = pd.read_excel(job_file)
        df['Equipment Code'] = df['Equipment Code'].astype(str).str.strip()

        eq_code_to_pairs = {}
        for _, row in df.iterrows():
            code_in_file = row['Equipment Code']
            # Some files might not have an Equipment Name for every row, so handle gracefully
            eq_name = str(row.get('Equipment Name', 'Unknown')).strip()
            job_title = str(row.get('Job Title', 'Unknown')).strip()

            # Append (Equipment Name, Job Title) to this code
            eq_code_to_pairs.setdefault(code_in_file, []).append((eq_name, job_title))

        # Keep track of all unique codes from this file
        all_eq_codes.update(eq_code_to_pairs.keys())

        # Sort the codes in this file
        sorted_codes = sorted(eq_code_to_pairs.keys())

        file_name = os.path.splitext(job_file.name)[0]
        file_dicts[file_name] = (eq_code_to_pairs, sorted_codes)

    # Convert all_eq_codes to a sorted list for consistent ordering
    master_eq_codes = sorted(all_eq_codes)

    # STEP B: Build the Comparison DataFrame
    # We'll also build a dictionary eq_code_to_all_names to store the union of eq_names
    eq_code_to_all_names = {mc: set() for mc in master_eq_codes}

    result_df = pd.DataFrame({'Equipment Code': master_eq_codes})
    file_names = list(file_dicts.keys())

    # We'll store the actual matched (eq_name, job_title) pairs for each (master_code, file_name)
    matched_map = {}

    # A progress bar so we can see how far we've gotten
    progress_bar = st.progress(0)
    total = len(master_eq_codes)

    # Create columns for each file
    for fn in file_names:
        result_df[fn] = None

    # For each master code, find how many pairs match in each file
    for i, master_code in enumerate(master_eq_codes):
        # Update progress
        progress_bar.progress((i + 1) / total)

        for fn in file_names:
            eq_code_to_pairs, sorted_codes_in_file = file_dicts[fn]

            # Gather all (eq_name, job_title) for codes that start with master_code
            matched_pairs = []
            for code_in_file, pairs in eq_code_to_pairs.items():
                if code_in_file.startswith(master_code):
                    matched_pairs.extend(pairs)

            count = len(matched_pairs)
            if count > 0:
                result_df.at[i, fn] = f"Y({count})"
            else:
                result_df.at[i, fn] = "N"

            # Store the matched pairs
            matched_map[(master_code, fn)] = matched_pairs

            # Update eq_code_to_all_names with the equipment names from these pairs
            for (ename, _) in matched_pairs:
                eq_code_to_all_names[master_code].add(ename)

    # STEP C: Add an "Equipment Name" column to show the union of eq_names
    # We'll display them as comma-separated strings
    eq_names_list = []
    for mc in master_eq_codes:
        all_names_for_mc = eq_code_to_all_names[mc]
        eq_names_list.append(", ".join(sorted(all_names_for_mc)) if all_names_for_mc else "")

    result_df.insert(1, "Equipment Name", eq_names_list)  # Insert after "Equipment Code"


    # STEP D: Mismatch Column
    def check_mismatch(row, cols):
        unique_vals = set(row[cols])
        return "Y" if len(unique_vals) > 1 else "N"


    result_df['Mismatch'] = result_df.apply(lambda r: check_mismatch(r, file_names), axis=1)

    # STEP E: Filter & Download
    show_only_mismatches = st.checkbox("Show only rows with Mismatch = Y")
    if show_only_mismatches:
        display_df = result_df[result_df['Mismatch'] == 'Y']
    else:
        display_df = result_df

    # Provide a download button for the displayed DataFrame
    csv_data = display_df.to_csv(index=False)
    st.download_button(
        label="Download Current View as CSV",
        data=csv_data,
        file_name="job_comparison.csv",
        mime="text/csv"
    )

    # STEP F: Configure & Show AgGrid
    gb = GridOptionsBuilder.from_dataframe(display_df)
    gb.configure_default_column(
        editable=False,
        groupable=True,
        enableRowGroup=True,
        aggFunc='sum',
        sortable=True,
        filter=True
    )
    gb.configure_side_bar()
    gb.configure_selection('multiple', use_checkbox=True)
    gridOptions = gb.build()

    st.write("### Equipment Code Comparison (With Equipment Name)")
    grid_response = AgGrid(
        display_df,
        gridOptions=gridOptions,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS,
        theme='material'
    )

    selected_rows = pd.DataFrame(grid_response['selected_rows'])

    # STEP G: Drill-Down for Selected Rows
    if not selected_rows.empty:
        st.write("### Drill-Down: (Equipment Name / Job Title) Pairs")
        for _, row in selected_rows.iterrows():
            eq_code = row['Equipment Code']
            eq_names_str = row['Equipment Name']

            st.subheader(f"Equipment Code: {eq_code} | Equipment Name(s): {eq_names_str}")

            for fn in file_names:
                pairs = matched_map.get((eq_code, fn), [])
                st.write(f"**{fn}** -> {len(pairs)} matches")
                # Show "Equipment Name / Job Title" for each match
                for (ename, jtitle) in pairs:
                    st.write(f"{ename} / {jtitle}")

    progress_bar.progress(1.0)  # Done
else:
    st.info("Please upload at least one Job Overview file.")
