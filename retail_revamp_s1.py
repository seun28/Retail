import streamlit as st
import pandas as pd
import os
from datetime import datetime


# Inject custom CSS to change font size
st.markdown(
    """
    <style>
    /* Increase font size of file uploader label */
    .css-1cpxqw2.e1fqkh3o2 {
        font-size: 25px; /* Adjust this value to your desired font size */
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# Define the function to append new records from uploaded data to the template file
def update_template_data(uploaded_data, template_data, template_path, filename_prefix):
    # Convert template data and uploaded data to a DataFrame
    template_df = (
        pd.read_excel(template_data)
        if "xlsx" in template_path
        else pd.read_csv(template_data)
    )
    uploaded_df = (
        pd.read_excel(uploaded_data)
        if "xlsx" in uploaded_data.name
        else pd.read_csv(uploaded_data)
    )

    # Find rows in uploaded data that are not in template data based on column A
    new_rows = uploaded_df[~uploaded_df.iloc[:, 0].isin(template_df.iloc[:, 0])]

    # Append the new rows to the template DataFrame
    updated_template_df = pd.concat([template_df, new_rows], ignore_index=True)

    # Save updated template with today’s date appended
    today_str = datetime.today().strftime("%d%m%Y")
    new_filename = f"{filename_prefix}_{today_str}.xlsx"
    updated_template_path = os.path.join(os.path.dirname(template_path), new_filename)

    updated_template_df.to_excel(updated_template_path, index=False)
    return updated_template_path


# Streamlit UI
st.title("Profile and Audit Data Updater")

# Step 1: File upload section
st.header("Upload Profile and Audit Data Files")
profile_data_file = st.file_uploader("Upload Profile Data (CSV format)", type="csv")
audit_data_file = st.file_uploader("Upload Audit Data (Excel format)", type="xlsx")

# Step 2: Directory inputs
st.header("Provide Directory Paths")
database_dir = st.text_input("Paste the directory path of the Database folder")
template_dir = st.text_input("Paste the directory path of the Template folder")

# Proceed only if all required inputs are provided
if (
    st.button("Update Data Files")
    and profile_data_file
    and audit_data_file
    and database_dir
    and template_dir
):
    try:
        # Step 3: Locate template files in the Template folder
        profile_template_path = None
        audit_template_path = None
        for file in os.listdir(template_dir):
            if "ProfileData" in file and file.endswith(".xlsx"):
                profile_template_path = os.path.join(template_dir, file)
            elif "AuditData" in file and file.endswith(".xlsx"):
                audit_template_path = os.path.join(template_dir, file)

        if not profile_template_path or not audit_template_path:
            st.error(
                "Template files for ProfileData or AuditData not found in the Template directory."
            )
        else:
            # Step 4: Update Profile Data
            updated_profile_path = update_template_data(
                profile_data_file,
                profile_template_path,
                profile_template_path,
                "ProfileData",
            )
            st.success(f"Profile data successfully updated!")
            st.success(f"Updated ProfileData file saved at: {updated_profile_path}")

            # Step 7: Update Audit Data
            updated_audit_path = update_template_data(
                audit_data_file, audit_template_path, audit_template_path, "AuditData"
            )
            st.success(f"Audit data successfully updated!")
            st.success(f"Updated AuditData file saved at: {updated_audit_path}")

    except Exception as e:
        st.error(f"An error occurred: {e}")
