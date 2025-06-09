import streamlit as st
import pandas as pd
import os
import tempfile
import shutil
import win32com.client as win32
import pythoncom  # Required for COM initialisation

st.set_page_config(page_title="Excel Split & Email Sender", layout="centered")
st.title("📧 Excel Splitter and Email Sender")

st.markdown("""
Upload your input Excel file and the distribution list. This app will:
1. Split data by selected columns.
2. Save files in respective folders.
3. Match names and designations from distribution list.
4. Send emails using Outlook.
5. Let you download output files and updated distribution sheet.
""")

# Upload input files
input_file = st.file_uploader("📄 Upload Input Excel File", type=["xlsx"])
distribution_file = st.file_uploader("📋 Upload Distribution List File", type=["xlsx"])

columns_to_split = st.multiselect(
    "🧩 Select columns to split by",
    options=["AM Team Member", "AM Team Lead"]
)

if st.button("🚀 Process and Send Emails"):
    if input_file and distribution_file and columns_to_split:
        with tempfile.TemporaryDirectory() as tmpdir:
            output_folder = os.path.join(tmpdir, "output_files")
            os.makedirs(output_folder, exist_ok=True)

            # Load files
            df = pd.read_excel(input_file)
            distribution_df = pd.read_excel(distribution_file)
            distribution_df['Sent_Flag'] = 'Not Sent'

            # Split and save files
            for column in columns_to_split:
                if column in df.columns:
                    column_folder = os.path.join(output_folder, column)
                    os.makedirs(column_folder, exist_ok=True)

                    for value in df[column].dropna().unique():
                        filtered_df = df[df[column] == value]

                        safe_value = "".join(c for c in str(value) if c.isalnum() or c in (" ", ".", "_")).strip()
                        filename = f"{safe_value}.xlsx"
                        file_path = os.path.join(column_folder, filename)

                        try:
                            filtered_df.to_excel(file_path, index=False, engine='openpyxl')
                        except Exception as e:
                            st.error(f"Error saving file {filename}: {e}")

            # Initialize COM for Outlook
            try:
                pythoncom.CoInitialize()
                outlook = win32.Dispatch('outlook.application')
            except Exception as e:
                st.error(f"❌ Unable to start Outlook COM. Error: {e}")
                st.stop()

            # Email loop
            for subfolder in os.listdir(output_folder):
                subfolder_path = os.path.join(output_folder, subfolder)

                if os.path.isdir(subfolder_path):
                    matched_designations = distribution_df[
                        distribution_df['Designation'].str.strip().str.lower() == subfolder.strip().lower()
                    ]

                    for file_name in os.listdir(subfolder_path):
                        file_path = os.path.join(subfolder_path, file_name)
                        name_only = os.path.splitext(file_name)[0]

                        matched_row = matched_designations[
                            matched_designations['Name'].str.strip().str.lower() == name_only.strip().lower()
                        ]

                        if not matched_row.empty:
                            email_id = matched_row['Email_ID'].values[0]

                            mail = outlook.CreateItem(0)
                            mail.To = email_id
                            mail.Subject = f"Attached: {name_only} Data ({subfolder})"
                            mail.Body = f"Dear {name_only},\n\nPlease find the attached file.\n\nBest regards,\nYour Name"
                            mail.Attachments.Add(file_path)

                            try:
                                mail.Send()
                                distribution_df.loc[matched_row.index, 'Sent_Flag'] = 'Sent'
                                st.success(f"✅ Email sent to {email_id} with {file_name}")
                            except Exception as e:
                                st.error(f"❌ Failed to send to {email_id}: {e}")
                                distribution_df.loc[matched_row.index, 'Sent_Flag'] = 'Failed'

            # Save updated distribution list
            dist_with_flags = os.path.join(tmpdir, "Distribution_list_with_flags.xlsx")
            distribution_df.to_excel(dist_with_flags, index=False, engine='openpyxl')

            with open(dist_with_flags, 'rb') as f:
                st.download_button("⬇️ Download Updated Distribution List", f.read(), file_name="Distribution_list_with_flags.xlsx")

            # Zip output folder
            zip_path = shutil.make_archive(os.path.join(tmpdir, "output_files"), 'zip', output_folder)

            with open(zip_path, 'rb') as f:
                st.download_button("⬇️ Download All Output Files as ZIP", f.read(), file_name="output_files.zip")

    else:
        st.warning("⚠️ Please upload both Excel files and select at least one column to split by.")
