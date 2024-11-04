import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import os
import zipfile

# Function to create a zip file of generated PDFs
def create_zip(folder_path):
    zip_filename = f"{folder_path}.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        for foldername, subfolders, filenames in os.walk(folder_path):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zip_file.write(file_path, os.path.relpath(file_path, folder_path))
    return zip_filename

# Upload Excel file and PDFs
st.title("Invitation PDF Generator")

excel_file = st.file_uploader("Upload Guest Details Excel", type=["xlsx"])
first_page_pdf = st.file_uploader("Upload First Page PDF", type=["pdf"])
second_pages_pdf = st.file_uploader("Upload Second Pages PDF", type=["pdf"])

# Process files if uploaded
if excel_file and first_page_pdf and second_pages_pdf:
    # Load the Excel sheet
    df = pd.read_excel(excel_file)

    # Read the PDFs
    first_page_reader = PdfReader(first_page_pdf)
    second_page_reader = PdfReader(second_pages_pdf)

    # Create output folder in current directory
    output_folder = "Generated_Invitations"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Merge PDFs for each guest
    for index, row in df.iterrows():
        guest_name = row['Name_English']

        # Create a PDF writer
        pdf_writer = PdfWriter()
        pdf_writer.add_page(first_page_reader.pages[0])  # Add first page
        pdf_writer.add_page(second_page_reader.pages[index])  # Add second page

        # Save each guest's invitation
        output_pdf_path = os.path.join(output_folder, f"{guest_name}_Invitation.pdf")
        with open(output_pdf_path, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)

    st.success(f"PDFs generated and saved in {output_folder}.")
    
    # Create a zip file of the generated invitations
    zip_file_path = create_zip(output_folder)

    # Provide a download button for the zip file
    with open(zip_file_path, 'rb') as f:
        st.download_button(
            label="Download All Invitations",
            data=f,
            file_name=os.path.basename(zip_file_path),
            mime="application/zip"
        )

    # Optionally, you can add a button to delete the zip file after download
    if st.button("Delete Generated Files"):
        os.remove(zip_file_path)
        for filename in os.listdir(output_folder):
            file_path = os.path.join(output_folder, filename)
            os.remove(file_path)
        os.rmdir(output_folder)
        st.success("Generated files deleted.")
