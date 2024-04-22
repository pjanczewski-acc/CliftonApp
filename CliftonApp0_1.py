import streamlit as st
import pandas as pd
import os
from zipfile import ZipFile
import CliftonCorresps
import json

def zip_files(file_paths, zip_name):
    with ZipFile(zip_name, 'w') as zip:
        for file in file_paths:
            zip.write(file)
            
def download_excel(output_excel_path):
    with open(output_excel_path, "rb") as file:
        contents = file.read()
    st.download_button(label="Download", data=contents, file_name="output_excel.xlsx", mime="application/octet-stream")
            
# Main function to run the Streamlit app
def main():
    st.title("Clifton Strenghts Application")

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Process the Excel file
        output_excel_path = CliftonCorresps.create_excel(uploaded_file)
        st.markdown("### Download Output Excel File")
        st.write("Click the button below to download the output Excel file and upload it to this SharePoint location:")
        st.markdown("[SharePoint Location](https://your-sharepoint-location)")
        st.button("Download Excel", key="download_excel", on_click=download_excel(output_excel_path), args=(output_excel_path,))
        
        st.markdown("### Power BI Report")
        st.components.v1.iframe("https://app.powerbi.com/reportEmbed?reportId=your_report_id&config=your_config_id", height=600)


if __name__ == "__main__":
    main()