import streamlit as st
import pandas as pd
import os
from zipfile import ZipFile

import CliftonPlots
import CliftonCorresps

import json
import openai

# Load configuration from JSON file
with open("config.json", mode="r") as f:
    config = json.load(f)

client = openai.AzureOpenAI(
        azure_endpoint=config["AZURE_ENDPOINT"],
        api_key= config["AZURE_API_KEY"],
        api_version="2023-12-01-preview")

# Function to process the uploaded Excel file and generate result files
def process_excel(input_file):
    # Load the Excel file
    
    CliftonCorresps.create_excel(input_file)

    # Save the processed Excel file
    output_excel_path = "test.xlsx"
    
    png_file_paths = CliftonPlots.create_plots()
    
    return output_excel_path, png_file_paths

def zip_files(file_paths, zip_name):
    with ZipFile(zip_name, 'w') as zip:
        for file in file_paths:
            zip.write(file)
            
# Main function to run the Streamlit app
def main():
    st.title("Clifton Strenghts Application")

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Process the Excel file
        output_excel_path, png_file_paths = process_excel(uploaded_file)

        zip_name = "result_files.zip"
        zip_files([output_excel_path] + png_file_paths, zip_name)

        with open("result_files.zip", "rb") as fp:
            btn = st.download_button(
                label="Download Results",
                data=fp,
                file_name="results.zip",
                mime="application/zip"
    )

if __name__ == "__main__":
    main()