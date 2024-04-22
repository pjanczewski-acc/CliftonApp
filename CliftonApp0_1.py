import streamlit as st
import pandas as pd
import os
import CliftonCorresps
import json

# Main function to run the Streamlit app
def main():
    st.title("Clifton Strenghts Application")

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Process the Excel file
        output_excel_path = CliftonCorresps.create_excel(uploaded_file)
        st.markdown("### Download Output Excel File")
        st.write("Click the button below to download the output Excel file and upload it to provided SharePoint location:")
        with open(output_excel_path, "rb") as file:
            contents = file.read()
        st.download_button(label="Download", data=contents, file_name="FactsTeam.xlsx", mime="application/octet-stream")
        st.markdown("[SharePoint Location](https://ts.accenture.com/sites/CliftonApp/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FCliftonApp%2FShared%20Documents%2FCliftonApp&viewid=b471a66b%2D616f%2D4bd1%2Db10a%2D804dfeaa698c)")
        
        st.markdown("### Power BI Report")
        
        st.markdown('<iframe title="CliftonApp.v.4.2" width="1140" height="541.25" src="https://app.powerbi.com/reportEmbed?reportId=c484c675-0935-46fb-a1fa-613daea31592&autoAuth=true&ctid=e0793d39-0939-496d-b129-198edd916feb" frameborder="0" allowFullScreen="true"></iframe>',unsafe_allow_html=True)

if __name__ == "__main__":
    main()