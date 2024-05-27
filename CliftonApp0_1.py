import streamlit as st
import pandas as pd
import os
import CliftonCorresps
import json
import base64
from azure.storage.blob import BlobServiceClient

# Main function to run the Streamlit app

def img_to_base64(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

st.set_page_config(
    page_title="Clifton Strengths Application",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get help": "https://github.com/pjanczewski-acc/CliftonApp",
        "Report a bug": "https://github.com/pjanczewski-acc/CliftonApp",
        "About": """
            ## Clifton Strengths Application
            Description:
            The Clifton Strengths Application is a powerful tool designed to gather and analyze the strengths of team members based on the Gallup CliftonStrengths assessment. This application enables users to input the strengths of team members, generate an output Excel file, and seamlessly generate a detailed Power BI report for enhanced visualization and analysis.

            ###Key Features:

            Input Strengths: Users can input the strengths of team members into the application. The strengths are based on the results of the Gallup CliftonStrengths assessment, which identifies an individual's top strengths across various domains.
            
            Output Excel File: The application generates an output Excel file containing comprehensive information about the strengths of team members. This file serves as a structured reference for further analysis and decision-making.
            
            Azure Blob Storage Integration: Users can upload the output Excel file to Azure Blob Storage. This integration ensures that the latest data is always accessible and centrally located for collaboration and sharing within the organization.
            
            Power BI Report: Upon replacing the Excel file in Blob Storage, users can effortlessly generate a detailed Power BI report. The report provides rich visualizations and insights into the strengths of team members, allowing for in-depth analysis and exploration of trends and patterns.
            
            ###Benefits : 
            
            Ease of Use: The Clifton Strengths Application offers a user-friendly interface, making it simple for team members to input their strengths and generate reports. Its intuitive design minimizes the learning curve, allowing users to quickly navigate through the application and access the information they need.
            
            Advanced Report Making: With seamless integration with Power BI, the application enables users to create sophisticated reports with rich visualizations and insights. By leveraging the capabilities of Power BI, users can explore data trends, patterns, and correlations in a dynamic and interactive manner, enhancing decision-making and strategic planning.
            
            Elimination of Manual Work: The application automates the process of gathering and analyzing strengths data, eliminating the need for manual data entry and manipulation. By streamlining workflows and reducing manual intervention, it saves time and resources, allowing team members to focus on more strategic tasks and initiatives.
            
            Creators:
            
            Piotr Janczewski
            Iga Korneta
            Damian Gortych
              
        """
    }
)

def main():
    # Azure Blob Storage configuration
    connect_str = "DefaultEndpointsProtocol=https;AccountName=polandaidevshared;AccountKey=D6PpHJzA9AXVVsWexsIpQMf1j787boepFdJu5sI2EJi9zYOdU3oXiWq8J0Kqwz1RfomZOhKxbuq5+ASt91NArQ==;EndpointSuffix=core.windows.net"
    container_name = "polandai-dev-01"  # Replace with your container name

    # Convert image to base64
    clifton_img = img_to_base64("images/clifton.png")
    logo1_img = img_to_base64("images/logo1.png")
    logo2_img = img_to_base64("images/logo2.jpg")
    skills_img = img_to_base64("images/skills34.png")
    sample_img = img_to_base64("images/sample_input.png")

    # Apply custom CSS to make sidebar responsive
    st.markdown(
        """
        <style>
        .sidebar .sidebar-content {
            width: 25%;
            height: 100%;
            position: fixed;
            top: 0;
            left: 0;
            overflow-y: auto;
            border-right: 1px solid #d0d0d0;
            padding-top: 20px;
        }
        .sidebar .sidebar-content img {
            width: 60%;
            height: 60%;
            max-width: 100px;
            display: block;
            margin: auto;
        }
        .cover-glow {
            width: 60%;
            height: 60%;
            padding: 3px;
            left: 20%;
            box-shadow: 
                0 0 5px #000033,
                0 0 10px #000066,
                0 0 15px #000099,
                0 0 20px #0000CC,
                0 0 25px #0000FF,
                0 0 30px #3333FF,
                0 0 35px #6666FF;
            position: relative;
            z-index: -1;
            border-radius: 30px;
        }
        .cover-glow2 {
            width: 30%;
            height: 30%;
            padding: 3px;
            left: 35%;
            box-shadow: 
                0 0 5px #000033,
                0 0 10px #000066,
                0 0 15px #000099,
                0 0 20px #0000CC,
                0 0 25px #0000FF,
                0 0 30px #3333FF,
                0 0 35px #6666FF;
            position: relative;
            z-index: -1;
            border-radius: 30px;
        }
        .logo {
            width: 100px;
            height: 100px;
            border-radius: 50%;
            overflow: hidden;
            margin-top: 60px;
            box-shadow:
                0 0 5px #000033,
                0 0 10px #000066,
                0 0 15px #000099,
                0 0 20px #0000CC,
                0 0 25px #0000FF,
                0 0 30px #3333FF,
                0 0 35px #6666FF;
            position: relative;
        }
        .divider {
            border-top: 1px solid #ccc;
            border-bottom: 1px solid #fff;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        .divider2 {
            border-top: 1px solid orange;
            border-bottom: 1px solid orange;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    
    st.sidebar.markdown(
        f'<img src="data:image/png;base64,{skills_img}" class="cover-glow">',
        unsafe_allow_html=True,
    )
    st.sidebar.markdown("---")
    
    show_advanced_info = st.sidebar.checkbox("Show Detailed Instructions", value=False)
    
    st.sidebar.markdown("---")
    
    st.sidebar.markdown("""
        ### Steps
        - **Upload Input File**: Upload Excel file with single sheet containing strengths for each individual.
        - **Upload to Azure Blob Storage**: The file will be uploaded to Azure Blob Storage.
        - **Analyze The Report**: You can now use the amazing report both in the application window and in a new tab after redirection.
        """)
        
    st.sidebar.markdown("---")
    
    st.sidebar.markdown(
        f'<img src="data:image/png;base64,{logo1_img}" class="cover-glow2">',
        unsafe_allow_html=True,
    )
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.markdown("<h1 style='font-size: 60px;text-align: center'>Clifton Strengths Application</h1>", unsafe_allow_html=True)

    with col2:
        st.markdown(
        f'<img src="data:image/png;base64,{logo2_img}" class="logo"/>',
        unsafe_allow_html=True
        )
    
    with col1:
        st.markdown(
            f"<div style='text-align: center;'><img src='data:image/png;base64,{clifton_img}' style='width: 70%;' /></div>",
            unsafe_allow_html=True
        )
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2>Upload Input File</h2>", unsafe_allow_html=True)
    
    if show_advanced_info:
        st.markdown("<div class='divider2'></div>", unsafe_allow_html=True)
        st.markdown("""
        ### Required format
        - **File Extension**: xlsx
        - **File Content**: Single sheet with strengths for each individual

        Sample input to download:
        """)
        with open("sample_input.xlsx", "rb") as file:
            sample_input = file.read() 
        st.download_button(label="Download Sample Input", data=sample_input, file_name="sample_input.xlsx", mime="application/octet-stream")
        st.markdown(
            f"<div><img src='data:image/png;base64,{sample_img}' style='width: 90%;margin-bottom: 20px' /></div>",
            unsafe_allow_html=True
        )
        st.markdown("<div class='divider2'></div>", unsafe_allow_html=True)
        
    uploaded_file = st.file_uploader("", type=["xlsx"], help="See sidebar for instructions")
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    if uploaded_file is not None:
        output_excel_path = CliftonCorresps.create_excel(uploaded_file)
        
        # Upload file to Azure Blob Storage
        blob_service_client = BlobServiceClient.from_connection_string(connect_str)
        blob_client = blob_service_client.get_blob_client(container=container_name, blob="test_FactsTeam.xlsx")
        
        with open(output_excel_path, "rb") as data:
            blob_client.upload_blob(data, overwrite=True)
         
        st.markdown("<p style='font-size:24px; margin-top:5px;'>File uploaded to Azure Blob Storage successfully.</p>", unsafe_allow_html=True)

        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
        report_path = "https://app.powerbi.com/reportEmbed?reportId=c40ef5e8-87e6-4415-be0c-370e990b95fb&autoAuth=true&ctid=e0793d39-0939-496d-b129-198edd916feb"
        
        # Adding a hyperlink to the report
        st.markdown("<h2>Link to Power BI Report</h2>", unsafe_allow_html=True)
        st.markdown(f'<a href="{report_path}" target="_blank">Open Power BI Report in a new tab</a>', unsafe_allow_html=True) 
        
        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
         
        st.markdown("<h2>Power BI Report</h2>", unsafe_allow_html=True)
        st.markdown(f'<iframe title="CliftonApp.v.4.2" width="100%" height="700" src={report_path} frameborder="0" allowFullScreen="true"></iframe>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
