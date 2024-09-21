import streamlit as st
import openai
from PyPDF2 import PdfReader
from docx import Document
import os
from dotenv import load_dotenv
import json
import pandas as pd
from io import BytesIO

# Load environment variables from .env file
load_dotenv()

# Get the API key from environment variables
api_key = os.getenv('OPENAI_API_KEY')

# Initialize the OpenAI client
client = openai

def extract_text_from_pdf(file):
    pdf_reader = PdfReader(file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    return text

def extract_text_from_docx(file):
    doc = Document(file)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text
    return text

def process_file(file, prompt):
    file_extension = os.path.splitext(file.name)[1].lower()
    if file_extension == ".pdf":
        text = extract_text_from_pdf(file)
    elif file_extension == ".docx":
        text = extract_text_from_docx(file)
    else:
        return None  # Unsupported file type
    
    # Make the OpenAI API call using client.chat.completions.create
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are a helpful assistant that processes document content."},
            {"role": "user", "content": f"Here's the content from the document:{text}Please process this content according to the following prompt:{prompt}"}
        ]
    )
    
    # Return the content of the response
    return response.choices[0].message.content

def save_to_excel(data_list):
    # Convert the list of data into a DataFrame
    df = pd.DataFrame({"Extracted Data": [json.dumps(item) for item in data_list]})

    # Create an Excel writer object and save the DataFrame to Excel in-memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Responses')

    # Set the output file's position to the start
    output.seek(0)

    return output

st.title("Invoice Extraction System")

uploaded_files = st.file_uploader("Upload PDF and/or DOCX files", type=["pdf", "docx"], accept_multiple_files=True)

prompt = st.text_area("Enter your multi-line prompt here:", height=150)

if st.button("Process Files"):
    if uploaded_files and prompt:
        with st.spinner("Processing files..."):
            results = []
            for file in uploaded_files:
                result = process_file(file, prompt)
                if result:
                    # Attempt to parse the result as JSON
                    try:
                        result_json = json.loads(result)
                        # Ensure the result is a dictionary or a list of dictionaries
                        if not isinstance(result_json, dict):
                            result_json = {"response": result_json}
                    except json.JSONDecodeError:
                        # If not valid JSON, wrap it in a dictionary to make it valid JSON
                        result_json = {"response": result}
                    
                    results.append(result_json)
                    
                    # Display the result in JSON format
                    st.write(f"### Extracted content from {file.name}:")
                    st.json(result_json)  # Always display in JSON format

        st.success("Processing complete!")
        
        # Provide an option to download the results as an Excel file
        if results:
            excel_data = save_to_excel(results)
            st.download_button(
                label="Download Excel File",
                data=excel_data,
                file_name="results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Please upload files and enter a prompt before processing.")
