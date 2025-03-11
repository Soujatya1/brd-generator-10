import streamlit as st
from langchain_groq import ChatGroq
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from docx import Document
from io import BytesIO
import os
import pdfplumber

# Set page configuration
st.set_page_config(page_title="Business Requirements Document Generator", layout="wide")

# Define hardcoded BRD format
BRD_FORMAT = """
1.0 Introduction
    1.1 Purpose
    1.2 To be process / High level solution
2.0 Impact Analysis
    2.1 System impacts – Primary and cross functional
    2.2 Impacted Products
    2.3 List of APIs required
3.0 Process / Data Flow diagram / Figma
4.0 Business / System Requirement
    4.1 Application / Module Name
    4.2 Application / Module Name
5.0 MIS / DATA Requirement
6.0 Communication Requirement
7.0 Test Scenarios
8.0 Questions / Suggestions
9.0 Reference Document
10.0 Appendix
"""

# Initialize the LLM
@st.cache_resource
def initialize_llm():
    model = ChatGroq(
        groq_api_key="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", 
        model_name="Llama3-70b-8192"
    )
    
    llm_chain = LLMChain(
        llm=model, 
        prompt=PromptTemplate(
            input_variables=['requirements', 'tables', 'brd_format'],
            template="""
            Create a comprehensive Business Requirements Document (BRD) based on the following details:
            
            Document Structure:
            {brd_format}
            
            Input Document Content:
            {requirements}
            
            Tables:
            {tables}
            
            Output:
            Generate a complete, well-structured Business Requirements Document.
            """
        )
    )
    return llm_chain

# Initialize the Test Scenarios Generator
@st.cache_resource
def initialize_test_scenario_generator():
    model = ChatGroq(
        groq_api_key="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", 
        model_name="Llama3-70b-8192"
    )
    
    test_scenario_chain = LLMChain(
        llm=model, 
        prompt=PromptTemplate(
            input_variables=['brd_content'],
            template="""
            Based on the following Business Requirements Document (BRD), generate detailed test scenarios for section 7.0 Test Scenarios:
            
            BRD Content:
            {brd_content}
            
            Special Instructions for Test Scenarios Section:
            Based on the entire BRD content, generate at least 5 detailed test scenarios that would comprehensively validate the requirements. For each test scenario:
            - Provide a clear test ID and descriptive name
            - Include test objective/purpose
            - List detailed test steps
            - Define expected results/acceptance criteria
            - Specify test data requirements if applicable
            - Indicate whether it's a positive or negative test case
            - Note any dependencies or prerequisites
            """
        )
    )
    return test_scenario_chain

# File uploader
uploaded_files = st.file_uploader("Upload requirement documents (PDF/DOCX):", accept_multiple_files=True)

# Processing files
if uploaded_files:
    combined_requirements = []
    all_tables_as_text = []
    
    for uploaded_file in uploaded_files:
        file_extension = os.path.splitext(uploaded_file.name)[-1].lower()
        st.write(f"Processing {uploaded_file.name}...")
        
        if file_extension == ".docx":
            doc = Document(uploaded_file)
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            combined_requirements.append(text)
        
        elif file_extension == ".pdf":
            with pdfplumber.open(uploaded_file) as pdf:
                text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
                combined_requirements.append(text)
        
        else:
            st.warning(f"Unsupported file format: {uploaded_file.name}")
    
    st.session_state.extracted_data = {
        'requirements': "\n\n".join(combined_requirements),
        'tables': "\n\n".join(all_tables_as_text)
    }

# Generate BRD
if st.button("Generate BRD") and uploaded_files:
    if not st.session_state.extracted_data['requirements']:
        st.error("No content extracted from documents.")
    else:
        st.write("Generating BRD...")
        llm_chain = initialize_llm()
        
        prompt_input = {
            "requirements": st.session_state.extracted_data['requirements'],
            "tables": st.session_state.extracted_data['tables'],
            "brd_format": BRD_FORMAT
        }
        
        output = llm_chain.run(prompt_input)
        
        # Generate test scenarios
        test_scenario_generator = initialize_test_scenario_generator()
        test_scenarios = test_scenario_generator.run({"brd_content": output})
        
        # Insert test scenarios under the existing '7.0 Test Scenarios' heading
        output = output.replace("7.0 Test Scenarios", "7.0 Test Scenarios\n" + test_scenarios)
        
        st.success("BRD generated successfully!")
        
        # Display the output
        st.subheader("Generated Business Requirements Document")
        st.markdown(output)
        
        # Create a Word document
        doc = Document()
        doc.add_heading('Business Requirements Document', level=0)
        
        for section in output.split('\n#'):
            if not section.strip():
                continue
            
            lines = section.strip().split('\n')
            heading_text = lines[0].lstrip('#').strip()
            heading_level = 1 if section.startswith('#') else 2
            doc.add_heading(heading_text, level=heading_level)
            
            content = '\n'.join(lines[1:]).strip()
            if content:
                doc.add_paragraph(content)
        
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="Download BRD as Word document",
            data=buffer,
            file_name="Business_Requirements_Document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
