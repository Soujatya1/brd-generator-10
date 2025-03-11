import streamlit as st
from langchain_groq import ChatGroq
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from docx import Document
from io import BytesIO
import hashlib
import os
import pdfplumber
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import difflib
import time

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
            input_variables=['requirements', 'tables'],
            template="""
            Create a comprehensive Business Requirements Document (BRD) based on the following details:

            Document Structure:
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

            Requirements:
            Analyze the content provided in the requirement documents and map the relevant information to each section defined in the BRD structure. Be detailed, specific, and maintain professional language.

            For each section:
            - Provide comprehensive descriptions that fully capture the requirements, not just summaries
            - Include relevant technical details, business context, and implementation considerations
            - Connect different sections to show how they relate to each other
            - Use specific language rather than generic statements

            Tables:
            If applicable, include the following tabular information extracted from the documents:
            {tables}

            Special Instructions for Test Scenarios Section:
            Based on the entire BRD content, generate at least 10 detailed test scenarios that would comprehensively validate the requirements. For each test scenario:
            - Provide a clear test ID and descriptive name
            - Include test objective/purpose
            - List detailed test steps
            - Define expected results/acceptance criteria
            - Specify test data requirements if applicable
            - Indicate whether it's a positive or negative test case
            - Note any dependencies or prerequisites

            Formatting Guidelines:
            1. Use proper heading levels (# for main sections, ## for subsections)
            2. Include bullet points or numbered lists for better readability
            3. Clearly differentiate between functional and non-functional requirements
            4. Maintain consistent formatting throughout the document
            5. Keep page number references accurate
            6. Include tables where information is best presented in tabular format

            Requirements Processing:
            1. Categorize extracted information based on the BRD structure
            2. Provide detailed explanations for each section, not just brief summaries
            3. Include explicit section headings with page numbers as shown above
            4. Ensure all content from input documents is represented in the appropriate sections
            5. When information is missing for a section, provide a reasonable placeholder based on available context

            Input Document Content:
            {requirements}

            Output:
            Generate a complete, well-structured Business Requirements Document following the exact structure provided, with appropriate page numbers. The document should be comprehensive, descriptive, professionally formatted, and ready for business use. Pay special attention to creating realistic and relevant test scenarios based on the BRD content.
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
            Based on the following Business Requirements Document (BRD), create a comprehensive set of test scenarios for section 7.0 Test Scenarios:

            BRD Content:
            {brd_content}

            Create at least 5 detailed test scenarios that would thoroughly validate all the requirements in this BRD. 
            For each test scenario, include:

            1. Test ID (TS-XXX format)
            2. Test Name (descriptive title)
            3. Test Category (Functional, Non-functional, Integration, etc.)
            4. Test Priority (High, Medium, Low)
            5. Test Objective (What is being tested and why)
            6. Preconditions (What must be in place before testing)
            7. Test Steps (Numbered steps for execution)
            8. Expected Results (What should happen if the test passes)
            9. Test Data Requirements (Sample data needed)
            10. Related Requirements (Which requirements from the BRD this test validates)

            Include a mix of:
            - Positive test cases (verifying things work correctly)
            - Negative test cases (verifying error handling works properly)
            - Boundary test cases (testing at the limits of acceptable values)
            - Performance test cases (if applicable)
            - Security test cases (if applicable)
            - Integration test cases (testing interactions between components)

            Format each test scenario as a clear, well-structured section with all components above.
            Your test scenarios should be realistic, specific to the requirements in the BRD, and provide adequate coverage of all key functionality.
            """
        )
    )
    
    return test_scenario_chain

# Extract content from various document formats
def extract_content_from_docx(file):
    doc = Document(file)
    paragraphs = []
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            paragraphs.append(paragraph.text)
    
    text = "\n".join(paragraphs)
    
    # Process tables
    tables = []
    for i, table in enumerate(doc.tables):
        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
        
        # Format table more cleanly
        formatted_table = f"Table {i+1}:\n"
        formatted_table += "\n".join([" | ".join(row) for row in table_data])
        tables.append(formatted_table)
    
    tables_as_text = "\n\n".join(tables)
    return text, tables_as_text

def extract_text_from_pdf(file):
    text_content = []
    tables_content = []
    
    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text() or ""
            if page_text:
                text_content.append(f"--- Page {i+1} ---\n{page_text}")
            
            # Extract tables
            tables = page.extract_tables()
            for j, table in enumerate(tables):
                table_text = f"Table {i+1}.{j+1}:\n"
                table_text += "\n".join([" | ".join([str(cell) if cell else "" for cell in row]) for row in table])
                tables_content.append(table_text)
    
    text = "\n\n".join(text_content)
    tables = "\n\n".join(tables_content)
    return text, tables

# Analyze content to identify key topics
def analyze_content_for_topics(text):
    # This function extracts key topics from the content
    # This could be enhanced with more sophisticated NLP techniques
    
    # Simple keyword extraction (this is a placeholder for more advanced analysis)
    keywords = []
    lines = text.split('\n')
    for line in lines:
        # Look for potential headings or key terms
        if line.strip() and len(line.strip()) < 100 and not line.startswith('---'):
            if any(char.isalpha() for char in line):
                keywords.append(line.strip())
    
    # Return unique keywords, limited to 20
    unique_keywords = list(set(keywords))
    return unique_keywords[:20]

# UI Components
st.title("AI_Powered BRD Generator")
st.write("Upload requirement documents to generate a professionally formatted Business Requirements Document.")

# File uploader
uploaded_files = st.file_uploader("Upload requirement documents (PDF/DOCX):", accept_multiple_files=True)

# Processing files
if uploaded_files:
    with st.spinner("Processing uploaded documents..."):
        if "extracted_data" not in st.session_state:
            combined_requirements = []
            all_tables_as_text = []
            
            for uploaded_file in uploaded_files:
                file_extension = os.path.splitext(uploaded_file.name)[-1].lower()
                st.write(f"Processing {uploaded_file.name}...")
                
                if file_extension == ".docx":
                    text, tables_as_text = extract_content_from_docx(uploaded_file)
                    if text:
                        combined_requirements.append(f"--- From {uploaded_file.name} ---\n{text}")
                    if tables_as_text:
                        all_tables_as_text.append(f"--- Tables from {uploaded_file.name} ---\n{tables_as_text}")
                
                elif file_extension == ".pdf":
                    text, tables = extract_text_from_pdf(uploaded_file)
                    if text:
                        combined_requirements.append(f"--- From {uploaded_file.name} ---\n{text}")
                    if tables:
                        all_tables_as_text.append(f"--- Tables from {uploaded_file.name} ---\n{tables}")
                
                else:
                    st.warning(f"Unsupported file format: {uploaded_file.name}")
            
            st.session_state.extracted_data = {
                'requirements': "\n\n".join(combined_requirements),
                'tables': "\n\n".join(all_tables_as_text)
            }
            
            # Analyze content for key topics
            if combined_requirements:
                all_text = "\n\n".join(combined_requirements)
                key_topics = analyze_content_for_topics(all_text)
                st.session_state.key_topics = key_topics
                
                # Show extraction success
                st.success("Documents successfully processed")
    
    # Preview extracted data
    with st.expander("Preview Extracted Content", expanded=False):
        st.subheader("Extracted Text")
        st.text_area("Text Content", st.session_state.extracted_data['requirements'], height=200)
        
        st.subheader("Extracted Tables")
        st.text_area("Table Content", st.session_state.extracted_data['tables'], height=200)
        
        if hasattr(st.session_state, 'key_topics'):
            st.subheader("Key Topics Identified")
            st.write(", ".join(st.session_state.key_topics))

# Caching mechanism
def generate_hash(requirements, tables, detail_level, test_count):
    combined_string = f"{requirements}{tables}{detail_level}{test_count}"
    return hashlib.md5(combined_string.encode()).hexdigest()

if "outputs_cache" not in st.session_state:
    st.session_state.outputs_cache = {}
if "test_scenarios_cache" not in st.session_state:
    st.session_state.test_scenarios_cache = {}

# Generate BRD
if st.button("Generate BRD") and uploaded_files:
    if not hasattr(st.session_state, 'extracted_data') or not st.session_state.extracted_data['requirements']:
        st.error("No content extracted from documents. Please check your uploaded files.")
    else:
        progress_placeholder = st.empty()
        status_text = st.empty()
        
        # Generate the main BRD
        status_text.write("Generating BRD... This may take a minute.")
        progress_bar = progress_placeholder.progress(0)
        
        llm_chain = initialize_llm()
        
        prompt_input = {
            "requirements": st.session_state.extracted_data['requirements'],
            "tables": st.session_state.extracted_data['tables'],
        }
        
        doc_hash = generate_hash(
            st.session_state.extracted_data['requirements'], 
            st.session_state.extracted_data['tables'],
            detail_level,
            test_scenario_count
        )
        
        if doc_hash in st.session_state.outputs_cache:
            output = st.session_state.outputs_cache[doc_hash]
            progress_bar.progress(50)
            status_text.write("Retrieved BRD from cache! Generating test scenarios...")
        else:
            output = llm_chain.run(prompt_input)
            st.session_state.outputs_cache[doc_hash] = output
            progress_bar.progress(50)
            status_text.write("BRD generated! Now enhancing test scenarios...")
        
        # Generate enhanced test scenarios if selected
        if generate_separate_test_scenarios:
            test_scenario_generator = initialize_test_scenario_generator()
            
            if doc_hash in st.session_state.test_scenarios_cache:
                test_scenarios = st.session_state.test_scenarios_cache[doc_hash]
                progress_bar.progress(100)
                status_text.write("Test scenarios retrieved from cache!")
            else:
                test_scenarios = test_scenario_generator.run({"brd_content": output})
                st.session_state.test_scenarios_cache[doc_hash] = test_scenarios
                progress_bar.progress(100)
                status_text.write("Test scenarios generated!")
            
            # Replace test scenarios section in the BRD
            if "# 7.0 Test Scenarios" in output:
                parts = output.split("# 7.0 Test Scenarios")
                section_end = parts[1].find("\n# ")
                if section_end == -1:
                    section_end = len(parts[1])
                
                # Replace the content of section 7
                enhanced_output = parts[0] + "# 7.0 Test Scenarios\n\n" + test_scenarios
                
                # Add the rest of the document back if there are sections after 7
                if section_end < len(parts[1]):
                    enhanced_output += parts[1][section_end:]
                
                output = enhanced_output
        
        progress_placeholder.empty()
        status_text.empty()
        st.success("BRD generated successfully!")
        
        # Display the output
        st.subheader("Generated Business Requirements Document")
        
        # Create tabs for different sections
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["Complete BRD"])
        
        with tab1:
            st.markdown(output)
        
        # Extract sections for the other tabs
        intro_section = ""
        if "# 1.0 Introduction" in output:
            intro_start = output.find("# 1.0 Introduction")
            intro_end = output.find("# 2.0 Impact Analysis")
            if intro_end > intro_start:
                intro_section = output[intro_start:intro_end]
        
        analysis_section = ""
        if "# 2.0 Impact Analysis" in output:
            analysis_start = output.find("# 2.0 Impact Analysis")
            analysis_end = output.find("# 3.0 Process")
            if analysis_end > analysis_start:
                analysis_section = output[analysis_start:analysis_end]
        
        requirements_section = ""
        if "# 4.0 Business / System Requirement" in output:
            req_start = output.find("# 4.0 Business / System Requirement")
            req_end = output.find("# 5.0 MIS / DATA Requirement")
            if req_end > req_start:
                requirements_section = output[req_start:req_end]
        
        test_section = ""
        if "# 7.0 Test Scenarios" in output:
            test_start = output.find("# 7.0 Test Scenarios")
            test_end = output.find("# 8.0 Questions")
            if test_end > test_start:
                test_section = output[test_start:test_end]
            else:
                test_section = output[test_start:]
        
        with tab2:
            st.markdown(intro_section)
            st.markdown(analysis_section)
        
        with tab3:
            st.markdown(requirements_section)
        
        with tab4:
            st.markdown(test_section)
        
        with tab5:
            # Display other sections like References, Appendix, etc.
            st.markdown(output[output.find("# 8.0 Questions"):])
        
        # Create a Word document
        doc = Document()
        
        # Add title
        doc.add_heading('Business Requirements Document', level=0)
        
        # Process the markdown output into Word
        sections = output.split('\n#')
        
        # Add the first part (before any headings)
        if not sections[0].startswith('#'):
            doc.add_paragraph(sections[0].strip())
            sections.pop(0)
        
        # Process each section
        for section in sections:
            if not section.strip():
                continue
            
            # Split into lines
            lines = section.strip().split('\n')
            
            # The first line should be the heading
            heading_text = lines[0].lstrip('#').strip()
            
            # Determine heading level by counting leading '#' characters
            heading_level = 1
            if section.startswith('#'):
                heading_level = len(section) - len(section.lstrip('#'))
            
            # Add the heading
            doc.add_heading(heading_text, level=heading_level)
            
            # Process the rest of the content in the section
            content = '\n'.join(lines[1:]).strip()
            if content:
                doc.add_paragraph(content)
        
        # Save the document to a buffer
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # Provide download button
        st.download_button(
            label="Download BRD as Word document",
            data=buffer,
            file_name="Business_Requirements_Document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
