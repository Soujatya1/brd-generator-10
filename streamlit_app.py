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
            1.0 Introduction [Page 4]
                1.1 Purpose [Page 4]
                1.2 To be process / High level solution [Page 4]
            2.0 Impact Analysis [Page 4]
                2.1 System impacts – Primary and cross functional [Page 4]
                2.2 Impacted Products [Page 5]
                2.3 List of APIs required [Page 5]
            3.0 Process / Data Flow diagram / Figma [Page 6]
            4.0 Business / System Requirement [Page 7]
                4.1 Application / Module Name [Page 7]
                4.2 Application / Module Name [Page 7]
            5.0 MIS / DATA Requirement [Page 8]
            6.0 Communication Requirement [Page 8]
            7.0 Test Scenarios [Page 8]
            8.0 Questions / Suggestions [Page 8]
            9.0 Reference Document [Page 9]
            10.0 Appendix [Page 9]

            Requirements:
            Analyze the content provided in the requirement documents and map the relevant information to each section defined in the BRD structure. Be concise, specific, and maintain professional language.

            Tables:
            If applicable, include the following tabular information extracted from the documents:
            {tables}

            Formatting Guidelines:
            1. Use proper heading levels (# for main sections, ## for subsections)
            2. Include bullet points or numbered lists for better readability
            3. Clearly differentiate between functional and non-functional requirements
            4. Maintain consistent formatting throughout the document
            5. Keep page number references accurate
            6. Include tables where information is best presented in tabular format

            Requirements Processing:
            1. Categorize extracted information based on the BRD structure
            2. Prioritize clarity and organization over verbosity
            3. Include explicit section headings with page numbers as shown above
            4. Ensure all content from input documents is represented in the appropriate sections

            Input Document Content:
            {requirements}

            Output:
            Generate a complete, well-structured Business Requirements Document following the exact structure provided, with appropriate page numbers. The document should be comprehensive yet concise, professionally formatted, and ready for business use.
            """
        )
    )
    
    return llm_chain

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

# UI Components
st.title("Enhanced BRD Generator")
st.write("Upload requirement documents to generate a professionally formatted Business Requirements Document.")

with st.expander("About this tool", expanded=False):
    st.write("""
    This tool analyzes your uploaded requirement documents and generates a well-structured Business Requirements Document (BRD) 
    following a standard format with proper page numbering. The generated document can be downloaded as a Word file.
    
    The BRD will follow this structure:
    ```
    {}
    ```
    """.format(BRD_FORMAT))

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
            
            # Show extraction success
            if combined_requirements:
                st.success("Documents successfully processed")
    
    # Preview extracted data
    with st.expander("Preview Extracted Content", expanded=False):
        st.subheader("Extracted Text")
        st.text_area("Text Content", st.session_state.extracted_data['requirements'], height=200)
        
        st.subheader("Extracted Tables")
        st.text_area("Table Content", st.session_state.extracted_data['tables'], height=200)

# Caching mechanism
def generate_hash(requirements, tables):
    combined_string = requirements + tables
    return hashlib.md5(combined_string.encode()).hexdigest()

if "outputs_cache" not in st.session_state:
    st.session_state.outputs_cache = {}

# Generate BRD
if st.button("Generate BRD") and uploaded_files:
    if not hasattr(st.session_state, 'extracted_data') or not st.session_state.extracted_data['requirements']:
        st.error("No content extracted from documents. Please check your uploaded files.")
    else:
        with st.spinner("Generating BRD... This may take a minute."):
            llm_chain = initialize_llm()
            
            prompt_input = {
                "requirements": st.session_state.extracted_data['requirements'],
                "tables": st.session_state.extracted_data['tables'],
            }
            
            doc_hash = generate_hash(
                st.session_state.extracted_data['requirements'], 
                st.session_state.extracted_data['tables']
            )
            
            if doc_hash in st.session_state.outputs_cache:
                output = st.session_state.outputs_cache[doc_hash]
                st.success("Retrieved from cache!")
            else:
                output = llm_chain.run(prompt_input)
                st.session_state.outputs_cache[doc_hash] = output
                st.success("BRD generated successfully!")
            
            # Display the output
            st.subheader("Generated Business Requirements Document")
            st.markdown(output)
            
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

# Similarity comparison
st.write("---")
st.subheader("Optional: Upload a sample BRD for comparison")
sample_file = st.file_uploader("Upload a sample BRD (PDF/DOCX):", type=["pdf", "docx"], key="sample_file")

def calculate_text_similarity(text1, text2):
    vectorizer = TfidfVectorizer().fit_transform([text1, text2])
    vectors = vectorizer.toarray()
    cosine_sim = cosine_similarity(vectors)
    return cosine_sim[0][1] * 100

def calculate_structural_similarity(tables1, tables2):
    sm = difflib.SequenceMatcher(None, tables1, tables2)
    return sm.ratio() * 100

if sample_file and "outputs_cache" in st.session_state and st.session_state.outputs_cache:
    with st.spinner("Analyzing similarity..."):
        file_extension = os.path.splitext(sample_file.name)[-1].lower()
        
        if file_extension == ".docx":
            sample_text, sample_table
