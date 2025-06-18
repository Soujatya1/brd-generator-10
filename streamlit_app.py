import streamlit as st
from langchain_openai import ChatOpenAI
from langchain_groq import ChatGroq
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain, SimpleSequentialChain
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from io import BytesIO
import os
import pdfplumber
import pandas as pd
import extract_msg
import re

BRD_FORMAT = """
## 1.0 Introduction
    ## 1.1 Purpose
    ## 1.2 To be process / High level solution
## 2.0 Impact Analysis
    ## 2.1 System impacts – Primary and cross functional
    ## 2.2 Impacted Products
    ## 2.3 List of APIs required
## 3.0 Process / Data Flow diagram / Figma
## 4.0 Business / System Requirement
## 5.0 MIS / DATA Requirement
## 6.0 Communication Requirement
## 7.0 Test Scenarios
## 8.0 Questions / Suggestions
## 9.0 Reference Document
## 10.0 Appendix
## 11.0 Risk Evaluation
"""

# Section-wise templates for sequential processing
SECTION_TEMPLATES = {
    "intro_impact": """
    You are a Business Analyst expert creating sections 1.0-2.0 of a comprehensive Business Requirements Document (BRD).
    
    SOURCE REQUIREMENTS:
    {requirements}
    
    Create ONLY the following sections with detailed content:
    
    ## 1.0 Introduction
    ### 1.1 Purpose
    Extract and elaborate on the business purpose, objectives, goals, or problem statement from the requirements.
    
    ### 1.2 To be process / High level solution
    Provide solution overview, high-level approach, or process descriptions based on the requirements.
    
    ## 2.0 Impact Analysis
    ### 2.1 System impacts – Primary and cross functional
    Identify affected systems, integrations, dependencies, upstream/downstream impacts.
    
    ### 2.2 Impacted Products
    List specific products, services, or business lines affected.
    
    ### 2.3 List of APIs required
    Extract API names, endpoints, integrations, web services, or technical interfaces.
    
    IMPORTANT: 
    - Use markdown formatting (## for main sections, ### for subsections)
    - If tables are present in requirements, preserve them using markdown table format
    - Include comprehensive content for each section
    - If information is not available, state "Not applicable based on provided requirements"
    """,
    
    "process_requirements": """
    You are a Business Analyst expert creating sections 3.0-4.0 of a comprehensive Business Requirements Document (BRD).
    
    Previous sections context: {previous_content}
    
    SOURCE REQUIREMENTS:
    {requirements}
    
    Create ONLY the following sections with detailed content:
    
    ## 3.0 Process / Data Flow diagram / Figma
    Look for process flows, workflow descriptions, data movement, user journeys, or references to diagrams.
    Include step-by-step processes, decision points, data transformations.
    
    ## 4.0 Business / System Requirement
    - Functional requirements (what the system should do)
    - Business rules and logic
    - User stories or use cases
    - Performance, security, and compliance requirements
    - Include any requirement tables from source documents here
    
    IMPORTANT:
    - Use markdown formatting (## for main sections, ### for subsections)
    - Preserve any tables using markdown table format with pipes (|)
    - Include comprehensive content for each section
    - Build upon the context from previous sections
    """,
    
    "data_communication": """
    You are a Business Analyst expert creating sections 5.0-6.0 of a comprehensive Business Requirements Document (BRD).
    
    Previous sections context: {previous_content}
    
    SOURCE REQUIREMENTS:
    {requirements}
    
    Create ONLY the following sections with detailed content:
    
    ## 5.0 MIS / DATA Requirement
    - Data requirements and specifications
    - Reporting needs, analytics requirements
    - Data sources and destinations
    - Include any data specification tables from source documents here
    
    ## 6.0 Communication Requirement
    - Stakeholder communication needs
    - Notification requirements
    - Email templates or communication workflows
    
    IMPORTANT:
    - Use markdown formatting (## for main sections, ### for subsections)
    - Preserve any tables using markdown table format with pipes (|)
    - Include comprehensive content for each section
    - Build upon the context from previous sections
    """,
    
    "testing_final": """
    You are a Business Analyst expert creating sections 7.0-11.0 of a comprehensive Business Requirements Document (BRD).
    
    Previous sections context: {previous_content}
    
    SOURCE REQUIREMENTS:
    {requirements}
    
    Create ONLY the following sections with detailed content:
    
    ## 7.0 Test Scenarios
    Generate at least 5 detailed test scenarios in a table format:
    | Test ID | Test Name | Objective | Test Steps | Expected Results | Test Data | Type |
    | ------- | --------- | --------- | ---------- | ---------------- | --------- | ---- |
    | TC001 | [Test Name] | [Objective] | [Steps] | [Results] | [Data] | [Type] |
    
    ## 8.0 Questions / Suggestions
    - Open questions from the source documents
    - Assumptions that need validation
    - Suggestions for improvement
    
    ## 9.0 Reference Document
    - Source documents mentioned
    - Related policies or procedures
    - External references or standards
    
    ## 10.0 Appendix
    - Supporting information
    - Detailed technical specifications
    - Include any supporting tables from source documents here
    
    ## 11.0 Risk Evaluation
    - Identified risks and mitigation strategies
    - Dependencies and constraints
    - Timeline and technical risks
    
    IMPORTANT:
    - Use markdown formatting (## for main sections, ### for subsections)
    - Preserve any tables using markdown table format with pipes (|)
    - Include comprehensive content for each section
    - Build upon the context from previous sections
    """
}

def estimate_content_size(text):
    """Rough estimation of content size (characters as proxy for tokens)"""
    return len(text)

def chunk_requirements(requirements, max_chunk_size=8000):
    """Split requirements into smaller chunks if too large"""
    if estimate_content_size(requirements) <= max_chunk_size:
        return [requirements]
    
    # Split by sections or paragraphs
    sections = requirements.split('\n\n')
    chunks = []
    current_chunk = ""
    
    for section in sections:
        if estimate_content_size(current_chunk + section) > max_chunk_size and current_chunk:
            chunks.append(current_chunk.strip())
            current_chunk = section
        else:
            current_chunk += "\n\n" + section if current_chunk else section
    
    if current_chunk:
        chunks.append(current_chunk.strip())
    
    return chunks

@st.cache_resource
def initialize_sequential_chains(api_provider, api_key):
    """Initialize the sequential chain for BRD generation"""
    
    if api_provider == "OpenAI":
        model = ChatOpenAI(
            openai_api_key=api_key,
            model_name="gpt-3.5-turbo-16k",
            temperature=0.2,
            top_p=0.2
        )
    else:
        model = ChatGroq(
            groq_api_key=api_key,
            model_name="llama3-70b-8192",
            temperature=0.2,
            top_p=0.2
        )
    
    # Create individual chains for each section group
    chains = []
    
    # Chain 1: Introduction and Impact Analysis
    chain1 = LLMChain(
        llm=model,
        prompt=PromptTemplate(
            input_variables=['requirements'],
            template=SECTION_TEMPLATES["intro_impact"]
        ),
        output_key="intro_impact_sections"
    )
    
    # Chain 2: Process and Business Requirements
    chain2 = LLMChain(
        llm=model,
        prompt=PromptTemplate(
            input_variables=['previous_content', 'requirements'],
            template=SECTION_TEMPLATES["process_requirements"]
        ),
        output_key="process_requirements_sections"
    )
    
    # Chain 3: Data and Communication Requirements
    chain3 = LLMChain(
        llm=model,
        prompt=PromptTemplate(
            input_variables=['previous_content', 'requirements'],
            template=SECTION_TEMPLATES["data_communication"]
        ),
        output_key="data_communication_sections"
    )
    
    # Chain 4: Testing and Final sections
    chain4 = LLMChain(
        llm=model,
        prompt=PromptTemplate(
            input_variables=['previous_content', 'requirements'],
            template=SECTION_TEMPLATES["testing_final"]
        ),
        output_key="testing_final_sections"
    )
    
    return [chain1, chain2, chain3, chain4]

def generate_brd_sequentially(chains, requirements):
    """Generate BRD using sequential chains"""
    
    # Check if content is too large and needs chunking
    req_chunks = chunk_requirements(requirements)
    
    if len(req_chunks) > 1:
        st.info(f"Large content detected. Processing in {len(req_chunks)} chunks...")
    
    all_sections = []
    
    for chunk_idx, req_chunk in enumerate(req_chunks):
        st.write(f"Processing chunk {chunk_idx + 1}/{len(req_chunks)}...")
        
        # Initialize variables for sequential processing
        previous_content = ""
        chunk_sections = []
        
        # Process each chain sequentially
        for i, chain in enumerate(chains):
            try:
                if i == 0:  # First chain doesn't need previous content
                    result = chain.run(requirements=req_chunk)
                else:
                    result = chain.run(previous_content=previous_content, requirements=req_chunk)
                
                chunk_sections.append(result)
                previous_content += "\n\n" + result
                
                st.write(f"✓ Completed section group {i+1}/4")
                
            except Exception as e:
                st.error(f"Error in chain {i+1}: {str(e)}")
                # Continue with next chain even if one fails
                chunk_sections.append(f"## Error in section group {i+1}\nError processing this section: {str(e)}")
        
        all_sections.extend(chunk_sections)
    
    # Combine all sections
    final_brd = "\n\n".join(all_sections)
    return final_brd

def create_fallback_chain(api_provider, api_key):
    """Fallback to original single chain if sequential fails"""
    
    if api_provider == "OpenAI":
        model = ChatOpenAI(
            openai_api_key=api_key,
            model_name="gpt-3.5-turbo-16k",
            temperature=0.2,
            top_p=0.2
        )
    else:
        model = ChatGroq(
            groq_api_key=api_key,
            model_name="llama3-70b-8192",
            temperature=0.2,
            top_p=0.2
        )
    
    return LLMChain(
        llm=model, 
        prompt=PromptTemplate(
            input_variables=['requirements', 'brd_format'],
            template="""
            You are a Business Analyst expert creating a comprehensive Business Requirements Document (BRD). 
            
            DOCUMENT STRUCTURE TO FOLLOW:
            {brd_format}

            SOURCE REQUIREMENTS:
            {requirements}

            CRITICAL INSTRUCTIONS FOR TABLE HANDLING:
            - When you encounter "TABLE:" sections in the requirements, PRESERVE them in the BRD output
            - Format all tables using markdown table syntax with pipes (|)
            - Include ALL table data from the source requirements
            
            Generate a complete BRD following the structure provided. Use markdown formatting and include comprehensive content for each section.
            """
        )
    )

# [Include all the existing helper functions: add_hyperlink, add_bookmark, create_clickable_toc, etc.]
# [These remain the same as in your original code]

def add_hyperlink(paragraph, text, url_or_bookmark, is_internal=True):
    """Add a hyperlink to a paragraph"""
    hyperlink = OxmlElement('w:hyperlink')
    
    if is_internal:
        hyperlink.set(qn('w:anchor'), url_or_bookmark)
    else:
        part = paragraph.part
        r_id = part.relate_to(url_or_bookmark, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        hyperlink.set(qn('r:id'), r_id)
    
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    # Add color and underline for hyperlink style
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')
    rPr.append(color)
    
    underline = OxmlElement('w:u')
    underline.set(qn('w:val'), 'single')
    rPr.append(underline)
    
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    
    paragraph._p.append(hyperlink)
    return hyperlink

def add_bookmark(paragraph, bookmark_name):
    """Add a bookmark to a paragraph"""
    bookmark_start = OxmlElement('w:bookmarkStart')
    bookmark_start.set(qn('w:id'), str(abs(hash(bookmark_name)) % 1000000))
    bookmark_start.set(qn('w:name'), bookmark_name)
    
    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), str(abs(hash(bookmark_name)) % 1000000))
    
    paragraph._p.insert(0, bookmark_start)
    paragraph._p.append(bookmark_end)

def create_clickable_toc(doc):
    """Create a clickable table of contents with page numbers"""
    toc_heading = doc.add_heading('Table of Contents', level=1)
    add_bookmark(toc_heading, 'TOC')
    
    # TOC entries with their bookmark names
    toc_entries = [
        ("1.0 Introduction", "introduction"),
        ("    1.1 Purpose", "purpose"),
        ("    1.2 To be process / High level solution", "process_solution"),
        ("2.0 Impact Analysis", "impact_analysis"),
        ("    2.1 System impacts – Primary and cross functional", "system_impacts"),
        ("    2.2 Impacted Products", "impacted_products"), 
        ("    2.3 List of APIs required", "apis_required"),
        ("3.0 Process / Data Flow diagram / Figma", "process_flow"),
        ("4.0 Business / System Requirement", "business_requirements"),
        ("5.0 MIS / DATA Requirement", "mis_data_requirement"),
        ("6.0 Communication Requirement", "communication_requirement"),
        ("7.0 Test Scenarios", "test_scenarios"),
        ("8.0 Questions / Suggestions", "questions_suggestions"),
        ("9.0 Reference Document", "reference_document"),
        ("10.0 Appendix", "appendix"),
        ("11.0 Risk Evaluation", "risk_evaluation")
    ]
    
    for entry_text, bookmark_name in toc_entries:
        toc_paragraph = doc.add_paragraph()
        
        # Add the entry text as hyperlink
        if entry_text.startswith("    "):
            toc_paragraph.add_run("    ")  # Add indentation
            link_text = entry_text.strip()
        else:
            link_text = entry_text
            
        add_hyperlink(toc_paragraph, link_text, bookmark_name, is_internal=True)
        
        # Add dots/leaders
        toc_paragraph.add_run(" " + "." * 50 + " ")
        
        # Add page number field
        page_run = toc_paragraph.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        
        instrText = OxmlElement('w:instrText')
        instrText.text = f'PAGEREF {bookmark_name} \\h'
        
        fldChar2 = OxmlElement('w:fldChar') 
        fldChar2.set(qn('w:fldCharType'), 'end')
        
        page_run._r.append(fldChar1)
        page_run._r.append(instrText)
        page_run._r.append(fldChar2)
    
    # Add note about updating TOC
    note_para = doc.add_paragraph()
    note_para.add_run("Note: ").bold = True
    note_para.add_run("Right-click on this Table of Contents and select 'Update Field' to refresh page numbers after opening in Microsoft Word.")
    
    return {entry[1]: entry[0] for entry in toc_entries}

def add_section_with_bookmark(doc, heading_text, bookmark_name, level=1):
    """Add a section heading with bookmark for TOC linking"""
    heading = doc.add_heading(heading_text, level=level)
    add_bookmark(heading, bookmark_name)
    return heading

def create_table_in_doc(doc, table_data):
    """Create a table in the Word document from table data"""
    if not table_data or len(table_data) < 2:
        return None
    
    # Create table with proper dimensions
    table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
    table.style = 'Table Grid'
    
    # Add header row styling
    for i, cell_text in enumerate(table_data[0]):
        cell = table.rows[0].cells[i]
        cell.text = str(cell_text) if cell_text else ""
        # Make header bold
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
    
    # Add data rows
    for row_idx, row_data in enumerate(table_data[1:], 1):
        for col_idx, cell_text in enumerate(row_data):
            if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                table.rows[row_idx].cells[col_idx].text = str(cell_text) if cell_text else ""
    
    return table

def parse_markdown_table(table_text):
    """Parse markdown table format into structured data"""
    lines = [line.strip() for line in table_text.split('\n') if line.strip()]
    
    if len(lines) < 2:
        return None
    
    # Remove separator line (usually second line with | --- | --- |)
    if len(lines) >= 2 and '---' in lines[1]:
        lines.pop(1)
    
    table_data = []
    for line in lines:
        if line.startswith('|') and line.endswith('|'):
            # Remove leading and trailing |
            line = line[1:-1]
            cells = [cell.strip() for cell in line.split('|')]
            table_data.append(cells)
        else:
            # Handle lines without proper markdown table format
            cells = [cell.strip() for cell in line.split('|')]
            table_data.append(cells)
    
    return table_data if table_data else None

# [Include all other existing helper functions: extract_content_from_docx, extract_content_from_pdf, etc.]
# [These remain the same as in your original code]

def extract_content_from_docx(doc_file):
    doc = Document(doc_file)
    content = []
    
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            content.append(paragraph.text.strip())
    
    # Extract tables with better formatting
    for table in doc.tables:
        table_content = []
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells]
            table_content.append(" | ".join(row_text))
        if table_content:
            content.append("TABLE:")
            content.append("\n".join(table_content))
    
    return "\n".join(content)

def extract_content_from_pdf(pdf_file):
    content = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            if page.extract_text():
                content.append(page.extract_text())
            
            # Extract tables with better formatting
            tables = page.extract_tables()
            for table in tables:
                if table:
                    table_text = []
                    for row in table:
                        row_text = [str(cell) if cell else "" for cell in row]
                        table_text.append(" | ".join(row_text))
                    if table_text:
                        content.append("TABLE:")
                        content.append("\n".join(table_text))
    
    return "\n".join(content)

def extract_content_from_excel(excel_file):
    content = []
    try:
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        for sheet_name, df in excel_data.items():
            if not df.empty:
                content.append(f"Excel Sheet: {sheet_name}")
                content.append(f"Dimensions: {df.shape[0]} rows × {df.shape[1]} columns")
                content.append(f"Columns: {', '.join(df.columns.tolist())}")
                
                # Add table representation
                content.append("TABLE:")
                table_lines = []
                table_lines.append(" | ".join(df.columns.tolist()))
                for _, row in df.iterrows():
                    table_lines.append(" | ".join([str(val) for val in row]))
                content.append("\n".join(table_lines))
    
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
    
    return "\n".join(content)

def extract_content_from_msg(msg_file):
    try:
        temp_file = BytesIO(msg_file.getvalue())
        temp_file.name = msg_file.name
        
        msg = extract_msg.Message(temp_file)
        body_content = msg.body
        
        # Clean email content
        cleaned_body = re.sub(r'^(From|To|Cc|Subject|Sent|Date):.*?\n', '', body_content, flags=re.MULTILINE)
        cleaned_body = re.sub(r'_{10,}[\s\S]*$', '', cleaned_body)
        cleaned_body = re.sub(r'-{10,}[\s\S]*$', '', cleaned_body)
        
        # Remove disclaimers
        disclaimer_pattern = r'DISCLAIMER:[\s\S]*?customercare@bajajallianz\.co\.in'
        cleaned_body = re.sub(disclaimer_pattern, '', cleaned_body)
        
        return cleaned_body.strip()
    
    except Exception as e:
        st.error(f"Error processing MSG file: {str(e)}")
        return ""

def add_header_with_logo(doc, logo_bytes):
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    
    run = header_para.add_run()
    logo_stream = BytesIO(logo_bytes)
    run.add_picture(logo_stream, width=Inches(1.5))

def create_word_document(content, logo_data=None):
    doc = Document()
    
    # Add logo to header if provided
    if logo_data:
        add_header_with_logo(doc, logo_data)
    
    # Add title page
    for _ in range(12):
        doc.add_paragraph()
    
    title = doc.add_heading('Business Requirements Document', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    
    # Add Version History
    doc.add_heading('Version History', level=1)
    version_table = doc.add_table(rows=5, cols=5)
    version_table.style = 'Table Grid'
    hdr_cells = version_table.rows[0].cells
    headers = ['Version', 'Date', 'Author', 'Change description', 'Review by']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    doc.add_paragraph('**To be reviewed and filled in by IT Team.**')
    
    # Add Sign-off Matrix
    doc.add_heading('Sign-off Matrix', level=1)
    signoff_table = doc.add_table(rows=5, cols=5)
    signoff_table.style = 'Table Grid'
    hdr_cells = signoff_table.rows[0].cells
    headers = ['Version', 'Sign-off Authority', 'Business Function', 'Sign-off Date', 'Email Confirmation']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    doc.add_page_break()
    
    # Create clickable table of contents
    bookmark_mapping = create_clickable_toc(doc)
    
    doc.add_page_break()
    
    # Process BRD content with table support
    sections = content.split('##')
    
    for section in sections:
        if section.strip():
            lines = section.strip().split('\n')
            if lines:
                # Extract heading
                heading_line = lines[0].strip()
                
                # Find the appropriate bookmark name for this heading
                bookmark_name = None
                for bookmark, heading_text in bookmark_mapping.items():
                    if heading_line.lower().replace('#', '').strip() in heading_text.lower():
                        bookmark_name = bookmark
                        break
                
                if heading_line.startswith('###'):
                    level = 2
                    heading_text = heading_line.replace('###', '').strip()
                else:
                    level = 1
                    heading_text = heading_line.replace('##', '').strip()
                
                # Add section heading with bookmark
                if bookmark_name:
                    add_section_with_bookmark(doc, heading_text, bookmark_name, level)
                else:
                    doc.add_heading(heading_text, level)
                
                # Process content with table detection
                i = 1
                while i < len(lines):
                    line = lines[i].strip()
                    
                    # Check if this line starts a markdown table
                    if line and '|' in line and line.count('|') >= 2:
                        # Collect all table lines
                        table_lines = []
                        while i < len(lines) and lines[i].strip() and '|' in lines[i]:
                            table_lines.append(lines[i].strip())
                            i += 1
                        
                        # Parse and create table
                        if table_lines:
                            table_data = parse_markdown_table('\n'.join(table_lines))
                            if table_data:
                                create_table_in_doc(doc, table_data)
                        continue
                    
                    # Regular content processing
                    if line:
                        if line.startswith('- ') or line.startswith('* '):
                            doc.add_paragraph(line[2:].strip(), style='List Bullet')
                        elif re.match(r'^\d+\.', line):
                            doc.add_paragraph(re.sub(r'^\d+\.\s*', '', line), style='List Number')
                        else:
                            doc.add_paragraph(line)
                    
                    i += 1
    
    return doc

# Enhanced Streamlit UI
st.title("Business Requirements Document Generator")
st.subheader("🔄 Enhanced with Sequential Chain Processing")

st.info("💡 This version uses sequential chain processing to handle large documents and avoid token limits!")

st.subheader("AI Model Selection")
api_provider = st.radio("Select API Provider:", ["OpenAI", "Groq"])

if api_provider == "OpenAI":
    api_key = st.text_input("Enter your OpenAI API Key:", type="password")
else:
    api_key = st.text_input("Enter your Groq API Key:", type="password")

# Processing method selection
st.subheader("Processing Method")
processing_method = st.radio(
    "Choose processing method:",
    ["Sequential Chain (Recommended)", "Single Chain (Fallback)"],
    help="Sequential Chain breaks down the BRD generation into smaller, manageable parts to avoid token limits."
)

st.subheader("Document Logo")
logo_file = st.file_uploader("Upload logo/icon for document (PNG):", type=['png'])

if logo_file:
    st.image(logo_file, caption="Logo Preview", width=100)
    st.success("Logo uploaded successfully!")

st.subheader("Requirement Documents")
uploaded_files = st.file_uploader("Upload requirement documents:", 
                                 accept_multiple_files=True, 
                                 type=['pdf', 'docx', 'xlsx', 'msg'])

if st.button("Generate BRD") and uploaded_files:
    if not api_key:
        st.error(f"Please enter your {api_provider} API Key.")
    else:
        st.write(f"Generating BRD using {api_provider} API with {processing_method}...")
        
        try:
            # Extract content from all uploaded files
            combined_requirements = []
            
            for uploaded_file in uploaded_files:
                file_extension = os.path.splitext(uploaded_file.name)[-1].lower()
                st.write(f"Processing {uploaded_file.name}...")
                
                if file_extension == ".docx":
                    content = extract_content_from_docx(uploaded_file)
                elif file_extension == ".pdf":
                    content = extract_content_from_pdf(uploaded_file)
                elif file_extension == ".xlsx":
                    content = extract_content_from_excel(uploaded_file)
                elif file_extension == ".msg":
                    content = extract_content_from_msg(uploaded_file)
                else:
                    st.warning(f"Unsupported file format: {uploaded_file.name}")
                    continue
                
                combined_requirements.append(content)
            
            all_requirements = "\n\n".join(combined_requirements)
            
            # Show content size estimation
            content_size = estimate_content_size(all_requirements)
            st.info(f"📊 Content size: ~{content_size:,} characters")
            
            # Generate BRD based on selected method
            if processing_method == "Sequential Chain (Recommended)":
                try:
                    chains = initialize_sequential_chains(api_provider, api_key)
                    output = generate_brd_sequentially(chains, all_requirements)
                    st.success("✅ BRD generated successfully using Sequential Chain!")
                    
                except Exception as e:
                    st.warning(f"Sequential chain failed: {str(e)}")
                    st.info("🔄 Falling back to Single Chain method...")
                    
                    # Fallback to single chain
                    fallback_chain = create_fallback_chain(api_provider, api_key)
                    
                    # Try chunking if content is too large
                    if content_size > 10000:
                        chunks = chunk_requirements(all_requirements, max_chunk_size=8000)
                        chunk_outputs = []
                        
                        for i, chunk in enumerate(chunks):
                            st.write(f"Processing fallback chunk {i+1}/{len(chunks)}...")
                            chunk_output = fallback_chain.run({
                                "requirements": chunk,
                                "brd_format": BRD_FORMAT
                            })
                            chunk_outputs.append(chunk_output)
                        
                        output = "\n\n".join(chunk_outputs)
                    else:
                        output = fallback_chain.run({
                            "requirements": all_requirements,
                            "brd_format": BRD_FORMAT
                        })
                    
                    st.success("✅ BRD generated successfully using Fallback Single Chain!")
            
            else:  # Single Chain method
                try:
                    single_chain = create_fallback_chain(api_provider, api_key)
                    
                    # Check if content needs chunking
                    if content_size > 10000:
                        st.info("🔄 Large content detected. Processing in chunks...")
                        chunks = chunk_requirements(all_requirements, max_chunk_size=8000)
                        chunk_outputs = []
                        
                        for i, chunk in enumerate(chunks):
                            st.write(f"Processing chunk {i+1}/{len(chunks)}...")
                            chunk_output = single_chain.run({
                                "requirements": chunk,
                                "brd_format": BRD_FORMAT
                            })
                            chunk_outputs.append(chunk_output)
                        
                        output = "\n\n".join(chunk_outputs)
                    else:
                        output = single_chain.run({
                            "requirements": all_requirements,
                            "brd_format": BRD_FORMAT
                        })
                    
                    st.success("✅ BRD generated successfully using Single Chain!")
                    
                except Exception as e:
                    st.error(f"Single chain processing failed: {str(e)}")
                    st.error("Please try the Sequential Chain method or check your API key.")
                    st.stop()
            
            # Display generated BRD
            st.subheader("Generated Business Requirements Document")
            
            # Add expandable sections for better readability
            with st.expander("📋 View Complete BRD Content", expanded=True):
                st.markdown(output)
            
            # Show BRD statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📄 Total Sections", len([s for s in output.split('##') if s.strip()]))
            with col2:
                st.metric("📝 Total Characters", len(output))
            with col3:
                st.metric("📊 Total Words", len(output.split()))
            
            # Create Word document
            st.info("🔄 Creating Word document...")
            logo_data = logo_file.getvalue() if logo_file else None
            doc = create_word_document(output, logo_data)
            
            # Save to BytesIO
            doc_buffer = BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            
            # Provide download
            st.success("✅ Word document created successfully!")
            st.download_button(
                label="📥 Download BRD as Word Document",
                data=doc_buffer.getvalue(),
                file_name="Business_Requirements_Document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # Additional features
            st.subheader("📋 Additional Features")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🔍 Analyze BRD Quality"):
                    # Simple BRD quality analysis
                    sections_found = []
                    required_sections = [
                        "Introduction", "Impact Analysis", "Process", "Business", 
                        "Requirement", "Test", "Risk", "Reference"
                    ]
                    
                    for section in required_sections:
                        if section.lower() in output.lower():
                            sections_found.append(section)
                    
                    st.write("**BRD Quality Analysis:**")
                    st.write(f"✅ Sections covered: {len(sections_found)}/{len(required_sections)}")
                    st.write(f"📝 Content completeness: {min(100, len(output) // 100)}%")
                    st.write(f"📊 Tables detected: {output.count('|')//3}")
                    
                    if len(sections_found) >= 6:
                        st.success("🎉 High quality BRD generated!")
                    elif len(sections_found) >= 4:
                        st.warning("⚠️ Good BRD, but could use more sections")
                    else:
                        st.error("❌ BRD needs improvement")
            
            with col2:
                if st.button("📊 Export BRD Summary"):
                    # Create a summary of the BRD
                    summary_lines = []
                    sections = output.split('##')
                    
                    for section in sections:
                        if section.strip():
                            lines = section.strip().split('\n')
                            if lines:
                                heading = lines[0].strip()
                                word_count = len(section.split())
                                summary_lines.append(f"• {heading}: {word_count} words")
                    
                    summary = "**BRD Summary:**\n\n" + "\n".join(summary_lines)
                    
                    st.text_area("BRD Summary", summary, height=200)
                    
                    # Download summary as text file
                    st.download_button(
                        label="📥 Download Summary",
                        data=summary,
                        file_name="BRD_Summary.txt",
                        mime="text/plain"
                    )
            
        except Exception as e:
            st.error(f"❌ Error generating BRD: {str(e)}")
            st.error("Please check your API key and try again.")
            
            # Show debug information
            with st.expander("🔧 Debug Information"):
                st.write(f"**API Provider:** {api_provider}")
                st.write(f"**Processing Method:** {processing_method}")
                st.write(f"**Content Size:** {content_size if 'content_size' in locals() else 'Unknown'}")
                st.write(f"**Files Processed:** {len(uploaded_files)}")
                st.write(f"**Error Details:** {str(e)}")

# Add helpful information in sidebar
with st.sidebar:
    st.markdown("### 📚 Help & Information")
    
    st.markdown("""
    **Sequential Chain Benefits:**
    - ✅ Handles large documents
    - ✅ Avoids token limits
    - ✅ Better content organization
    - ✅ More reliable processing
    
    **Supported File Types:**
    - 📄 PDF files
    - 📝 Word documents (.docx)
    - 📊 Excel files (.xlsx)
    - 📧 Outlook messages (.msg)
    
    **Tips for Best Results:**
    - Use clear, structured requirement documents
    - Include tables and diagrams when possible
    - Provide complete business context
    - Review generated BRD before finalizing
    """)
    
    st.markdown("### 🔧 Troubleshooting")
    st.markdown("""
    **Common Issues:**
    - **API Key Error:** Verify your API key is correct
    - **Large Files:** Try Sequential Chain method
    - **Missing Sections:** Check source document quality
    - **Token Limits:** Use chunking or Sequential Chain
    """)
    
    st.markdown("### 📞 Support")
    st.info("For technical support, please check your API provider's documentation.")

# Footer
st.markdown("---")
st.markdown(
    "**Business Requirements Document Generator** | "
    "Enhanced with Sequential Chain Processing | "
    "Built with Streamlit & LangChain"
)
