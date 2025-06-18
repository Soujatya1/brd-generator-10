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
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_BREAK
from docx.enum.style import WD_STYLE_TYPE

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
    - Timeline and technical risks
    - If Risk Assessment tables/data are found in source requirements, PRESERVE the complete table structure using markdown format and add the same
    
    IMPORTANT:
    - Use markdown formatting (## for main sections, ### for subsections)
    - Preserve any tables using markdown table format with pipes (|)
    - Include comprehensive content for each section
    - If content found for the mentioned sections, put it in the BRD as-is, else build upon the context from previous sections
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
    
    # Process all chunks together to create ONE BRD
    combined_requirements = "\n\n=== DOCUMENT BREAK ===\n\n".join(req_chunks)
    
    # Initialize variables for sequential processing
    previous_content = ""
    final_sections = []
    
    # Process each chain sequentially with ALL content
    for i, chain in enumerate(chains):
        try:
            if i == 0:  # First chain doesn't need previous content
                result = chain.run(requirements=combined_requirements)
            else:
                result = chain.run(previous_content=previous_content, requirements=combined_requirements)
            
            final_sections.append(result)
            previous_content += "\n\n" + result
            
            st.write(f"✓ Completed section group {i+1}/4")
            
        except Exception as e:
            st.error(f"Error in chain {i+1}: {str(e)}")
            # Continue with next chain even if one fails
            final_sections.append(f"## Error in section group {i+1}\nError processing this section: {str(e)}")
    
    # Combine all sections into ONE final BRD
    final_brd = "\n\n".join(final_sections)
    return final_brd

def create_toc_styles(doc):
    """Create TOC styles if they don't exist"""
    styles = doc.styles
    
    # Create TOC 1 style if it doesn't exist
    try:
        toc1_style = styles['TOC 1']
    except KeyError:
        toc1_style = styles.add_style('TOC 1', WD_STYLE_TYPE.PARAGRAPH)
        toc1_style.font.name = 'Calibri'
        toc1_style.font.size = Pt(11)
        toc1_style.paragraph_format.left_indent = Inches(0)
        toc1_style.paragraph_format.space_after = Pt(0)
    
    # Create TOC 2 style if it doesn't exist
    try:
        toc2_style = styles['TOC 2']
    except KeyError:
        toc2_style = styles.add_style('TOC 2', WD_STYLE_TYPE.PARAGRAPH)
        toc2_style.font.name = 'Calibri'
        toc2_style.font.size = Pt(11)
        toc2_style.paragraph_format.left_indent = Inches(0.25)
        toc2_style.paragraph_format.space_after = Pt(0)

def create_clickable_toc(doc):
    """Create a clickable table of contents with page numbers"""
    toc_heading = doc.add_heading('Table of Contents', level=1)
    add_bookmark(toc_heading, 'TOC')
    
    # Create TOC styles
    create_toc_styles(doc)
    
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
        
        # Apply appropriate style
        try:
            if entry_text.startswith("    "):
                toc_paragraph.style = 'TOC 2'
            else:
                toc_paragraph.style = 'TOC 1'
        except KeyError:
            # Fallback: manually format if styles still don't work
            if entry_text.startswith("    "):
                toc_paragraph.paragraph_format.left_indent = Inches(0.25)
            toc_paragraph.paragraph_format.space_after = Pt(0)
        
        # Add the entry text as hyperlink
        if entry_text.startswith("    "):
            toc_paragraph.add_run("    ")  # Add indentation
            link_text = entry_text.strip()
        else:
            link_text = entry_text
            
        add_hyperlink(toc_paragraph, link_text, bookmark_name, is_internal=True)
        
        # Add tab stop for page numbers (important!)
        toc_paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(6.0))
        
        # Add tab character before page number
        toc_paragraph.add_run("\t")
        
        # Add page number field with proper XML structure
        page_run = toc_paragraph.add_run()
        
        # Create field XML elements
        fldChar_begin = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>')
        instrText = parse_xml(f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> PAGEREF {bookmark_name} \\h </w:instrText>')
        fldChar_end = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>')
        
        # Add field elements to the run
        page_run._r.append(fldChar_begin)
        page_run._r.append(instrText)
        page_run._r.append(fldChar_end)
        
        # Add placeholder text that will be replaced when field updates
        page_run.add_text(" ")  # Placeholder page number
    
    # Add comprehensive note about updating TOC
    doc.add_paragraph()  # Empty line
    note_para = doc.add_paragraph()
    note_run = note_para.add_run("IMPORTANT: ")
    note_run.bold = True
    note_run.font.color.rgb = RGBColor(255, 0, 0)  # Red color
    
    note_para.add_run("To see actual page numbers in this Table of Contents:")
    
    # Add numbered instructions
    instructions = [
        "Press Ctrl+A to select all, then F9 to update all fields in the document."
    ]

def add_hyperlink(paragraph, text, url_or_bookmark, is_internal=True):
    """Add a hyperlink to a paragraph with proper styling"""
    # Get the document part for relationship handling
    part = paragraph.part
    
    # Create hyperlink element
    hyperlink = OxmlElement('w:hyperlink')
    
    if is_internal:
        # Internal bookmark link
        hyperlink.set(qn('w:anchor'), url_or_bookmark)
    else:
        # External URL link
        r_id = part.relate_to(url_or_bookmark, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        hyperlink.set(qn('r:id'), r_id)
    
    # Create run element with hyperlink styling
    new_run = OxmlElement('w:r')
    
    # Create run properties for hyperlink style
    rPr = OxmlElement('w:rPr')
    
    # Set hyperlink color (blue)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')
    rPr.append(color)
    
    # Set underline
    underline = OxmlElement('w:u')
    underline.set(qn('w:val'), 'single')
    rPr.append(underline)
    
    # Add run properties to run
    new_run.append(rPr)
    
    # Add text to run
    text_element = OxmlElement('w:t')
    text_element.text = text
    new_run.append(text_element)
    
    # Add run to hyperlink
    hyperlink.append(new_run)
    
    # Add hyperlink to paragraph
    paragraph._p.append(hyperlink)
    
    return hyperlink

def add_bookmark(paragraph, bookmark_name):
    """Add a bookmark to a paragraph with unique ID"""
    # Generate unique bookmark ID
    bookmark_id = str(abs(hash(bookmark_name)) % 1000000)
    
    # Create bookmark start element
    bookmark_start = OxmlElement('w:bookmarkStart')
    bookmark_start.set(qn('w:id'), bookmark_id)
    bookmark_start.set(qn('w:name'), bookmark_name)
    
    # Create bookmark end element
    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), bookmark_id)
    
    # Insert bookmark at the beginning and end of paragraph
    paragraph._p.insert(0, bookmark_start)
    paragraph._p.append(bookmark_end)

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

def extract_content_from_excel(excel_file, max_rows_per_sheet=70, max_sample_rows=10):
    content = []
    try:
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        for sheet_name, df in excel_data.items():
            if df.empty:
                continue
                
            content.append(f"=== EXCEL SHEET: {sheet_name} ===")
            content.append(f"Total Dimensions: {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Add column information
            content.append(f"Columns ({len(df.columns)}): {', '.join(df.columns.tolist())}")
            
            # Add data type information
            data_types = df.dtypes.to_dict()
            type_summary = []
            for col, dtype in data_types.items():
                type_summary.append(f"{col}: {str(dtype)}")
            content.append(f"Data Types: {', '.join(type_summary[:10])}...")  # Limit to first 10
            
            # Add statistical summary for numeric columns
            numeric_cols = df.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                content.append(f"Numeric Columns: {', '.join(numeric_cols.tolist()[:5])}...")  # First 5 only
                
            # Add sample data (limited rows)
            sample_size = min(max_sample_rows, len(df))
            if sample_size > 0:
                content.append(f"\nSample Data (first {sample_size} rows):")
                content.append("TABLE:")
                
                # Create sample table with limited columns if too many
                display_df = df.head(sample_size)
                if len(df.columns) > 10:
                    # If too many columns, show first 8 and indicate there are more
                    display_cols = df.columns[:8].tolist() + [f"... +{len(df.columns)-8} more columns"]
                    display_df = df[df.columns[:8]].head(sample_size)
                    
                    # Add header row
                    header_row = " | ".join(display_cols)
                    content.append(header_row)
                else:
                    # Normal case - show all columns
                    header_row = " | ".join(df.columns.tolist())
                    content.append(header_row)
                
                # Add sample data rows
                for _, row in display_df.iterrows():
                    row_data = []
                    for val in row:
                        # Truncate long text values
                        str_val = str(val)
                        if len(str_val) > 50:
                            str_val = str_val[:47] + "..."
                        row_data.append(str_val)
                    content.append(" | ".join(row_data))
                
                # Add summary of remaining data if applicable
                if len(df) > sample_size:
                    content.append(f"... and {len(df) - sample_size} more rows")
            
            # Add key insights if possible
            content.append(f"\nData Summary:")
            
            # Check for key patterns or important columns
            key_columns = []
            for col in df.columns:
                col_lower = col.lower()
                if any(keyword in col_lower for keyword in ['id', 'name', 'title', 'status', 'type', 'category', 'priority', 'requirement']):
                    key_columns.append(col)
            
            if key_columns:
                content.append(f"Key Columns Identified: {', '.join(key_columns[:5])}")
                
                # Show unique values for key categorical columns
                for col in key_columns[:3]:  # Limit to first 3 key columns
                    if df[col].dtype == 'object':  # Text/categorical column
                        unique_vals = df[col].dropna().unique()
                        if len(unique_vals) <= 20:  # Only if manageable number of unique values
                            content.append(f"{col} Values: {', '.join(map(str, unique_vals[:10]))}")
                        else:
                            content.append(f"{col}: {len(unique_vals)} unique values")
            
            # Add missing data info
            missing_data = df.isnull().sum()
            if missing_data.sum() > 0:
                missing_cols = missing_data[missing_data > 0].head(5)  # Top 5 columns with missing data
                missing_info = [f"{col}: {count} missing" for col, count in missing_cols.items()]
                content.append(f"Missing Data: {', '.join(missing_info)}")
            
            content.append("="*50)  # Separator between sheets
    
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
        content.append(f"Error processing Excel file: {str(e)}")
    
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
st.subheader("🔄 Sequential Chain Processing")

st.info("💡 This version uses sequential chain processing to handle large documents and avoid token limits!")

st.subheader("AI Model Selection")
api_provider = st.radio("Select API Provider:", ["OpenAI", "Groq"])

if api_provider == "OpenAI":
    api_key = st.text_input("Enter your OpenAI API Key:", type="password")
else:
    api_key = st.text_input("Enter your Groq API Key:", type="password")

st.subheader("Document Logo")

logo_file = st.file_uploader("Upload Company Logo (optional):", type=['png', 'jpg', 'jpeg'])
logo_data = None
if logo_file:
    logo_data = logo_file.getvalue()
    st.success("✅ Logo uploaded successfully!")

st.subheader("Upload Requirements Documents")
uploaded_files = st.file_uploader(
    "Choose files", 
    type=['txt', 'docx', 'pdf', 'xlsx', 'xls', 'msg'],
    accept_multiple_files=True
)

# Text input option
st.subheader("Or Enter Requirements Manually")
manual_requirements = st.text_area(
    "Paste your requirements here:",
    height=200,
    placeholder="Enter your business requirements, user stories, or project specifications here..."
)

if st.button("🚀 Generate BRD", type="primary"):
    if not api_key:
        st.error("❌ Please enter your API key!")
    elif not uploaded_files and not manual_requirements.strip():
        st.error("❌ Please upload files or enter requirements manually!")
    else:
        try:
            # Initialize chains
            with st.spinner("🔧 Initializing AI chains..."):
                chains = initialize_sequential_chains(api_provider, api_key)
            
            # Extract content from uploaded files
            all_requirements = []
            
            if manual_requirements.strip():
                all_requirements.append("=== MANUAL REQUIREMENTS ===")
                all_requirements.append(manual_requirements.strip())
                all_requirements.append("="*50)
            
            if uploaded_files:
                st.info(f"📁 Processing {len(uploaded_files)} uploaded files...")
                
                for uploaded_file in uploaded_files:
                    file_extension = uploaded_file.name.split('.')[-1].lower()
                    
                    try:
                        st.write(f"📄 Processing: {uploaded_file.name}")
                        
                        if file_extension == 'txt':
                            content = str(uploaded_file.read(), "utf-8")
                        elif file_extension == 'docx':
                            content = extract_content_from_docx(uploaded_file)
                        elif file_extension == 'pdf':
                            content = extract_content_from_pdf(uploaded_file)
                        elif file_extension in ['xlsx', 'xls']:
                            content = extract_content_from_excel(uploaded_file)
                        elif file_extension == 'msg':
                            content = extract_content_from_msg(uploaded_file)
                        else:
                            st.warning(f"⚠️ Unsupported file type: {file_extension}")
                            continue
                        
                        if content.strip():
                            all_requirements.append(f"=== FILE: {uploaded_file.name} ===")
                            all_requirements.append(content.strip())
                            all_requirements.append("="*50)
                            st.success(f"✅ Successfully processed: {uploaded_file.name}")
                        else:
                            st.warning(f"⚠️ No content extracted from: {uploaded_file.name}")
                            
                    except Exception as e:
                        st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
                        continue
            
            if not all_requirements:
                st.error("❌ No valid content found in uploaded files!")
                st.stop()
            
            # Combine all requirements
            combined_requirements = "\n\n".join(all_requirements)
            
            # Display content size info
            content_size = estimate_content_size(combined_requirements)
            st.info(f"📊 Total content size: {content_size:,} characters")
            
            # Generate BRD using sequential chains
            st.subheader("🤖 AI Processing Progress")
            
            with st.spinner("🔄 Generating comprehensive BRD using sequential processing..."):
                brd_content = generate_brd_sequentially(chains, combined_requirements)
            
            if brd_content:
                st.success("✅ BRD generated successfully!")
                
                # Display generated content
                st.subheader("📋 Generated BRD Content")
                
                # Create expandable sections for preview
                with st.expander("🔍 Preview Generated BRD", expanded=False):
                    st.markdown(brd_content)
                
                # Generate Word document
                st.subheader("📄 Download Options")
                
                try:
                    with st.spinner("📝 Creating Word document..."):
                        doc = create_word_document(brd_content, logo_data)
                        
                        # Save to BytesIO
                        doc_buffer = BytesIO()
                        doc.save(doc_buffer)
                        doc_buffer.seek(0)
                        
                        # Download button for Word document
                        st.download_button(
                            label="📥 Download BRD (Word Document)",
                            data=doc_buffer.getvalue(),
                            file_name="Business_Requirements_Document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                        st.success("✅ Word document ready for download!")
                        
                except Exception as e:
                    st.error(f"❌ Error creating Word document: {str(e)}")
                    st.info("📋 You can still copy the content above manually.")
                
                # Also provide markdown download option
                try:
                    st.download_button(
                        label="📥 Download BRD (Markdown)",
                        data=brd_content,
                        file_name="Business_Requirements_Document.md",
                        mime="text/markdown"
                    )
                except Exception as e:
                    st.error(f"❌ Error creating markdown download: {str(e)}")
                
                # Content statistics
                st.subheader("📊 Document Statistics")
                word_count = len(brd_content.split())
                char_count = len(brd_content)
                section_count = brd_content.count('##')
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("📝 Word Count", f"{word_count:,}")
                with col2:
                    st.metric("🔤 Character Count", f"{char_count:,}")
                with col3:
                    st.metric("📑 Sections", section_count)
                
            else:
                st.error("❌ Failed to generate BRD content!")
                
        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
            st.info("💡 Try reducing the input size or check your API key.")
