import streamlit as st
from langchain_openai import ChatOpenAI
from langchain_groq import ChatGroq
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
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

@st.cache_resource
def initialize_llm(api_provider, api_key):
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
    
    llm_chain = LLMChain(
        llm=model, 
        prompt=PromptTemplate(
            input_variables=['requirements', 'brd_format'],
            template="""
            You are a Business Analyst expert creating a comprehensive Business Requirements Document (BRD). 
            
            DOCUMENT STRUCTURE TO FOLLOW:
            {brd_format}

            SOURCE REQUIREMENTS:
            {requirements}

            DETAILED INSTRUCTIONS:

            **1.0 Introduction**
            - 1.1 Purpose: Extract the business purpose, objectives, goals, or problem statement
            - 1.2 To be process / High level solution: Look for solution overview, high-level approach, or process descriptions

            **2.0 Impact Analysis**
            - 2.1 System impacts: Identify affected systems, integrations, dependencies, upstream/downstream impacts
            - 2.2 Impacted Products: List specific products, services, or business lines affected
            - 2.3 List of APIs required: Extract API names, endpoints, integrations, web services, or technical interfaces

            **3.0 Process / Data Flow diagram / Figma**
            - Look for: Process flows, workflow descriptions, data movement, user journeys, or references to diagrams
            - Include: Step-by-step processes, decision points, data transformations

            **4.0 Business / System Requirement**
            - Functional requirements (what the system should do)
            - Business rules and logic
            - User stories or use cases
            - Performance, security, and compliance requirements

            **5.0 MIS / DATA Requirement**
            - Data requirements and specifications
            - Reporting needs, analytics requirements
            - Data sources and destinations

            **6.0 Communication Requirement**
            - Stakeholder communication needs
            - Notification requirements
            - Email templates or communication workflows

            **7.0 Test Scenarios**
            - Generate at least 5 detailed test scenarios in a table format
            - Include: Test ID, Test Name, Objective, Test Steps, Expected Results, Test Data, Type

            **8.0 Questions / Suggestions**
            - Open questions from the source documents
            - Assumptions that need validation
            - Suggestions for improvement

            **9.0 Reference Document**
            - Source documents mentioned
            - Related policies or procedures
            - External references or standards

            **10.0 Appendix**
            - Supporting information
            - Detailed technical specifications

            **11.0 Risk Evaluation**
            - Identified risks and mitigation strategies
            - Dependencies and constraints
            - Timeline and technical risks

            OUTPUT REQUIREMENTS:
            - Use markdown formatting (## for main sections, ### for subsections)
            - Include comprehensive content for each section
            - Maintain professional business document tone
            - If information is not available for a section, state "Not applicable based on provided requirements"

            Generate the complete BRD now:"""
        )
    )
    return llm_chain

def extract_content_from_docx(doc_file):
    doc = Document(doc_file)
    content = []
    
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            content.append(paragraph.text.strip())
    
    # Extract tables
    for table in doc.tables:
        table_content = []
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells]
            table_content.append(" | ".join(row_text))
        content.append("TABLE:\n" + "\n".join(table_content))
    
    return "\n".join(content)

def extract_content_from_pdf(pdf_file):
    content = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            if page.extract_text():
                content.append(page.extract_text())
            
            # Extract tables
            tables = page.extract_tables()
            for table in tables:
                table_text = []
                for row in table:
                    row_text = [str(cell) if cell else "" for cell in row]
                    table_text.append(" | ".join(row_text))
                content.append("TABLE:\n" + "\n".join(table_text))
    
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
                table_lines = []
                table_lines.append(" | ".join(df.columns.tolist()))
                for _, row in df.head(10).iterrows():  # First 10 rows
                    table_lines.append(" | ".join([str(val) for val in row]))
                content.append("TABLE:\n" + "\n".join(table_lines))
    
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
    
    # Process BRD content
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
                
                # Process content
                for line in lines[1:]:
                    line = line.strip()
                    if line:
                        if line.startswith('- ') or line.startswith('* '):
                            doc.add_paragraph(line[2:].strip(), style='List Bullet')
                        elif re.match(r'^\d+\.', line):
                            doc.add_paragraph(re.sub(r'^\d+\.\s*', '', line), style='List Number')
                        else:
                            doc.add_paragraph(line)
    
    return doc

# Streamlit UI
st.title("Business Requirements Document Generator")

st.subheader("AI Model Selection")
api_provider = st.radio("Select API Provider:", ["OpenAI", "Groq"])

if api_provider == "OpenAI":
    api_key = st.text_input("Enter your OpenAI API Key:", type="password")
else:
    api_key = st.text_input("Enter your Groq API Key:", type="password")

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
        st.write(f"Generating BRD using {api_provider} API...")
        
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
            
            # Generate BRD
            llm_chain = initialize_llm(api_provider, api_key)
            
            prompt_input = {
                "requirements": "\n\n".join(combined_requirements),
                "brd_format": BRD_FORMAT
            }
            
            output = llm_chain.run(prompt_input)
            
            st.success("BRD generated successfully!")
            
            # Display generated BRD
            st.subheader("Generated Business Requirements Document")
            st.markdown(output)
            
            # Create Word document
            logo_data = logo_file.getvalue() if logo_file else None
            doc = create_word_document(output, logo_data)
            
            # Save to BytesIO
            doc_buffer = BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            
            # Provide download
            st.download_button(
                label="Download BRD as Word Document",
                data=doc_buffer.getvalue(),
                file_name="Business_Requirements_Document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Error generating BRD: {str(e)}")
            st.error("Please check your API key and try again.")
