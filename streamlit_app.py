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
import uuid
import re
import copy

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
    # Create the w:hyperlink tag and add needed values
    hyperlink = OxmlElement('w:hyperlink')
    
    if is_internal:
        # For internal bookmarks, just set the anchor attribute
        hyperlink.set(qn('w:anchor'), url_or_bookmark)
    else:
        # For external URLs, create a relationship
        part = paragraph.part
        r_id = part.relate_to(url_or_bookmark, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        hyperlink.set(qn('r:id'), r_id)
    
    # Create a new run object and add the text
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
    # Create bookmark start
    bookmark_start = OxmlElement('w:bookmarkStart')
    bookmark_start.set(qn('w:id'), str(abs(hash(bookmark_name)) % 1000000))
    bookmark_start.set(qn('w:name'), bookmark_name)
    
    # Create bookmark end
    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), str(abs(hash(bookmark_name)) % 1000000))
    
    # Insert bookmarks
    paragraph._p.insert(0, bookmark_start)
    paragraph._p.append(bookmark_end)

def add_page_field(paragraph):
    """Add a page number field to paragraph"""
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    
    run = paragraph.add_run()
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

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
        
        # Add dots/leaders (simplified version)
        toc_paragraph.add_run(" " + "." * 50 + " ")
        
        # Add page number (this will be updated when document is opened in Word)
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
            input_variables=['requirements', 'tables', 'brd_format'],
            template="""
            Create a Business Requirements Document (BRD) based on the following details:

        Document Structure To Follow:
        {brd_format}

        SOURCE REQUIREMENTS:
        {requirements}
        
        Tables:
        {tables}
           
        INSTRUCTIONS:
1. Create a BRD following the exact structure provided in the document format above
2. Map content from the source requirements to the appropriate BRD sections
3. If you find content that matches a BRD section header, include ALL relevant information from that section
4. Be comprehensive but concise - include all important details without unnecessary verbosity

TABLE HANDLING:
- When tables should be included, use the marker [[TABLE_ID:identifier]] exactly as provided in the tables section
- Do NOT recreate or reformat tables - only use the provided markers
- Place table markers in the most appropriate location within each section

SPECIFIC SECTION REQUIREMENTS:
- Section 4.0 (Business/System Requirements): Include business process flows, functional requirements, and any process-related tables
- Section 7.0 (Test Scenarios): Only include the placeholder "[[TEST_SCENARIOS_PLACEHOLDER]]" - this will be generated separately

OUTPUT FORMAT:
- Use proper markdown heading structure (## for main sections, ### for subsections)
- Include bullet points and numbered lists for clarity
- Maintain professional business document tone
- Ensure each section has relevant content or clearly state if information is not available

IMPORTANT: FOR ALL THE SECTIONS MENTIONED, PLEASE INCLUDE THE MOST ACCURATE INFORMATION AS PER THE UPLOADED DOCUMENTS, DO NOT MENTION GENERIC INFORMATION

Generate the complete BRD now:"""
        )
    )
    return llm_chain

@st.cache_resource
def initialize_test_scenario_generator(api_provider, api_key):
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
    
    test_scenario_chain = LLMChain(
        llm=model, 
        prompt=PromptTemplate(
            input_variables=['brd_content'],
            template="""
            Based on the following Business Requirements Document (BRD), generate detailed test scenarios for section 7.0 Test Scenarios:
            
            BRD Content:
            {brd_content}
            
            Special Instructions for Test Scenarios Section:
            Based on the entire BRD content, generate at least 5 detailed test scenarios in a tabular format that would comprehensively validate the requirements. For each test scenario:
            - Provide a clear test ID and descriptive name
            - Include test objective/purpose
            - List detailed test steps as per serial numbers
            - Define expected results/acceptance criteria
            - Specify test data requirements if applicable
            - Indicate whether it's a positive or negative test case
            - Note any dependencies or prerequisites
            
            Format your response EXACTLY as follows:
            1. First, create a paragraph or two introducing the test scenarios and their purpose.
            2. Then include EXACTLY this marker: [[TEST_SCENARIO_TABLE:start]]
            3. Format the test scenarios as a properly formatted markdown table with the following columns:
               | Test ID | Test Name | Objective | Test Steps | Expected Results | Test Data | Type |
               |---------|-----------|-----------|------------|-----------------|-----------|------|
               | TS-001  | ...       | ...       | ...        | ...             | ...       | ...  |
            4. End the table with EXACTLY this marker: [[TEST_SCENARIO_TABLE:end]]
            5. Add any concluding remarks after the table.
            
            IMPORTANT: Ensure the [[TEST_SCENARIO_TABLE:start]] and [[TEST_SCENARIO_TABLE:end]] markers appear exactly as shown, with no extra whitespace or characters.
            """
        )
    )
    return test_scenario_chain

def process_test_scenarios(output, doc):
    # Find test scenario table markers
    table_start_pattern = r'\[\[TEST_SCENARIO_TABLE:start\]\]'
    table_end_pattern = r'\[\[TEST_SCENARIO_TABLE:end\]\]'
    
    start_match = re.search(table_start_pattern, output)
    end_match = re.search(table_end_pattern, output)
    
    if start_match and end_match:
        # Extract the text before the table
        pre_text = output[:start_match.start()].strip()
        if pre_text:
            doc.add_paragraph(pre_text)
        
        # Extract the table content
        table_content = output[start_match.end():end_match.start()].strip()
        table_rows = [row.strip() for row in table_content.split('\n') if row.strip() and '|' in row]
        
        if len(table_rows) >= 2:  # Header row + at least one data row
            # Extract column count from header row
            header_row = table_rows[0]
            columns = [col.strip() for col in header_row.split('|') if col.strip()]
            col_count = len(columns)
            
            # Create the table in Word
            test_table = doc.add_table(rows=1, cols=col_count)  # Start with just the header row
            test_table.style = 'Table Grid'
            
            # Add header row
            header_cells = test_table.rows[0].cells
            for i, header in enumerate(columns):
                if i < len(header_cells):
                    header_cells[i].text = header
                    # Make header bold
                    for paragraph in header_cells[i].paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
            
            # Skip the separator row and add data rows
            data_rows = [row for row in table_rows[2:] if not all(c == '-' or c == '|' or c.isspace() for c in row)]
            for row in data_rows:
                cells = [cell.strip() for cell in row.split('|') if cell.strip()]
                if cells:  # Make sure we have data
                    row_cells = test_table.add_row().cells
                    for j, cell_content in enumerate(cells):
                        if j < col_count:
                            row_cells[j].text = cell_content
        else:
            # If we couldn't parse the table properly, just add it as text
            doc.add_paragraph(table_content)
        
        # Extract the text after the table
        post_text = output[end_match.end():].strip()
        if post_text:
            doc.add_paragraph(post_text)
    else:
        # If no proper table markers are found, just add the content as paragraphs
        st.warning("Test scenario table markers not found. Adding content as regular paragraphs.")
        doc.add_paragraph(output)
    
    return doc

def normalize_header(header):
    return header.lower().strip().replace('/', ' ').replace('  ', ' ')

def extract_content_from_docx(doc_file):
    doc = Document(doc_file)
    structured_content = []
    current_heading = "General"
    original_tables = {}
    
    for element in doc.element.body:
        if element.tag.endswith('p'):
            if len(structured_content) < len(doc.paragraphs):
                paragraph = doc.paragraphs[len(structured_content)]
                text = paragraph.text.strip()
                
                if text:
                    if paragraph.style.name.startswith('Heading'):
                        current_heading = text
                    
                    structured_content.append({
                        'type': 'paragraph',
                        'heading': current_heading,
                        'content': text
                    })
        
        elif element.tag.endswith('tbl'):
            table_index = len([e for e in structured_content if e['type'] == 'table'])
            if table_index < len(doc.tables):
                table = doc.tables[table_index]
                
                table_id = f"table_{uuid.uuid4().hex[:8]}"
                
                original_tables[table_id] = table
                
                table_content = []
                for row in table.rows:
                    row_text = [cell.text.strip() for cell in row.cells]
                    table_content.append(" | ".join(row_text))
                
                structured_content.append({
                    'type': 'table',
                    'heading': current_heading,
                    'content': f"[[TABLE_ID:{table_id}]]\n" + "\n".join(table_content),
                    'table_id': table_id
                })
    
    return structured_content, original_tables

def extract_tables_from_excel(excel_file):
    original_tables = {}
    table_markers = []
    
    try:
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        for sheet_name, df in excel_data.items():
            if not df.empty:
                table_id = f"excel_{sheet_name}_{uuid.uuid4().hex[:8]}".replace(' ', '_')
                df.columns = ["Inset Column Name" if str(col).startswith('Unnamed') else str(col)
                              for col in df.columns]
                
                original_tables[table_id] = df
                
                table_content = [f"Excel Sheet: {sheet_name} ({df.shape[0]} rows × {df.shape[1]} columns)"]
                table_content.append("| " + " | ".join(df.columns.tolist()) + " |")
                table_content.append("| " + " | ".join(["---"] * len(df.columns)) + " |")
                
                for _, row in df.head(5).iterrows():
                    table_content.append("| " + " | ".join([str(val) for val in row]) + " |")
                
                if df.shape[0] > 5:
                    table_content.append("| ... | " + " | ".join(["..."] * (len(df.columns)-1)) + " |")
                
                marker = f"[[TABLE_ID:{table_id}]]\n" + "\n".join(table_content)
                table_markers.append(marker)
    
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
    
    return original_tables, table_markers

def summarize_excel_data(excel_file):
    summaries = []
    
    try:
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        for sheet_name, df in excel_data.items():
            if not df.empty:
                summaries.append(f"Sheet '{sheet_name}':")
                summaries.append(f"- Dimensions: {df.shape[0]} rows × {df.shape[1]} columns")
                summaries.append(f"- Column names: {', '.join(df.columns.tolist())}")
                numeric_cols = df.select_dtypes(include=['number']).columns
                if not numeric_cols.empty:
                    summaries.append("- Numeric columns summary:")
                    for col in numeric_cols[:5]:
                        summaries.append(f"  {col}: min={df[col].min()}, max={df[col].max()}, avg={df[col].mean():.2f}")
                summaries.append("\n")
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
    
    return "\n".join(summaries)

def extract_content_from_msg(msg_file, save_as_txt=True):
    try:
        temp_file = BytesIO(msg_file.getvalue())
        temp_file.name = msg_file.name
        
        msg = extract_msg.Message(temp_file)
        
        body_content = msg.body
        
        cleaned_body = body_content
        
        cleaned_body = re.sub(r'^From:.*?\n', '', cleaned_body, flags=re.MULTILINE)
        
        cleaned_body = re.sub(r'^To:.*?\n', '', cleaned_body, flags=re.MULTILINE)
        
        cleaned_body = re.sub(r'^Cc:.*?\n', '', cleaned_body, flags=re.MULTILINE)
        
        cleaned_body = re.sub(r'^Subject:.*?\n', '', cleaned_body, flags=re.MULTILINE)
        
        cleaned_body = re.sub(r'^(Sent|Date):.*?\n', '', cleaned_body, flags=re.MULTILINE)
        
        cleaned_body = re.sub(r'_{10,}[\s\S]*$', '', cleaned_body)
        cleaned_body = re.sub(r'-{10,}[\s\S]*$', '', cleaned_body)
        
        balic_disclaimer = "DISCLAIMER: This email communication/message, including any attachments, may contain proprietary"
        if balic_disclaimer in cleaned_body:
            cleaned_body = cleaned_body.split(balic_disclaimer)[0]
        
        disclaimer_pattern = r'DISCLAIMER: This email communication/message[\s\S]*?customercare@bajajallianz\.co\.in'
        cleaned_body = re.sub(disclaimer_pattern, '', cleaned_body)
        
        cleaned_body = cleaned_body.strip()
        
        if save_as_txt:
            txt_filename = os.path.splitext(msg_file.name)[0] + ".txt"
            temp_txt_path = os.path.join("/tmp", txt_filename)
            with open(temp_txt_path, "w", encoding="utf-8") as txt_file:
                txt_file.write(cleaned_body)
            
            return cleaned_body, txt_filename
        
        return cleaned_body
    except Exception as e:
        st.error(f"Error processing MSG file: {str(e)}")
        return ""

def insert_table_into_doc(doc, table_to_insert, table_id, max_rows=50):
    if isinstance(table_to_insert, pd.DataFrame):
        df = table_to_insert
        
        if df.shape[0] > max_rows:
            df = df.head(max_rows)
        
        rows, cols = df.shape
        
        word_table = doc.add_table(rows=rows+1, cols=cols)
        word_table.style = 'Table Grid'
        
        for col_idx, column_name in enumerate(df.columns):
            if str(column_name).startswith('Unnamed'):
                column_display_name = "Insert Column Name"
            else:
                column_display_name = str(column_name)
            word_table.cell(0, col_idx).text = str(column_name)
        
        for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
            for col_idx, cell_value in enumerate(row):
                display_value = '-' if pd.isna(cell_value) else str(cell_value)
                word_table.cell(row_idx, col_idx).text = display_value
        
        return word_table
    else:
        new_table = doc.add_table(rows=len(table_to_insert.rows), cols=len(table_to_insert.rows[0].cells))
        new_table.style = 'Table Grid'
        
        for i, row in enumerate(table_to_insert.rows):
            for j, cell in enumerate(row.cells):
                cell_text = cell.text.replace('nan', '-') if cell.text == 'nan' else cell.text
                new_table.cell(i, j).text = cell_text
        
        return new_table

def add_header_with_logo(doc, logo_bytes):
    section = doc.sections[0]
    
    header = section.header
    
    header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    
    run = header_para.add_run()
    logo_stream = BytesIO(logo_bytes)
    run.add_picture(logo_stream, width=Inches(1.5))

st.title("Business Requirements Document Generator")

st.subheader("AI Model Selection")
api_provider = st.radio("Select API Provider:", ["OpenAI", "Groq"])

if api_provider == "OpenAI":
    api_key = st.text_input("Enter your OpenAI API Key:", help="Your API key will not be stored and is only used for this session")
else:
    api_key = st.text_input("Enter your Groq API Key:", help="Your API key will not be stored and is only used for this session")

st.subheader("Document Logo")
logo_file = st.file_uploader("Upload logo/icon for document (PNG):", type=['png'])

if logo_file is not None:
    st.image(logo_file, caption="Logo Preview", width=100)
    st.success("Logo uploaded successfully! It will be added to the document header.")
    if 'logo_data' not in st.session_state:
        st.session_state.logo_data = logo_file.getvalue()
else:
    st.info("Please upload a PNG logo/icon that will appear in the document header.")

st.subheader("Requirement Documents")
uploaded_files = st.file_uploader("Upload requirement documents (PDF/DOCX/XLSX/MSG):", 
                                 accept_multiple_files=True, 
                                 type=['pdf', 'docx', 'xlsx', 'msg'])

if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = {'requirements': '', 'tables': '', 'original_tables': {}}

if uploaded_files:
    combined_requirements = []
    all_tables_as_text = []
    all_original_tables = {}
    
    for uploaded_file in uploaded_files:
        file_extension = os.path.splitext(uploaded_file.name)[-1].lower()
        st.write(f"Processing {uploaded_file.name}...")
        
        if file_extension == ".docx":
            structured_content, original_tables = extract_content_from_docx(uploaded_file)
            
            all_original_tables.update(original_tables)
    
            organized_content = {}
            for item in structured_content:
                heading = item['heading']
                if heading not in organized_content:
                    organized_content[heading] = {'paragraphs': [], 'tables': []}
        
                if item['type'] == 'paragraph':
                    organized_content[heading]['paragraphs'].append(item['content'])
                else:
                    organized_content[heading]['tables'].append(item['content'])
    
            for heading, content in organized_content.items():
                section_text = [heading]
                section_text.extend(content['paragraphs'])
                combined_requirements.append("\n".join(section_text))
        
                if content['tables']:
                    table_text = [f"Tables for section {heading}:"]
                    table_text.extend(content['tables'])
                    all_tables_as_text.append("\n".join(table_text))
        
        elif file_extension == ".pdf":
            with pdfplumber.open(uploaded_file) as pdf:
                text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
                combined_requirements.append(text)
                
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        table_text = []
                        for row in table:
                            row_text = [str(cell) if cell else "" for cell in row]
                            table_text.append(" | ".join(row_text))
                        all_tables_as_text.append("\n".join(table_text))
        
        elif file_extension == ".xlsx":
            excel_tables, table_markers = extract_tables_from_excel(uploaded_file)
            
            all_original_tables.update(excel_tables)
            all_tables_as_text.extend(table_markers)
            
            excel_summary = summarize_excel_data(uploaded_file)
            combined_requirements.append(f"Excel file content from {uploaded_file.name}:\n{excel_summary}")
        if 'msg_content' not in st.session_state:
            st.session_state.msg_content = {}
        elif file_extension == ".msg":
            msg_content, txt_filename = extract_content_from_msg(uploaded_file, save_as_txt=True)
            if msg_content:
                combined_requirements.append(msg_content)
        
                st.session_state.msg_content[txt_filename] = msg_content
        
                st.success(f"Email body extracted and saved as: {txt_filename}")
        
                with st.expander(f"View content of {txt_filename}"):
                    st.text_area("Email Body", msg_content, height=300, key=f"txt_{txt_filename}")
        
            else:
                st.warning(f"Unsupported file format: {uploaded_file.name}")
    
    st.session_state.extracted_data = {
        'requirements': "\n\n".join(combined_requirements),
        'tables': "\n\n".join(all_tables_as_text),
        'original_tables': all_original_tables
    }

if st.button("Generate BRD") and uploaded_files:
    if not api_key:
        st.error(f"Please enter your {api_provider} API Key.")
    elif not st.session_state.extracted_data['requirements']:
        st.error("No content extracted from documents.")
    else:
        st.write(f"Generating BRD using {api_provider} API...")
        try:
            llm_chain = initialize_llm(api_provider, api_key)
            
            prompt_input = {
                "requirements": st.session_state.extracted_data['requirements'],
                "tables": st.session_state.extracted_data['tables'],
                "brd_format": BRD_FORMAT
            }
            
            # Generate the main BRD content
            output = llm_chain.run(prompt_input)
            
            # Generate test scenarios as a separate step
            test_scenario_generator = initialize_test_scenario_generator(api_provider, api_key)
            test_scenarios = test_scenario_generator.run({"brd_content": output})
            
            # Use an easily identifiable placeholder
            test_scenario_placeholder = "[[TEST_SCENARIOS_PLACEHOLDER]]"
            final_output = re.sub(r'[•\-\*]\s*\[\[TEST_SCENARIOS_PLACEHOLDER\]\]', "[[TEST_SCENARIOS_PLACEHOLDER]]", output)
            final_output = re.sub(r'\s*\[\[TEST_SCENARIOS_PLACEHOLDER\]\]\s*', "[[TEST_SCENARIOS_PLACEHOLDER]]", final_output)
            
            st.success("BRD generated successfully!")
            
            st.subheader("Generated Business Requirements Document")
            display_output = re.sub(r'\[\[TABLE_ID:[a-zA-Z0-9_]+\]\]', '[TABLE WILL BE INSERTED HERE]', final_output)
            display_output = display_output.replace(test_scenario_placeholder, "[TEST SCENARIOS WILL BE INSERTED HERE]")
            st.markdown(display_output)
            
            # For debugging - show the test scenarios content
            with st.expander("View Generated Test Scenarios"):
                st.text(test_scenarios)
            
            # Create Word document
            doc = Document()
            
            # Add spacing before title
            for _ in range(12):
                doc.add_paragraph()
            
            # Add logo to header if provided
            if logo_file is not None:
                add_header_with_logo(doc, st.session_state.logo_data)
            
            # Add title
            title = doc.add_heading('Business Requirements Document', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add page break
            doc.add_page_break()
            
            # Add Version History section
            doc.add_heading('Version History', level=1)
            version_table = doc.add_table(rows=1, cols=5)
            version_table.style = 'Table Grid'
            hdr_cells = version_table.rows[0].cells
            hdr_cells[0].text = 'Version'
            hdr_cells[1].text = 'Date'
            hdr_cells[2].text = 'Author'
            hdr_cells[3].text = 'Change description'
            hdr_cells[4].text = 'Review by'

            for _ in range(4):
                version_table.add_row()

            doc.add_paragraph('**To be reviewed and filled in by IT Team.**', style='Caption')

            # Add Sign-off Matrix section
            doc.add_heading('Sign-off Matrix', level=1)
            signoff_table = doc.add_table(rows=1, cols=5)
            signoff_table.style = 'Table Grid'
            hdr_cells = signoff_table.rows[0].cells
            hdr_cells[0].text = 'Version'
            hdr_cells[1].text = 'Sign-off Authority'
            hdr_cells[2].text = 'Business Function'
            hdr_cells[3].text = 'Sign-off Date'
            hdr_cells[4].text = 'Email Confirmation'

            for _ in range(4):
                signoff_table.add_row()

            doc.add_page_break()

            # Add Table of Contents
            doc.add_heading('Table of Contents', level=1)

            toc_paragraph = doc.add_paragraph()
            toc_paragraph.bold = True

            toc_entries = [
                "1.0 Introduction",
                "    1.1 Purpose",
                "    1.2 To be process / High level solution",
                "2.0 Impact Analysis",
                "    2.1 System impacts – Primary and cross functional",
                "    2.2 Impacted Products",
                "    2.3 List of APIs required",
                "3.0 Process / Data Flow diagram / Figma",
                "4.0 Business / System Requirement",
                "5.0 MIS / DATA Requirement",
                "6.0 Communication Requirement",
                "7.0 Test Scenarios",
                "8.0 Questions / Suggestions",
                "9.0 Reference Document",
                "10.0 Appendix",
                "11.0 Risk Evaluation"
            ]

            for entry in toc_entries:
                if entry.startswith("    "):
                    doc.add_paragraph(entry.strip(), style='Heading 3')
                else:
                    doc.add_paragraph(entry, style='Heading 2')

            # Create clickable table of contents
            bookmark_mapping = create_clickable_toc(doc)
            
            # Add page break after TOC
            doc.add_page_break()
            
            # Process each section of the BRD
            sections = final_output.split('##')
            
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
                        
                        # Determine heading level
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
                        content_lines = lines[1:]
                        i = 0

                        if "7.0 Test Scenarios" in heading_text or heading_text.startswith("7.0"):
                            doc = process_test_scenarios(test_scenarios, doc)
                            continue
                        
                        while i < len(content_lines):
                            line = content_lines[i].strip()
                            
                            # Handle table markers
                            table_match = re.search(r'\[\[TABLE_ID:([a-zA-Z0-9_]+)\]\]', line)
                            if table_match:
                                table_id = table_match.group(1)
                                if table_id in st.session_state.extracted_data['original_tables']:
                                    insert_table_into_doc(doc, st.session_state.extracted_data['original_tables'][table_id], table_id)
                                i += 1
                                continue
                            
                            
                            # Handle regular content
                            if line:
                                # Handle bullet points
                                if line.startswith('- ') or line.startswith('* '):
                                    bullet_content = line[2:].strip()
                                    para = doc.add_paragraph(bullet_content, style='List Bullet')
                                
                                # Handle numbered lists
                                elif re.match(r'^\d+\.', line):
                                    numbered_content = re.sub(r'^\d+\.\s*', '', line)
                                    para = doc.add_paragraph(numbered_content, style='List Number')
                                
                                # Handle regular paragraphs
                                else:
                                    para = doc.add_paragraph(line)
                            
                            i += 1
            
            # Save document to BytesIO
            doc_buffer = BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            
            # Provide download button
            st.download_button(
                label="Download BRD as Word Document",
                data=doc_buffer.getvalue(),
                file_name="Business_Requirements_Document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # Store generated content in session state for potential regeneration
            st.session_state.generated_brd = final_output
            st.session_state.generated_test_scenarios = test_scenarios
            
            # Optional: Display test scenarios separately
            with st.expander("View Generated Test Scenarios"):
                st.markdown(test_scenarios)
                
        except Exception as e:
            st.error(f"Error generating BRD: {str(e)}")
            st.error("Please check your API key and try again.")
        
        if st.session_state.extracted_data['tables']:
            st.subheader("Extracted Tables")
            st.text_area("Tables", st.session_state.extracted_data['tables'], height=200, key="tables_review")
