import streamlit as st
from langchain_openai import ChatOpenAI
from langchain_groq import ChatGroq
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
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

@st.cache_resource
def initialize_llm(api_provider, api_key):
    if api_provider == "OpenAI":
        model = ChatOpenAI(
            openai_api_key=api_key,
            model_name="gpt-4.5-preview",
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

        Document Structure:
        {brd_format}

        Requirements:
        Analyze the content provided in the requirement documents and map the relevant information to each section defined in the BRD format according to the {requirements}. Be concise and specific.
        Wherever you find the similar header, please pick all the information from the {requirements} files and put it into the BRD as per the format.
        
        Tables:
        {tables}

        IMPORTANT: DO NOT try to recreate or reformat tables. Instead, when a table should be included, insert a marker [[TABLE_ID:identifier]] in the exact location where the table should appear.

        In the BRD format, section: 4.0 Business / System Requirement, should contain the business process flows from the {requirements} which should be in a table format as well as textual data.
        
        Formatting:
        1. Use headings and subheadings for clear organization.
        2. Include bullet points or numbered lists where necessary for better readability.
        3. Clearly differentiate between functional and non-functional requirements.
        4. For tables, include the table marker [[TABLE_ID:identifier]] as provided in the tables section.
        
        Key Points:
        1. Use the given format `{brd_format}` strictly as the base structure for the BRD.
        2. Ensure all relevant information from the requirements is displayed under the corresponding section.
        3. Avoid including irrelevant or speculative information.
        4. Summarize lengthy content while preserving its meaning.
        5. Do not attempt to recreate tables - use the table markers exactly as provided.

        Output:
        The output must be formatted cleanly as a Business Requirements Document, following professional standards. Avoid verbose language and stick to the structure defined above.
        """
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
    """Extract only the body content from Outlook MSG file and optionally save as TXT"""
    try:
        temp_file = BytesIO(msg_file.getvalue())
        temp_file.name = msg_file.name
        
        msg = extract_msg.Message(temp_file)
        
        # Extract only the body content
        body_content = msg.body
        
        # Save as text file if requested
        if save_as_txt:
            txt_filename = os.path.splitext(msg_file.name)[0] + ".txt"
            temp_txt_path = os.path.join("/tmp", txt_filename)
            with open(temp_txt_path, "w", encoding="utf-8") as txt_file:
                txt_file.write(body_content)
            
            # Return both the content and the filename
            return body_content, txt_filename
        
        return body_content
    except Exception as e:
        st.error(f"Error processing MSG file: {str(e)}")
        return ""

def insert_table_into_doc(doc, table_to_insert, table_id):
    if isinstance(table_to_insert, pd.DataFrame):
        df = table_to_insert
        rows, cols = df.shape
        
        word_table = doc.add_table(rows=rows+1, cols=cols)
        word_table.style = 'Table Grid'
        
        for col_idx, column_name in enumerate(df.columns):
            word_table.cell(0, col_idx).text = str(column_name)
        
        for row_idx, (_, row) in enumerate(df.iterrows()):
            for col_idx, cell_value in enumerate(row):
                word_table.cell(row_idx+1, col_idx).text = str(cell_value)
        
        return word_table
    else:
        new_table = doc.add_table(rows=len(table_to_insert.rows), cols=len(table_to_insert.rows[0].cells))
        new_table.style = 'Table Grid'
        
        for i, row in enumerate(table_to_insert.rows):
            for j, cell in enumerate(row.cells):
                new_table.cell(i, j).text = cell.text
        
        return new_table

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
        
        elif file_extension == ".msg":
            msg_content, txt_filename = extract_content_from_msg(uploaded_file, save_as_txt=True)
            if msg_content:
                combined_requirements.append(msg_content)
        
                # Store the content in session state with the filename as key
                st.session_state.msg_content[txt_filename] = msg_content
        
                # Show a success message
                st.success(f"Email extracted and saved as: {txt_filename}")
        
                # Display the content in a collapsible section
                with st.expander(f"View content of {txt_filename}"):
                    st.text_area("Email Body", msg_content, height=300, key=f"txt_{txt_filename}")
        
            else:
                st.warning(f"Unsupported file format: {uploaded_file.name}")
    
    st.session_state.extracted_data = {
        'requirements': "\n\n".join(combined_requirements),
        'tables': "\n\n".join(all_tables_as_text),
        'original_tables': all_original_tables
    }

def add_header_with_logo(doc, logo_bytes):
    section = doc.sections[0]
    
    header = section.header
    
    header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    
    run = header_para.add_run()
    logo_stream = BytesIO(logo_bytes)
    run.add_picture(logo_stream, width=Inches(1.5))

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
            
            output = llm_chain.run(prompt_input)
            
            test_scenario_generator = initialize_test_scenario_generator(api_provider, api_key)
            test_scenarios = test_scenario_generator.run({"brd_content": output})
            
            output = output.replace("7.0 Test Scenarios", "7.0 Test Scenarios\n" + test_scenarios)
            
            st.success("BRD generated successfully!")
            
            st.subheader("Generated Business Requirements Document")
            display_output = re.sub(r'\[\[TABLE_ID:[a-zA-Z0-9_]+\]\]', '[TABLE WILL BE INSERTED HERE]', output)
            st.markdown(display_output)
            
            doc = Document()
            doc.add_heading('Business Requirements Document', level=0)
            
            if logo_file:
                logo_bytes = logo_file.getvalue()
                add_header_with_logo(doc, logo_bytes)
            
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

            doc.add_paragraph('**Review by should be someone from IT function.**', style='Caption')

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

            doc.add_page_break()
            
            sections = output.split('\n#')
            
            for section in sections:
                if not section.strip():
                    continue
                
                lines = section.strip().split('\n')
                heading_text = lines[0].lstrip('#').strip()
                heading_level = 1 if section.startswith('#') else 2
                doc.add_heading(heading_text, level=heading_level)
                
                remaining_content = '\n'.join(lines[1:]).strip()
                
                table_pattern = r'\[\[TABLE_ID:([a-zA-Z0-9_]+)\]\]'
                matches = list(re.finditer(table_pattern, remaining_content))
                
                last_pos = 0
                for match in matches:
                    pre_text = remaining_content[last_pos:match.start()].strip()
                    if pre_text:
                        doc.add_paragraph(pre_text)
                    
                    table_id = match.group(1)
                    if table_id in st.session_state.extracted_data['original_tables']:
                        st.write(f"Inserting table {table_id}")
                        table_to_insert = st.session_state.extracted_data['original_tables'][table_id]
                        insert_table_into_doc(doc, table_to_insert, table_id)
                    else:
                        doc.add_paragraph(f"[TABLE {table_id} NOT FOUND]")
                    
                    last_pos = match.end()
                
                remaining_text = remaining_content[last_pos:].strip()
                if remaining_text:
                    
                    lines = remaining_text.split('\n')
                    clean_lines = []
                    skip_mode = False
                    
                    for line in lines:
                        if '|' in line and skip_mode:
                            continue
                        elif not line.strip() and skip_mode:
                            skip_mode = False
                        elif not skip_mode:
                            clean_lines.append(line)
                    
                    clean_text = '\n'.join(clean_lines)
                    if clean_text.strip():
                        doc.add_paragraph(clean_text)
            
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.download_button(
                label="Download BRD as Word document",
                data=buffer,
                file_name="Business_Requirements_Document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.info(f"This might be due to an invalid {api_provider} API key or connection issues. Please check your API key and try again.")
