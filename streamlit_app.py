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
from langchain_openai import AzureChatOpenAI
from openpyxl import load_workbook

BRD_FORMAT = """
## 1.0 Introduction
    ## 1.1 Purpose
    ## 1.2 As-is process
    ## 1.3 To be process / High level solution
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

SECTION_TEMPLATES = {

    "intro_impact": """

You are a Business Analyst expert creating sections 1.0–2.0 of a comprehensive Business Requirements Document (BRD).

IMPORTANT: Do not output any ``` code fences or Mermaid syntax.
All text should be plain markdown (headings, lists, tables) only - no code blocks or fenced content.

SOURCE REQUIREMENTS:

{requirements}

CRITICAL INSTRUCTIONS:

- Extract information ONLY from the provided source requirements
- Identify the ACTUAL business problem being solved from the source
- Focus on what is explicitly mentioned, not assumptions
- Do NOT create, assume, or fabricate any content not present in the source
- If a section has no relevant information in the source, leave it BLANK
- Adapt to any domain (training, payments, integration, access control, etc.)

Create ONLY the following sections with detailed content in markdown:

## 1.0 Introduction

### 1.1 Purpose

Extract the EXACT business purpose from the source, focusing on:
- What is the main business objective or problem being addressed?
- What specific functionality or capability is being implemented?
- What restrictions, validations, or controls are being introduced?
- What business processes are being improved or changed?
- What compliance, security, or operational requirements are being met?

Look for key phrases that indicate purpose: "purpose", "objective", "requirement", "need", "problem", "solution", "implement", "restrict", "validate", "improve", "ensure"

### 1.2 As-is process

Extract the CURRENT state/process from the source:
- How does the current system/process work?
- What are the existing workflows or user journeys?
- What problems or limitations exist in the current approach?
- What manual processes or workarounds are currently used?
- What system behaviors need to be changed?
- Any screenshots, process flows, or current state descriptions

Look for indicators: "currently", "as-is", "existing", "present", "manual", "workaround", "problem with current", "limitations"

### 1.3 To be process / High level solution

Extract the PROPOSED solution from the source:
- What is the new process or system behavior?
- What workflow steps or validation logic will be implemented?
- How will the new solution address current problems?
- What automated processes will replace manual ones?
- What new capabilities or features will be added?
- Any conditional logic, decision trees, or multi-step processes

Look for indicators: "to-be", "proposed", "solution", "new process", "will be", "should be", "automated", "enhanced", "improved", "step-by-step", "workflow", "condition", "if-then"

## 2.0 Impact Analysis

### 2.1 System impacts – Primary and cross functional

Extract SPECIFIC system impacts mentioned in the source:
- Which primary systems/applications will be modified?
- What cross-functional or integrated systems are affected?
- What databases, APIs, or services need changes?
- What new integration points are required?
- Any performance, security, or infrastructure impacts
- Dependencies on other systems or services

### 2.2 Impacted Products

List ONLY the products/services explicitly mentioned in the source:
- Specific products, product lines, or offerings affected
- Business processes or departments impacted
- User groups, roles, or personas affected
- Proposal types, application types, or transaction types
- Geographic or market scope if mentioned

### 2.3 List of APIs required

Extract SPECIFIC technical requirements from the source:
- New APIs or services that need to be created
- Existing APIs that need modification
- Third-party integrations or external system connections
- Database access or query requirements
- Authentication, authorization, or security services
- Any technical specifications or interface requirements

IMPORTANT:

- Use markdown headings (##, ###)
- Preserve any tables in markdown format
- Extract content based on what's ACTUALLY in the source, regardless of domain
- Adapt language and focus to match the source content type
- If no content found for a subsection, leave it blank

VALIDATION CHECK:

Before finalizing each section, verify that every piece of information can be traced back to the source requirements. Remove any content that cannot be directly attributed to the source documents.

OUTPUT FORMAT:
Provide ONLY the markdown sections (## 1.0 Introduction, ### 1.1 Purpose, etc.) with the extracted content. Do not include any of these instructions, validation checks, or processing guidelines in your response.

""",

    "process_requirements": """

You are a Business Analyst expert creating sections 3.0–4.0 of a comprehensive BRD.

PREVIOUS CONTENT:

{previous_content}

SOURCE REQUIREMENTS:

{requirements}

CRITICAL INSTRUCTIONS:

- Extract information ONLY from the provided source requirements
- Identify ACTUAL workflows, processes, and business rules from the source
- Focus on step-by-step logic, conditions, and decision points mentioned
- Adapt to any business domain (training, validation, integration, access control, etc.)
- Do NOT create, assume, or fabricate any content not explicitly present in the source

Create ONLY the following sections with detailed content in markdown:

## 3.0 Process / Data Flow diagram / Figma

Extract DETAILED workflow/process information from the source:

### 3.1 Workflow Description

Create step-by-step process based on what's described in the source:
- What triggers the process or workflow?
- What are the sequential steps or stages?
- What decision points, conditions, or validations occur?
- What are the different paths or outcomes?
- How are errors, exceptions, or edge cases handled?
- What user interactions or system responses are involved?

Format as logical flow:
- Step 1: [Action/Trigger]
  - If [condition mentioned in source]: [result/next step]
  - If [alternative condition]: [alternative result]
- Step 2: [Next Action]
  - [Continue based on source content]

Look for process indicators: "workflow", "process", "steps", "sequence", "flow", "journey", "condition", "if", "then", "when", "trigger", "action", "response"

## 4.0 Business / System Requirement

### 4.1 Application / Module Name: [Extract exact application/module name from source]

Create detailed requirement table based on source content:

| **Rule ID** | **Rule Description** | **Expected Result** | **Dependency** |
|-------------|---------------------|-------------------|----------------|
| **4.1.1** | [Extract specific business rule from source] | [Exact expected behavior mentioned] | [Technical/system dependencies noted] |

Focus on extracting:
- Specific business rules, validations, or logic mentioned
- Functional requirements and expected system behaviors
- User access controls, permissions, or restrictions
- Data validation, processing, or transformation rules
- Integration requirements and system interactions

### 4.2 Functional Requirements

Extract BUSINESS functional requirements from source:
- What specific business functions or capabilities need to be delivered?
- What actions, operations, or tasks should users be able to perform?
- What business processes or workflows need to be supported?
- What decisions, approvals, or validations are required in the process?
- What information needs to be captured, reviewed, or processed?
- What business rules, policies, or compliance requirements must be enforced?
- What notifications, communications, or alerts need to be generated?
- What access controls, permissions, or restrictions are needed?
- What calculations, computations, or business logic processing is required?
- What reporting, tracking, or audit capabilities are needed for the business?

### 4.3 Business Rules and Logic

Extract BUSINESS rules and decision logic from source:
- What business decisions need to be made and under what conditions?
- What criteria determine whether a business action is allowed or restricted?
- What business policies or guidelines govern the process or workflow?
- When should different business outcomes or paths be followed?
- What business conditions trigger specific actions, approvals, or notifications?
- What are the business exceptions, special cases, or override scenarios?
- How do different user roles, responsibilities, or authorities affect business decisions?
- What business validations ensure compliance with regulations or company policies?
- What escalation rules or approval hierarchies apply to business decisions?
- What business status changes or state transitions are required in the process?

IMPORTANT:

- Use markdown headings
- Create detailed requirement tables with multiple columns
- Base all content on what's explicitly stated in the source
- Adapt terminology and focus to match the source domain
- Leave blank if no content found

VALIDATION CHECK:

Before finalizing each section, verify that every piece of information can be traced back to the source requirements. Remove any content that cannot be directly attributed to the source documents.

OUTPUT FORMAT:
Provide ONLY the markdown sections (## 3.0, ### 3.1, etc.) with the extracted content. Do not include any of these instructions, validation checks, or processing guidelines in your response.

""",

    "data_communication": """

You are a Business Analyst expert creating sections 5.0–6.0 of a comprehensive BRD.

PREVIOUS CONTENT:

{previous_content}

SOURCE REQUIREMENTS:

{requirements}

CRITICAL INSTRUCTIONS:

- Extract information ONLY from the provided source requirements
- Identify ACTUAL data and communication needs from the source
- Adapt to any type of data requirements (user data, transaction data, training data, etc.)
- Do NOT create, assume, or fabricate any content not explicitly present in the source

Create ONLY the following sections with detailed content in markdown:

## 5.0 MIS / DATA Requirement

### 5.1 Data Specifications

Extract SPECIFIC data requirements from the source:
- What data elements, fields, or attributes are needed?
- What data sources, databases, or systems provide this data?
- What data formats, structures, or schemas are required?
- What data validation, quality, or integrity requirements exist?
- What data processing, transformation, or calculation needs are mentioned?
- Any data retention, archival, or lifecycle requirements

### 5.2 Reporting and Analytics Needs

Extract any mentioned:
- What reports, dashboards, or analytics are required?
- What metrics, KPIs, or measurements need to be tracked?
- What data visualization or presentation requirements exist?
- What frequency or scheduling of reports is needed?
- What user roles or audiences need access to reports?
- Any real-time monitoring or alerting requirements

### 5.3 Data Sources and Destinations

Extract from source:
- Source systems, databases, or applications providing data
- Target systems, repositories, or destinations for data
- Integration points, APIs, or data exchange mechanisms
- Data flow directions and transformation requirements
- External systems, third-party sources, or partner integrations
- Master data management or reference data needs

## 6.0 Communication Requirement

**SEARCH FOR .MSG FILES ONLY:**
- Scan ALL source documents for original email files (.msg format)
- Extract ONLY emails that are explicitly provided in .msg file format

**IF .msg files ARE FOUND:**
- Include top 3 most relevant original emails from the .msg files
- Preserve original email format and content

**IF NO .msg files are found:**
- Leave this section completely BLANK
- Do NOT generate, create, or simulate any email content
- Do NOT include any sample communications

**CRITICAL:** Never generate, create, assume, or fabricate any email content. Only extract from actual .msg source files.

VALIDATION CHECK:

Before finalizing each section, verify that every piece of information can be traced back to the source requirements. Remove any content that cannot be directly attributed to the source documents.

OUTPUT FORMAT:
Provide ONLY the markdown sections (## 5.0, ### 5.1, etc.) with the extracted content. Do not include any of these instructions, validation checks, or processing guidelines in your response.

""",

    "testing_final": """

You are a Business Analyst expert creating sections 7.0–11.0 of a comprehensive BRD.

PREVIOUS CONTENT:

{previous_content}

SOURCE REQUIREMENTS:

{requirements}

CRITICAL INSTRUCTIONS FOR ALL SECTIONS:

- Extract information ONLY from the provided source requirements
- For Test Scenarios: PRIORITY CHECK - First look for existing test scenarios in source
- Adapt to any business domain or requirement type
- Do NOT create, assume, or fabricate any content not explicitly present in the source

Create ONLY the following sections with detailed content in markdown:

## 7.0 Test Scenarios

**PRIMARY APPROACH - Extract Existing Test Scenarios:**
FIRST, thoroughly scan ALL source requirements documents for existing test content using these keywords:
- "Test Scenarios" / "Test Scenario"
- "Test Cases" / "Test Case" 
- "Test case Scenarios"
- "Testing" / "Test Plan"
- "Verification" / "Validation"

**IF existing test scenarios/cases ARE FOUND in source documents:**
- Extract and preserve ALL the EXACT test scenarios from the source (require all the test scenarios from the source)
- Maintain original test structure, format, and content
- Convert to standardized markdown table format:

| **Test ID** | **Test Scenario Name** | **Objective** | **Test Steps** | **Expected Results** | **Test Data** | **Type** |
|-------------|---------------|---------------|----------------|---------------------|---------------|----------|
| [Extract ID] | [Extract Name] | [Extract Objective] | [Extract Steps] | [Extract Results] | [Extract Data] | [Extract Type] |

**STOP HERE - Do not proceed to Secondary Approach if existing tests are found**

---

**SECONDARY APPROACH - Generate from Functional Requirements:**
**ONLY EXECUTE IF PRIMARY APPROACH YIELDS NO RESULTS**

IF NO existing test scenarios are found in ANY source documents, THEN generate test scenarios based EXCLUSIVELY on functionality explicitly described in source requirements:

| **Test ID** | **Test Name** | **Objective** | **Test Steps** | **Expected Results** | **Test Data** | **Type** |
|-------------|---------------|---------------|----------------|---------------------|---------------|----------|

Create exactly 5 test scenarios covering:
- Primary functional requirements mentioned in source
- Different user roles, permissions, or access levels described  
- Various input conditions, data scenarios, or edge cases noted
- Error conditions, exceptions, or validation failures mentioned
- Integration points, API calls, or system interactions described

**CRITICAL:** Base ALL generated test scenarios ONLY on what is explicitly described in the source requirements. Do not infer or assume functionality not documented.

**EXECUTION RULE:** Use Primary Approach OR Secondary Approach - NEVER BOTH.

## 8.0 Questions / Suggestions

**SEARCH STRATEGY:**
- Scan ALL source documents for explicit questions, suggestions, or clarifications
- Look for keywords: "Question", "Clarification", "Unknown", "Assumption", "Dependency", "Suggestion", "Recommendation"

**IF questions/suggestions ARE FOUND in source:**
- List exact questions, clarifications, or unknowns from source
- List exact assumptions, dependencies, or prerequisites from source  
- List exact suggestions, recommendations, or enhancements from source

**IF NO questions/suggestions are found in source:**
- Leave this section completely BLANK
- Do NOT generate, create, or assume any questions or suggestions
- Do NOT infer potential issues or recommendations

**CRITICAL:** Only extract content that is explicitly stated in the source documents. Never generate, create, assume, or fabricate any questions, suggestions, or recommendations not present in the source.

## 9.0 Reference Document

List exact source documents

**CRITICAL:** Only extract content that is explicitly stated in the source documents. Never generate, create, assume, or fabricate anything

## 10.0 Appendix

**SEARCH STRATEGY:**
- Scan ALL source documents for explicit appendix and supporting information
- Look for keywords: "Appendix"

**IF appendix/supporting information ARE FOUND in source:**
- List exact appendix and supporting information

**IF NO appendix/supporting information are found in source:**
- Leave this section completely BLANK
- Do NOT generate, create, or assume any appendix or supporting information
- Do NOT infer potential supporting information

## 11.0 Risk Evaluation

**SEARCH STRATEGY:**
- Scan for sections: "Risk Evaluation", "Risk Assessment", "Risk Analysis"
- Look for ANY structured data that appears to be in rows and columns

**EXTRACTION AND FORMATTING:**
- **IF content appears to be tabular data (even if extracted as plain text):**
  * Identify column headers from the source
  * Identify data rows from the source  
  * Convert to markdown table format using this structure:

| [Header 1 from source] | [Header 2 from source] | [Header 3 from source] | [Additional headers as found] |
|------------------------|------------------------|------------------------|------------------------------|
| [Row 1 data from source] | [Row 1 data] | [Row 1 data] | [Row 1 data] |
| [Row 2 data from source] | [Row 2 data] | [Row 2 data] | [Row 2 data] |

- **Preserve exact content** from source but **force table formatting**
- **Do not modify** any text content, only structure it into table format

**IF NO Risk Evaluation content found:**
- State: "No Risk Evaluation section found in source document"

**CRITICAL:** Even if the source content appears as plain text, if it contains structured risk data, format it as a markdown table while preserving all original content exactly.

IMPORTANT:

- Use markdown headings
- Preserve tables with markdown table format using pipe syntax
- Do NOT output code fences
- Create test scenarios based on ACTUAL requirements described in source
- For sections 8.0-11.0: Leave blank if no content found in source
- Adapt all content to match the domain and scope of the source requirements

VALIDATION CHECK:

Before finalizing sections 8.0-11.0, verify that every piece of information can be traced back to the source requirements. Remove any content that cannot be directly attributed to the source documents.

OUTPUT FORMAT:
Provide ONLY the markdown sections (## 7.0, ### 7.1, etc.) with the extracted content. Do not include any of these instructions, validation checks, or processing guidelines in your response.

"""

}

def estimate_content_size(text):
    return len(text)

def chunk_requirements(requirements, max_chunk_size=8000):
    if estimate_content_size(requirements) <= max_chunk_size:
        return [requirements]
    
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
def initialize_sequential_chains(api_provider, api_key, azure_endpoint=None, azure_deployment=None, api_version=None):
    
    if api_provider == "OpenAI":
        model = ChatOpenAI(
            openai_api_key=api_key,
            model_name="gpt-3.5-turbo-16k",
            temperature=0.2,
            top_p=0.2
        )
    elif api_provider == "AzureOpenAI":
        model = AzureChatOpenAI(
            azure_endpoint=azure_endpoint,
            openai_api_key=api_key,
            azure_deployment=azure_deployment,
            api_version=api_version,
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
    
    chains = []
    chain1 = LLMChain(
        llm=model,
        prompt=PromptTemplate(
            input_variables=['requirements'],
            template=SECTION_TEMPLATES["intro_impact"]
        ),
        output_key="intro_impact_sections"
    )
    
    chain2 = LLMChain(
        llm=model,
        prompt=PromptTemplate(
            input_variables=['previous_content', 'requirements'],
            template=SECTION_TEMPLATES["process_requirements"]
        ),
        output_key="process_requirements_sections"
    )
    
    chain3 = LLMChain(
        llm=model,
        prompt=PromptTemplate(
            input_variables=['previous_content', 'requirements'],
            template=SECTION_TEMPLATES["data_communication"]
        ),
        output_key="data_communication_sections"
    )
    
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
    req_chunks = chunk_requirements(requirements)
    
    if len(req_chunks) > 1:
        st.info(f"Large content detected. Processing in {len(req_chunks)} chunks...")
    
    combined_requirements = "\n\n=== DOCUMENT BREAK ===\n\n".join(req_chunks)
    
    # Print the combined requirements being sent to the first chain
    st.write("="*80)
    st.write("COMBINED REQUIREMENTS SENT TO LLM:")
    st.write("="*80)
    st.write(combined_requirements)
    st.write("="*80)
    st.write(f"Total characters: {len(combined_requirements)}")
    st.write("="*80)
    
    previous_content = ""
    final_sections = []
    
    for i, chain in enumerate(chains):
        try:
            print(f"\n{'='*60}")
            print(f"CHAIN {i+1} INPUT:")
            print(f"{'='*60}")
            
            if i == 0:
                print("Input to Chain 1 (intro_impact):")
                print(f"Requirements length: {len(combined_requirements)} characters")
                print("First 500 characters of requirements:")
                print(combined_requirements[:500] + "..." if len(combined_requirements) > 500 else combined_requirements)
                
                result = chain.run(requirements=combined_requirements)
            else:
                print(f"Input to Chain {i+1}:")
                print(f"Previous content length: {len(previous_content)} characters")
                print(f"Requirements length: {len(combined_requirements)} characters")
                print("Previous content (first 300 chars):")
                print(previous_content[:300] + "..." if len(previous_content) > 300 else previous_content)
                print("\nRequirements (first 300 chars):")
                print(combined_requirements[:300] + "..." if len(combined_requirements) > 300 else combined_requirements)
                
                result = chain.run(previous_content=previous_content, requirements=combined_requirements)
            
            print(f"\nCHAIN {i+1} OUTPUT:")
            print(f"Response length: {len(result)} characters")
            print("First 500 characters of response:")
            print(result[:500] + "..." if len(result) > 500 else result)
            print(f"{'='*60}")
            
            final_sections.append(result)
            previous_content += "\n\n" + result
            
            st.write(f"✓ Completed section group {i+1}/4")
            
        except Exception as e:
            print(f"ERROR in chain {i+1}: {str(e)}")
            st.error(f"Error in chain {i+1}: {str(e)}")
            final_sections.append(f"## Error in section group {i+1}\nError processing this section: {str(e)}")
    
    final_brd = "\n\n".join(final_sections)
    
    st.write("\n" + "="*80)
    st.write("FINAL BRD CONTENT:")
    st.write("="*80)
    st.write(f"Total final BRD length: {len(final_brd)} characters")
    st.write("Final BRD (first 1000 characters):")
    st.write(final_brd[:1000] + "..." if len(final_brd) > 1000 else final_brd)
    st.write("="*80)
    
    return final_brd

def create_toc_styles(doc):
    styles = doc.styles
    
    try:
        toc1_style = styles['TOC 1']
    except KeyError:
        toc1_style = styles.add_style('TOC 1', WD_STYLE_TYPE.PARAGRAPH)
        toc1_style.font.name = 'Calibri'
        toc1_style.font.size = Pt(11)
        toc1_style.paragraph_format.left_indent = Inches(0)
        toc1_style.paragraph_format.space_after = Pt(0)
    
    try:
        toc2_style = styles['TOC 2']
    except KeyError:
        toc2_style = styles.add_style('TOC 2', WD_STYLE_TYPE.PARAGRAPH)
        toc2_style.font.name = 'Calibri'
        toc2_style.font.size = Pt(11)
        toc2_style.paragraph_format.left_indent = Inches(0.25)
        toc2_style.paragraph_format.space_after = Pt(0)

def create_clickable_toc(doc):
    toc_heading = doc.add_heading('Table of Contents', level=1)
    add_bookmark(toc_heading, 'TOC')
    
    create_toc_styles(doc)
    
    toc_entries = [
        ("1.0 Introduction", "introduction"),
        ("    1.1 Purpose", "purpose"),
        ("    1.2 As-is process", "process_solution"),
        ("    1.3 To be process / High level solution", "process_solution"),
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

    bookmark_mapping = {}
    for entry_text, bookmark_name in toc_entries:
        bookmark_mapping[bookmark_name] = entry_text
        
    for entry_text, bookmark_name in toc_entries:
        toc_paragraph = doc.add_paragraph()
        
        try:
            if entry_text.startswith("    "):
                toc_paragraph.style = 'TOC 2'
            else:
                toc_paragraph.style = 'TOC 1'
        except KeyError:
            if entry_text.startswith("    "):
                toc_paragraph.paragraph_format.left_indent = Inches(0.25)
            toc_paragraph.paragraph_format.space_after = Pt(0)
        
        if entry_text.startswith("    "):
            toc_paragraph.add_run("    ")
            link_text = entry_text.strip()
        else:
            link_text = entry_text
            
        add_hyperlink(toc_paragraph, link_text, bookmark_name, is_internal=True)
        
        toc_paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(6.0))
        
        toc_paragraph.add_run("\t")
        
        page_run = toc_paragraph.add_run()
        
        fldChar_begin = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>')
        instrText = parse_xml(f'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> PAGEREF {bookmark_name} \\h </w:instrText>')
        fldChar_end = parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>')
        
        page_run._r.append(fldChar_begin)
        page_run._r.append(instrText)
        page_run._r.append(fldChar_end)
        
        page_run.add_text(" ")
    
    doc.add_paragraph()
    note_para = doc.add_paragraph()
    note_run = note_para.add_run("IMPORTANT: ")
    note_run.bold = True
    note_run.font.color.rgb = RGBColor(255, 0, 0)
    
    note_para.add_run("To see actual page numbers in this Table of Contents:")
    note_para.add_run("Press 'Ctrl + A' to select all, then F9 to update all fields in the document.")
    
    return bookmark_mapping

def add_hyperlink(paragraph, text, url_or_bookmark, is_internal=True):
    part = paragraph.part
    
    hyperlink = OxmlElement('w:hyperlink')
    
    if is_internal:
        hyperlink.set(qn('w:anchor'), url_or_bookmark)
    else:
        r_id = part.relate_to(url_or_bookmark, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        hyperlink.set(qn('r:id'), r_id)
    
    new_run = OxmlElement('w:r')
    
    rPr = OxmlElement('w:rPr')
    
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')
    rPr.append(color)
    
    underline = OxmlElement('w:u')
    underline.set(qn('w:val'), 'single')
    rPr.append(underline)
    
    new_run.append(rPr)
    
    text_element = OxmlElement('w:t')
    text_element.text = text
    new_run.append(text_element)
    
    hyperlink.append(new_run)
    
    paragraph._p.append(hyperlink)
    
    return hyperlink

def add_bookmark(paragraph, bookmark_name):
    bookmark_id = str(abs(hash(bookmark_name)) % 1000000)
    bookmark_start = OxmlElement('w:bookmarkStart')
    bookmark_start.set(qn('w:id'), bookmark_id)
    bookmark_start.set(qn('w:name'), bookmark_name)
    
    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), bookmark_id)
    
    paragraph._p.insert(0, bookmark_start)
    paragraph._p.append(bookmark_end)

def add_section_with_bookmark(doc, heading_text, bookmark_name, level=1):
    heading = doc.add_heading(heading_text, level=level)
    add_bookmark(heading, bookmark_name)
    
    return heading

def create_table_in_doc(doc, table_data):
    if not table_data or len(table_data) < 2:
        return None
    
    table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
    table.style = 'Table Grid'
    
    for i, cell_text in enumerate(table_data[0]):
        cell = table.rows[0].cells[i]
        cell.text = str(cell_text) if cell_text else ""
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
    
    for row_idx, row_data in enumerate(table_data[1:], 1):
        for col_idx, cell_text in enumerate(row_data):
            if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                table.rows[row_idx].cells[col_idx].text = str(cell_text) if cell_text else ""
    
    return table

def parse_markdown_table(table_text):
    lines = [line.strip() for line in table_text.split('\n') if line.strip()]
    
    if len(lines) < 2:
        return None
    
    if len(lines) >= 2 and '---' in lines[1]:
        lines.pop(1)
    
    table_data = []
    for line in lines:
        if line.startswith('|') and line.endswith('|'):
            line = line[1:-1]
            cells = [cell.strip() for cell in line.split('|')]
            table_data.append(cells)
        else:
            cells = [cell.strip() for cell in line.split('|')]
            table_data.append(cells)
    
    return table_data if table_data else None

def extract_content_from_docx(doc_file):
    doc = Document(doc_file)
    content = []
    
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            content.append(paragraph.text.strip())
    
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

def extract_content_from_excel(excel_file, max_rows_per_sheet=70, max_sample_rows=10, visible_only=False):
    content = []
    try:
        if visible_only:
            wb = load_workbook(excel_file)
            visible_sheets = []
            
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                if sheet.sheet_state == 'visible':
                    visible_sheets.append(sheet_name)
            
            if visible_sheets:
                excel_data = pd.read_excel(excel_file, sheet_name=visible_sheets)
            else:
                return "No visible sheets found in the Excel file"
            
            if not isinstance(excel_data, dict):
                excel_data = {visible_sheets[0]: excel_data}
        else:
            excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        for sheet_name, df in excel_data.items():
            if df.empty:
                continue
            
            if max_rows_per_sheet and len(df) > max_rows_per_sheet:
                df = df.head(max_rows_per_sheet)
                content.append(f"Note: Processing first {max_rows_per_sheet} rows only")
                
            content.append(f"=== EXCEL SHEET: {sheet_name} ===")
            content.append(f"Total Dimensions: {df.shape[0]} rows × {df.shape[1]} columns")
            
            content.append(f"Columns ({len(df.columns)}): {', '.join(df.columns.tolist())}")
            
            data_types = df.dtypes.to_dict()
            type_summary = []
            for col, dtype in data_types.items():
                type_summary.append(f"{col}: {str(dtype)}")
            content.append(f"Data Types: {', '.join(type_summary[:10])}...")
            
            numeric_cols = df.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                content.append(f"Numeric Columns: {', '.join(numeric_cols.tolist()[:5])}...")
            
            sample_size = min(max_sample_rows, len(df))
            if sample_size > 0:
                content.append(f"\nSample Data (first {sample_size} rows):")
                content.append("TABLE:")
                
                display_df = df.head(sample_size)
                
                if len(df.columns) > 10:
                    display_cols = df.columns[:8].tolist() + [f"... +{len(df.columns)-8} more columns"]
                    display_df = df[df.columns[:8]].head(sample_size)
                    header_row = " | ".join(display_cols)
                    content.append(header_row)
                else:
                    header_row = " | ".join(df.columns.tolist())
                    content.append(header_row)
                
                for _, row in display_df.iterrows():
                    row_data = []
                    for val in row:
                        str_val = str(val)
                        if len(str_val) > 50:
                            str_val = str_val[:47] + "..."
                        row_data.append(str_val)
                    content.append(" | ".join(row_data))
                
                if len(df) > sample_size:
                    content.append(f"... and {len(df) - sample_size} more rows")
            
            content.append(f"\nData Summary:")
            
            key_columns = []
            for col in df.columns:
                col_lower = col.lower()
                if any(keyword in col_lower for keyword in ['id', 'name', 'title', 'status', 'type', 'category', 'priority', 'requirement']):
                    key_columns.append(col)
            
            if key_columns:
                content.append(f"Key Columns Identified: {', '.join(key_columns[:5])}")
                
                for col in key_columns[:3]:
                    if df[col].dtype == 'object':
                        unique_vals = df[col].dropna().unique()
                        if len(unique_vals) <= 20:
                            content.append(f"{col} Values: {', '.join(map(str, unique_vals[:10]))}")
                        else:
                            content.append(f"{col}: {len(unique_vals)} unique values")
            
            missing_data = df.isnull().sum()
            if missing_data.sum() > 0:
                missing_cols = missing_data[missing_data > 0].head(5)
                missing_info = [f"{col}: {count} missing" for col, count in missing_cols.items()]
                content.append(f"Missing Data: {', '.join(missing_info)}")
            
            content.append("="*50)
    
    except Exception as e:
        content.append(f"Error processing Excel file: {str(e)}")
    
    return "\n".join(content)

def extract_content_from_msg(msg_file):
    try:
        temp_file = BytesIO(msg_file.getvalue())
        temp_file.name = msg_file.name
        
        msg = extract_msg.Message(temp_file)
        body_content = msg.body
        
        cleaned_body = re.sub(r'^(From|To|Cc|Subject|Sent|Date):.*?\n', '', body_content, flags=re.MULTILINE)
        cleaned_body = re.sub(r'_{10,}[\s\S]*$', '', cleaned_body)
        cleaned_body = re.sub(r'-{10,}[\s\S]*$', '', cleaned_body)
        
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
    
    if logo_data:
        add_header_with_logo(doc, logo_data)
    
    for _ in range(12):
        doc.add_paragraph()
    
    title = doc.add_heading('Business Requirements Document', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    
    doc.add_heading('Version History', level=1)
    version_table = doc.add_table(rows=5, cols=5)
    version_table.style = 'Table Grid'
    hdr_cells = version_table.rows[0].cells
    headers = ['Version', 'Date', 'Author', 'Change description', 'Review by']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    doc.add_paragraph('**To be reviewed and filled in by IT Team.**')
    
    doc.add_heading('Sign-off Matrix', level=1)
    signoff_table = doc.add_table(rows=5, cols=5)
    signoff_table.style = 'Table Grid'
    hdr_cells = signoff_table.rows[0].cells
    headers = ['Version', 'Sign-off Authority', 'Business Function', 'Sign-off Date', 'Email Confirmation']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    doc.add_page_break()
    
    bookmark_mapping = create_clickable_toc(doc)
    if bookmark_mapping is None:
        bookmark_mapping = {}
    
    doc.add_page_break()
    
    sections = content.split('##')
    
    introduction_started = False

    for i, section in enumerate(sections):
        if section.strip():
            lines = section.strip().split('\n')
            if lines:
                heading_line = lines[0].strip()
            
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
            
                section_name_lower = heading_text.lower()
                if 'introduction' in section_name_lower or section_name_lower.startswith('1.0'):
                    introduction_started = True
            
                if level == 1 and i > 0 and not introduction_started:
                    doc.add_page_break()
            
                if bookmark_name:
                    add_section_with_bookmark(doc, heading_text, bookmark_name, level)
                else:
                    doc.add_heading(heading_text, level)
            
                j = 1
                while j < len(lines):
                    line = lines[j].strip()
                
                    if line and '|' in line and line.count('|') >= 2:
                        table_lines = []
                        while j < len(lines) and lines[j].strip() and '|' in lines[j]:
                            table_lines.append(lines[j].strip())
                            j += 1
                    
                        if table_lines:
                            table_data = parse_markdown_table('\n'.join(table_lines))
                            if table_data:
                                create_table_in_doc(doc, table_data)
                        continue
                
                    if line:
                        if line.startswith('- ') or line.startswith('* '):
                            doc.add_paragraph(line[2:].strip(), style='List Bullet')
                        elif re.match(r'^\d+\.', line):
                            doc.add_paragraph(re.sub(r'^\d+\.\s*', '', line), style='List Bullet')
                        else:
                            doc.add_paragraph(line)
                
                    j += 1
    
    return doc

st.title("Business Requirements Document Generator")

st.subheader("AI Model Selection")
api_provider = st.radio("Select API Provider:", ["OpenAI", "Groq", "AzureOpenAI"])

if api_provider == "OpenAI":
    api_key = st.text_input("Enter your OpenAI API Key:", type="password")
elif api_provider == "AzureOpenAI":
    api_key = st.text_input("Enter your Azure OpenAI API Key:", type="password")
    azure_endpoint = st.text_input("Enter your Azure OpenAI Endpoint:", 
                                   placeholder="https://your-resource.openai.azure.com/")
    azure_deployment = st.text_input("Enter your Azure Deployment Name:", 
                                     placeholder="gpt-35-turbo")
    api_version = st.text_input("Enter API Version (optional):", 
                                value="2024-02-15-preview",
                                placeholder="2024-02-15-preview")
else:
    api_key = st.text_input("Enter your Groq API Key:", type="password")

st.subheader("Document Logo")

logo_file = st.file_uploader("Upload Company Logo (optional):", type=['png', 'jpg', 'jpeg'])
logo_data = None
if logo_file:
    logo_data = logo_file.getvalue()
    st.success("Logo uploaded successfully!")

st.subheader("Upload Requirements Documents")
uploaded_files = st.file_uploader(
    "Choose files", 
    type=['txt', 'docx', 'pdf', 'xlsx', 'xls', 'msg'],
    accept_multiple_files=True
)

st.subheader("Or Enter Requirements Manually")
manual_requirements = st.text_area(
    "Paste your requirements here:",
    height=200,
    placeholder="Enter your business requirements, user stories, or project specifications here..."
)

if st.button("Generate BRD", type="primary"):
    if not api_key:
        st.error("Please enter your API key!")
    elif not uploaded_files and not manual_requirements.strip():
        st.error("Please upload files or enter requirements manually!")
    else:
        try:
            with st.spinner("Initializing AI chains"):
                chains = initialize_sequential_chains(api_provider=api_provider,
            api_key=api_key,
            azure_endpoint=azure_endpoint,
            azure_deployment=azure_deployment,
            api_version=api_version)
            
            all_requirements = []
            
            if manual_requirements.strip():
                all_requirements.append("=== MANUAL REQUIREMENTS ===")
                all_requirements.append(manual_requirements.strip())
                all_requirements.append("="*50)
            
            if uploaded_files:
                st.info(f"Processing {len(uploaded_files)} uploaded files...")
                
                for uploaded_file in uploaded_files:
                    file_extension = uploaded_file.name.split('.')[-1].lower()
                    
                    try:
                        st.write(f"Processing: {uploaded_file.name}")
                        
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
                            st.warning(f"⚠Unsupported file type: {file_extension}")
                            continue
                        
                        if content.strip():
                            all_requirements.append(f"=== FILE: {uploaded_file.name} ===")
                            all_requirements.append(content.strip())
                            all_requirements.append("="*50)
                            st.success(f"Successfully processed: {uploaded_file.name}")
                        else:
                            st.warning(f"No content extracted from: {uploaded_file.name}")
                            
                    except Exception as e:
                        st.error(f"Error processing {uploaded_file.name}: {str(e)}")
                        continue
            
            if not all_requirements:
                st.error("No valid content found in uploaded files!")
                st.stop()
            
            combined_requirements = "\n\n".join(all_requirements)
            
            content_size = estimate_content_size(combined_requirements)
            st.info(f"Total content size: {content_size:,} characters")
            
            st.subheader("AI Processing Progress")
            
            with st.spinner("Generating comprehensive BRD using sequential processing..."):
                brd_content = generate_brd_sequentially(chains, combined_requirements)
            
            if brd_content:
                st.success("BRD generated successfully!")
                
                st.subheader("Generated BRD Content")
                
                with st.expander("Preview Generated BRD", expanded=False):
                    st.markdown(brd_content)
                
                st.subheader("Download Options")
                
                try:
                    with st.spinner("Creating Word document..."):
                        doc = create_word_document(brd_content, logo_data)
                        
                        doc_buffer = BytesIO()
                        doc.save(doc_buffer)
                        doc_buffer.seek(0)
                        
                        st.download_button(
                            label="Download BRD (Word Document)",
                            data=doc_buffer.getvalue(),
                            file_name="Business_Requirements_Document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                        st.success("Word document ready for download!")
                        
                except Exception as e:
                    st.error(f"Error creating Word document: {str(e)}")
                    st.info("You can still copy the content above manually.")
                
                try:
                    st.download_button(
                        label="Download BRD (Markdown)",
                        data=brd_content,
                        file_name="Business_Requirements_Document.md",
                        mime="text/markdown"
                    )
                except Exception as e:
                    st.error(f"Error creating markdown download: {str(e)}")
                
            else:
                st.error("Failed to generate BRD content!")
                
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.info("Try reducing the input size or check your API key.")
