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
import json

BRD_FORMAT = """
## 1.0 Introduction
    ## 1.1 Purpose
    ## 1.2 As-is process
    ## 1.3 To be process / High level solution
## 2.0 Impact Analysis
    ## 2.1 Impacted Products
    ## 2.2 Applications Impacted
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

EXCEL FILE PROCESSING INSTRUCTIONS:

**FOR EXCEL FILES (.xlsx/.xls):**
- Process ALL sheets EXCEPT "Test Scenarios" sheet
- **PRIORITY FOCUS**: Look specifically for "PART B : (Mandatory) Detailed Requirement" section in any sheet
- Include data from sheets: "Requirement", "Ops Risk Assessment", and any other available sheets
- Extract content from ALL relevant columns and rows in each sheet
- Look for business requirements, processes, impacts, and technical specifications across all sheets
- If sheet names are different from expected, process all sheets except those explicitly containing test scenarios

**SPECIAL INSTRUCTION FOR PURPOSE AND TO-BE PROCESS:**
For sections 1.1 Purpose and 1.3 To be process / High level solution:
- **PRIMARY PRIORITY**: Search for and extract information from "PART B : (Mandatory) Detailed Requirement" section
- Look for this exact text or similar variations like:
  - "PART B"
  - "Mandatory Detailed Requirement" 
  - "Detailed Requirement"
  - "Part B - Detailed Requirement"
  - "PART B : Detailed Requirement"
- If found, prioritize this section's content for Purpose and To-be process extraction
- If not found, then search across ALL other processed sheets for relevant content

CRITICAL INSTRUCTIONS:

- Extract information from ALL available sheets (except Test Scenarios sheet)
- **For Purpose and To-be process: PRIORITIZE "PART B : (Mandatory) Detailed Requirement" content**
- Identify the ACTUAL business problem being solved from any relevant sheet
- Focus on what is explicitly mentioned across all processed sheets
- Do NOT create, assume, or fabricate any content not present in the source
- If a section has no relevant information across ALL processed sheets, leave it BLANK
- Adapt to any domain (training, payments, integration, access control, etc.)

Create ONLY the following sections with detailed content in markdown:

## 1.0 Introduction

### 1.1 Purpose

**SEARCH STRATEGY FOR PURPOSE:**
1. **FIRST PRIORITY**: Look specifically for "PART B : (Mandatory) Detailed Requirement" section
2. **SECOND PRIORITY**: Search other sections in "Requirement" sheet

Extract the EXACT business purpose, focusing on:
- Capture from "PART B : (Mandatory) Detailed Requirement" if available
- If PART B not found, capture from "Detailed Requirement" sections in any sheet
- What is the main business objective or problem being addressed?
- What specific functionality or capability is being implemented?
- What restrictions, validations, or controls are being introduced?
- What business processes are being improved or changed?
- What compliance, security, or operational requirements are being met?

Search across ALL processed sheets for key phrases: "purpose", "objective", "requirement", "need", "problem", "solution", "implement", "restrict", "validate", "improve", "ensure"

**EXTRACTION PRIORITY ORDER:**
1. Content from "PART B : (Mandatory) Detailed Requirement" 
2. Content from other "Detailed Requirement" sections
3. Content from other relevant sections across all sheets

OUTPUT FORMAT:
1. In form of bullet pointers

### 1.2 As-is process

Extract the CURRENT state/process from ANY relevant sheet:
- How does the current system/process work?
- What are the existing workflows or user journeys?
- What problems or limitations exist in the current approach?
- What manual processes or workarounds are currently used?
- What system behaviors need to be changed?
- Any screenshots, process flows, or current state descriptions

Look for indicators across ALL sheets: "currently", "as-is", "existing", "present", "manual", "workaround", "problem with current", "limitations"

### 1.3 To be process / High level solution

**SEARCH STRATEGY FOR TO-BE PROCESS:**
1. **FIRST PRIORITY**: Look specifically for "PART B : (Mandatory) Detailed Requirement" section
2. **SECOND PRIORITY**: Search other sections in "Requirement" sheet  
3. **THIRD PRIORITY**: Search "Ops Risk Assessment" and other sheets

Extract the PROPOSED solution, prioritizing "PART B : (Mandatory) Detailed Requirement" content:
- Content from "PART B : (Mandatory) Detailed Requirement" if available
- What is the new process or system behavior?
- What workflow steps or validation logic will be implemented?
- How will the new solution address current problems?
- What automated processes will replace manual ones?
- What new capabilities or features will be added?
- Any conditional logic, decision trees, or multi-step processes

Look for indicators across ALL sheets: "to-be", "proposed", "solution", "new process", "will be", "should be", "automated", "enhanced", "improved", "step-by-step", "workflow", "condition", "if-then"

**EXTRACTION PRIORITY ORDER:**
1. Content from "PART B : (Mandatory) Detailed Requirement"
2. Content from other "Detailed Requirement" sections  
3. Content from other relevant sections across all sheets

## 2.0 Impact Analysis

### 2.1 Impacted Products

**PRIORITY SEARCH STRATEGY:**
1. **FIRST PRIORITY**: Look for tables/sections with headers ONLY like "Products Impacted", "Product Impacted", "Type of Product" from "Part C"
2. **SECOND PRIORITY**: Look for tables containing product names: ULIP, Term, Endowment, Annuity, Health, Group
3. **THIRD PRIORITY**: Search for any mention of products across all sheets

**EXTRACTION INSTRUCTIONS:**
- If a structured table is found (like the example with ULIP, Term, Endowment, etc.), extract it in the following format:
  - Create a markdown table preserving the original structure
  - Show product types and their impact status (Yes/No/etc.)
  - Include any additional columns or classifications found
  
**TABLE FORMAT EXAMPLE (if structured table found):**
| Product Type | |--------------| | [Extract from source]
|Impact Status|---------------| [Extract Yes/No/etc.] |

**IF NO STRUCTURED TABLE FOUND:**
- List ONLY the products/platforms explicitly mentioned across ALL processed sheets which are impacted
- Extract from any column/row mentioning affected products/platforms
- Check all sheets for product names, service names, or system names, platform names

**CRITICAL**: If a "Products Impacted" table exists in the source, reproduce it exactly as a markdown table. Do NOT create a generic list.

### 2.2 Applications Impacted

**PRIORITY SEARCH STRATEGY:**
1. **FIRST PRIORITY**: Look for tables/sections with headers ONLY like "Applications Impacted", "Application Impacted", "Application Name" from "Part C"
2. **SECOND PRIORITY**: Look for tables containing application names: OPUS, INSTAB, NGIN, PMAC, CRM, Cashier
3. **THIRD PRIORITY**: Search for any mention of applications across all sheets

**EXTRACTION INSTRUCTIONS:**
- If a structured table is found (like the example with OPUS, INSTAB, NGIN, etc.), extract it in the following format:
  - Create a markdown table preserving the original structure
  - Show application names and their impact status (Yes/No/etc.)
  - Include any additional columns or classifications found
  
**TABLE FORMAT EXAMPLE (if structured table found):**
| Application Name | Impact Status | Additional Notes |
|------------------|---------------|------------------|
| [Extract from source] | [Extract Yes/No/etc.] | [Any other info] |

**IF NO STRUCTURED TABLE FOUND:**
- List ONLY the applications explicitly mentioned across ALL processed sheets which are impacted
- Extract from any column/row mentioning affected applications
- Check all sheets for application names

**CRITICAL**: If an "Applications Impacted" table exists in the source, reproduce it exactly as a markdown table. Do NOT create a generic list.

### 2.3 List of APIs required

Understand the what the input requirement states and as per that from the below list select the show the API required:

Policy
API related to details about premium payment,account statments , profile detail, fund details , statement, products etc.

POST
​/GetDetailsService​/getDetails
This service is used to get the details of customer based on policy no and request type.

POST
​/GetUnclaimedAmountService​/getUnclaimedAmount
This API is used to get Unclaimed Amount

POST
​/MyProfileService​/getMyProfile
The Api provides the profile details of a user based on the customer_id passed

POST
​/GetDashboardDetailsService​/getDashboardDetail
This Service is used to fetch dashboard details

POST
​/RenewalDueListService​/getRenewalDueList
get Renewal Due List , providers are Opus and Ngin

POST
​/FetchCategoryDetailService​/fetchCategoryDetail
fetch details for given category , provider is NGIN

POST
​/GetChargeDetailsService​/getChargeDetails
calculate charges , provider is NGIN

POST
​/InterestWaiverDetailService​/getInterestWaiverDetails
get InterestWaiver Details , provide is CRM

POST
​/LoyaltyDetailService​/getLoyaltyDetailsByPolicy
get Loyalty Details , provider in CRM Legacy and Ngin

POST
​/SumAssuredCalculationService​/calculateSumAssured
calculates sum assured on basis of premium amount, providers are Opus and Ngin

POST
​/getPremiumReceiptService​/getPremiumReceipt
get Premium Receipt Details , provider is Opus/ICSM

POST
​/NameSearchService​/searchName
Retrieves registered PH,Policy number against agent code , providers are Opus and Ngin

POST
​/CRMAppSrchMsupDtls​/CRMAppSrchMsupDtlsProxy​/getDetails
Mashup Service to retrieve policydetails , providers are CRM and NGIN

POST
​/SMSDetailsByPolicyService​/getSMSDetailsByPolicy
get SMS Details By Policy , provider is CRM

POST
​/RenewalMailerDetailService​/getRenewalMailerDetails
Retrieves mail details sent for renewal tracking , providers are CRM and Ngin

POST
​/FatcaNRIService​/getFatcaNRI
FATCA NRI details of the customer for the given policy number , providers are CRM and Ngin (to-be)

POST
​/AutoCallingDetailsServices​/getAutoCallingDetails
generates calling details for given number of policy no

POST
​/PasLoginService​/pasLogin
returns url for each sr category post processing , provider is Ngin

POST
​/UpdateFundAlertSMSService​/AlertSMS
updates frequncy of sms alert notification opted , providers are Ngin and CRM

POST
​/FundDetailsService​/getFundDetails
Retrieves Policy Fund & Nav (Net Asset Value) details , providers are Opus , CRM and Ngin

POST
​/GetFundsService​/getFunds
This API is used to get the Funds Bencmarking Details. provider is GBO (Goal Based Orientation)

POST
​/PolicyDetailsService​/getPolicyDetails
Retrives Policy Details (Inception Date,Frequency,etc) , providers are Ngin and Opus

POST
​/GetPayeeDetailsService​/getPayeeDetails
Get Payee Details

POST
​/GenerateNGINTokenService​/generateNGINToken
generate NGIN Token


Policy Servicing
Customer service requests related to exisitng policy , like address change , PAN/Aadhar Update, name and dob corrections are part of Policy Servicing.



POST
​/UpdateFTReceiptStatusService​/updateFTReceiptStatus
This API is used to update FT Receipt status

POST
​/ValidatePartialWithdrawalAmountService​/validatePartialWithdrawalAmount
This API is used to validate Partial Withdrawal Amount

POST
​/SaveAnswerService​/saveAnswer
The Api saves the question list and gives a status which indicates its success

POST
​/SubmitRevivalDGHResponseService​/submitRevivalDGHResponse
The Api is used to submit the revival DGH request for the policy no

POST
​/SubmitPartialWithdrawalRequestService​/updatePartialWithdrawal
This API is used to submit the partial withdrawal request and generate a service request number for same.

POST
​/GPARequestDetailsService​/saveGPARequestDetails

POST
​/AutopayDetailsService​/saveAutopayDetails
Save Autopayment Details

POST
​/CalculateFeesGSTandTDSService​/calculateFeesGSTandTDS
calculate taxes applicable on transactions , provider is PAS .

POST
​/FTRequestDetailsService​/saveFTRequestDetails
save fund transfer Request , update happens in CRM or Ngin system

POST
​/RetrieveModalPremiumForSelectedFrequency​/getRetrieveModalPremium
get ModalPremium For given policy and payment frequency, providers are NGIN and Opus

POST
​/NAVHistoryService​/getNAVHistory
get NAV(Net Asset value) History , providers are Ngin and Opus

POST
​/PayoutDetailsService​/getPayoutDetails
get Payout Details, providers are Ngin and CRM

POST
​/PremiumRedirectionService​/savePremiumRedirection
save Premium Redirection Details , update happens in Ngin or Customer Portal


POST
​/UWCommentsService​/getUWComments
get underwriter Comments for Policy , providers are NGIN or CMR and iONE

POST
​/RenewalIntimationCoverDetailsService​/getcoverdetails
returns coverage details of a policy

POST
​/GenerateSMSService​/generateSMS
generate SMS , provider is CRM

POST
​/validateCOEService​/validateCOE
validate Certificate of Existence , provider is Opus

POST
​/LoanDetailService​/getLoanDetails
get Loan Details , provider is CRM

POST
​/BonusDetailService​/getBonusDetails
get Bonus Details , provider is CRM

POST
​/NomineeDetailsService​/NomineeDetailsService_PS​/getNomineeDetails
get policy holder's Nominee Detail , providers are Ngin, Opus and CRM

POST
​/ValidateSettlementService​/validateSettlement
Validate settlement for the given policy number , providers are CRM and Ngin

POST
​/UpdatePremiumFrequencyService​/updatePremiumFrequency
Update Premium frequency for the given policy detail , details update in Ngin and CRM system based on Policy Number

POST
​/ClickToCallService​/getCall
Customer details & mobile number are saved to opt for clicktocall , providers are Ngin and Opus

POST
​/ViewRequestHistoryService​/viewRequestHistory
checks serice request history against each policy , providers are Ngin and Opus

POST
​/MRNachStatusService​/updateNachStatus
Updates Nach(Bank,transaction , etc) details for policyNo , update in Ngin(CRT) and Opus

POST
​/SaveFundSwitchService​/saveFund
Save fund switch details based on the policy number , providers are Opus and Ngin

POST
​/SaveFundSwitchSTPService​/saveFundSwitchSTP
Save fund switch systematic transfer plan details based on the policy number , data updates in Ngin or Opus System based on Policy

POST
​/ProductDetailsService​/ProductDetailsService_PS​/getProductList
List of products , provider is CRM

POST
​/FundValueService​/fetchPolicyData
Fund Value details based on the policy no, application no, insurance type and manufacturer name , providers are Ngin and Opus

POST
​/GetCashReceiptValService​/getCashReceiptValService
Cash Receipt details based on the policy no, application no , provider is Ngin(Cashiering and DC Activation) .

POST
​/SaveTopupPaymentService​/saveTopupPayment
save topup details against Policy , provider is Opus and Ngin

POST
​/SimultaneousCaseDetailsService​/getSimultaneousCaseDetails
retrieves simultaneous case details against application no,policy details

POST
​/OTPVerificationService​/fetchAgentData
retrieves agent profile details. Provider is Opus

POST
​/PremiumRedirectionService​/getPremiumRedirection
retrieves fund aportionment of policies , providers are Ngin and Opus

POST
​/MandateTransDetailService​/getMandateTransDetails
Used to get the mandate Transaction details , providers are CRM and Ngin

POST
​/PolicyListService​/getPolicyList
Retrieves payment url for policies eligible for renewal against user id . Providers are CRM and Ngin

POST
​/PreviousRequestSTPService​/getPreviousRequestSTP
This service is used to display the details of fund transfer for a given policy. The details of funds from which the transfer occurs is displayed in the transfer_from list, and the funds to which the transfer is made is displayed in the transfer_to_list.

POST
​/validateBalicPolicy​/rest​/ValidatePolicyNo​/
Used to validate the policy, providers are Opus and Ngin

POST
​/ReverseUpdateStatusLeadsService​/reverseUpdateStatusLeads
Used to get the item key , provider is CRM

POST
​/UWCOPUSService​/getUWCDetails
used for underwriting computation , provider is Ngin

POST
​/SaveNomineeDetailsService​/saveNomineeDetails
saves the details of nominee for a given policy number , provider is CRM

POST
​/ValidateAadhaarService​/validateAadhaar
validate details for a given aadhaar number , provider is Opus

POST
​/UpdateAadharService​/UpdateAadhar
updates Aadhar details for given number of policy , detail updates in Ngin and Opus based on given Policy

POST
​/KYCDetailsService​/getKYCDetails
generates KYC details for given number of policy , providers are Ngin and Opus

POST
​/SavePayoutDetailsService​/savePayoutDetails
saves payout details for given number of policy , providers are Ngin and CRM

POST
​/ReverseUpdateStatusService​/getReverseStatus
provides reverse SR status , provider is CRM

POST
​/SRStatusService​/getSRStatus
This API gives the current status of SR , provider is CRM

POST
​/ExtendRevivalService​/getExtendRevival
generates message for extending revival period for given policy , provider is CRM and Ngin

POST
​/getInstallmentAmount​/InstallmentAmount
Get Installment amount , provider is Opus

POST
​/getTotalNetAmount​/TotNetAmount
Get total net amount , provider is Opus

POST
​/DocumentEligilbleService​/getDocumentEligible
Get Document Eligible , provider is GBO (Goal Based Orientation) platform

POST
​/CurrentPayableByAgentCodeService​/getCurrentPayableByAgentCode
get the amount payable by the respective agent code , providers are ICSM portal and Ngin

POST
​/updateApprovalStatus​/updateApproval
Updates application Approval Status for PLVC, detail updates in Ngin and Opus

POST
​/SearchFundRequestService​/getSearchFundRequest
To get service request details of funds , provider is ICSM

POST
​/PreLoginContactUpdateService​/PreLoginContactUpdate
To get the Contact Updated in Ngin and Opus sysems for given application number

POST
​/AllocationDetailsService​/getAllocationDetails
To retrieve policy fund allocation details, providers are Ngin and Opus

POST
​/ActivityLogService​/getActivityLog
.Get Activity Details against policy(address change, contact update etc) , providers are Ngin and CRM

POST
​/UploadCOEService​/uploadCOE
This API is used to upload Certificate of Existence , updates in CRM

POST
​/CancelSTPRequestService​/cancelSTPRequest
Customer portals use this API to discontinue STP fund tranfer for given policy number , updates in Ngin and CRM system

POST
​/FamilyDetailsService​/getFamilyDetailsByPolicy
In this service we will get Details of Family members , providers are Ngin and CRM

POST
​/PartnerDetailService​/getPartnerDetails
In this service we will get Details of Partner , providers are Ngin and Opus

POST
​/PolicyEligibilityDetailsService​/getPolicyEligibilityDetails
Retrieves details of surrender value, partial withdrwal etc , providers are Ngin and CRM

POST
​/FTAppDetailService​/getFTApplnRequestDetails
Mashup details for Fund transfer application details , providers are Ngin and Opus

POST
​/PremiumDetailsService​/getPremiumDetails
API provides premium payment detals based on Policy , providers are Ngin and Opus

POST
​/STPDetailsService​/saveSTPDetails
Saves information for SRs like portfolio switch, fund switch name change for given policy , detail updates in Ngin or Opus

POST
​/ProfileInfoService​/getProfileInfo
fetches profile details on the basis of either profile number or mobile number from PMAC

POST
​/CustomerRetentionFutureFundDetailsService​/customerRetentionFutureFundDetails
This API is used to fetch the customer retention future fund details.

POST
​/CPIDMergeService​/cpidMerge
This service is used to merge/change the cp_id of policy

POST
​/SubmitMedicalClaimService​/submitmedicaldetails
This API is used to submit medical claim details.

POST
​/CPMergingStatusService​/getCPMergingStatus
getCPMergingStatus , provider is OPUS

POST
​/GetMedicalClaimdetailsService​/GetMedicalClaimdetails
GetMedicalClaimdetails , provider is OPUS

New Business
New Business takes care of onboarding new Customer policy applications creation,kyc validation , product selection , proposal creation using insTAB ,which is part of digital journey that provides business functionalites for Partner and Balic sales team ,Similarity STP web portal is partner specific(institutional Business).


CRT
API which are part of Customer Retention Team module



POST
​/RenewalIntimationDetailsService​/getRenewalIntimationDetails
get RenwalInitmation Details , provider is CRM

POST
​/AutoMandateRegService​/getAutoMandateReg
details about auto mandate registration for given policy , providers are CRM and Ngin

POST
​/RegisterMandateWithGroupIDService​/registerMandateWithGroupIDService
mandate registration with group ID , providers are Ngin(CRT) and Opus

POST
​/MandateRegDetailsService​/MandateRegDetailsService_PS​/regDtls
Mandate Reg details based on the policy number , providers are Ngin Opus and CRM

POST
​/RenewalDetailsService​/getRenewalDetails
Renewal info details based on the policy number , providers are CRM and Ngin

POST
​/RevivalDetailsService​/getRevivalDetails
Used to get the revival details , providers are Ngin and CRM

POST
​/GroupIDForMultiplePolicyMandateRegistrationService​/getGroupIDfromMultiplePolicy
generates group id of multiple polices for Mandate Registration , providers are Ngin (CRT) and Opus

POST
​/GroupIdDetailsForMRService​/getGroupIdDetails
To get the policy details against particular group id from Opu and Ngin systems.
Performance Portal
API which are part of performace portal and integrates with Ismart , imanage and Sales portal to provide Agent related information, and their performance metrics.



POST
​/CustomerPortfolioSearchByName​/SearchCustomerPortfolio
searches customer portfolio on basis customer name for an agent , providers and Opus and Ngin

POST
​/SearchCustomerService​/searchCustomer
retrieves customer details on basis of mobile/dob/name , provider is Opus

POST
​/CustomerPortfolioWithChartService​/getSelectedCustomerPortfolioWithChart
provide customer portfolio list for the selected customer for a given policy number , providers are Ngin and Opus (ICSM)

POST
​/PolicySnapshotWithChartService​/getPolicySnapshotWithChart
To get Customer Fund Value Details , provider is Opus(ICSM)

POST
​/CustomerPortfolioSearchByDateService​/customerPortfolioSearch
searches customer portfolio on basis customer date for an agent , provider is ICSM
Cashiering and Receipting
APIs realted to receipt details.


claims
API for claims processings and retrieving



POST
​/ClaimsPartnerService​/getPartnerDetails
Claims Partner Details, provider is NGIN and Opus

POST
​/ClaimPayoutDetailsService​/getClaimPayoutDetails
get Claim Payout Details , providers are NGIN and CRM legacy system

POST
​/ClaimRepudiationDetailsService​/getClaimRepudiationDetails
generates claim repudiation details for given number of policy , providers are Ngin (clamis) and CRM

POST
​/ClaimPolicyListService​/getClaimPolicyList
Get claim policy list , provider is Opus

POST
​/ClaimNotificationService​/getClaimNotification
get Claim Notification
Common
API which serves functionality used by multiple features



POST
​/LOVDetailService​/getLOV
get LOV Details , provider is Common Master


POST
​/RoutingInfoService​/{RoutingOperation}
getRoutingInfo - target System

POST
​/CPIDFromUserId​/getCPIDFromUserId
CPID based on the user ID

POST
​/GenerateAadharOTPService​/generateOTP
Sends OTP on the registered mobile linked with aadhar Number. Provider is aadhar Voult/Opus

POST
​/ValidateAadhaarOTPService​/ValidateAadhaarOTP
Validates Aadhar with OTP generated from AadharOTP API and provides Reference Number for aadhar. Additional features of name match with PAS system is added . Providers are aadhar Voult , Opus and Ngin
Sentimeter
Sentimeter COV details & sentimeter influencer details



POST
​/SentimeterCOVDetailsService​/getCOVDetails
This Service is used to get the sentimeter COV details

POST
​/SentimeterInfluencerDetailsService​/getInfluencerDetails
This Service is used to get the sentimeter influencer details
Websales


POST
​/GetAccountStatementService​/GetAccountStatement_PS​/getAccountStatement
This API is used to generate the account statement .

POST
​/BalicWSRest​/BalicWSRest
saves application details in NB Journey for partner services, updates the data in Opus
SFDC


POST
​/MasterPolicyInsuranceService​/getInsuranceDetails
This service is used to save the master policy insurance data

POST
​/ReceiptCreationService​/generateReceipt
generate receipt number , provider is OPUS.

POST
​/CCMPushDataService​/saveCCMData
save CCM Data in DWH , provider is OPUS.

POST
​/PartnerIdCreationService​/getPartnerIdCreation
This service is used to create the partner id for the customer based on the details provided.

POST
​/QuotationIdCreationService​/getQuotationIdCreation
This service is used to create the quotation id for the customer based on the details provided.
external_api


POST
​/GetApplicationTrackerDetailsService​/getApplicationTrackerDetails
This API is used to track APplication details for a given Application/Policy Number

POST
​/GetUserIdFlagService​/getUserIdFlag
This API is used to get user_flag for a given login Id.Login Id can be a mobile number / email, user Id.

POST
​/TrackerDocumentService​/getDocument
This API is used to send list of documents to user via mail or SMS, Providers are OPUS, NGIN & CCM

POST
​/PinCodeDetailService​/getPincodeDetails
Get Pincode Details, provider is CRM,CP,GBO.

POST
​/CreateMerchantAndSendNotificationService​/createMerchantAndSendNotification
Create Merchant for PIVC journey. Provider is Signzy

POST
​/ProfileDetailsService​/getProfileDetails
Get Profile Details , providers are OPUS and NGIN

POST
​/PolicyDetailsDashboardService​/getPolicyDetailsDashboard
Get Policy Details of a given User Id , providers are OPUS and NGIN

POST
​/getMicrDetails​/MicrDetails
Get MICR Details , Provider is Opus

POST
​/GetDocumentService​/getDocument
API provides enrycpted pdf statements , providers are PAS systems and CCM

POST
​/DownloadLatestReceiptDetailsService​/downloadReceipt
provides Receipting statement in pdf , provides are Opus , cashiering and CCM


POST
​/getReceiptDetails​/getReceiptDetails
get Receipt Details , providers are Opus and Ngin

POST
​/CustomerDetailsService​/fetchCustomerDetails
retrieves customer details from Opus system for given Policy and Search Criteria

POST
​/WhatsappPolicyChartService​/getWhatsappPolicyChart
get Whatsapp Policy Chart Details , provider is PAS system

POST
​/GenerateReceiptService​/generateReceiptNumber
generates receipt on basis of payment transaction details for given application or policy number , provider is Opus

POST
​/RenewalPolicyDetailService​/getRenewalPolicyDetailService
Used to retrieve policy renewal premium details , providers are cashiering and Opus

POST
​/GetRenewalDetailsService​/getRenewalDetails
retrieves policy renewal details against agent , provider is Opus

POST
​/UpdatePANService​/UpdatePAN
Used to update the PAN details , systems involved are Opus , Ngin and CRM

POST
​/PANValidationService​/ValidatePAN
Used to validate the PAN details , provider is Opus

POST
​/ApplicationSearchService​/searchApplication
Used to get the application details , providers are Ngin and Opus

POST
​/CustPaymentLinkService​/getCustPaymentLink
get CustPaymentLink for a given policy number , providers are Opus , Ngin and CRM

POST
​/UpdateContactDetailsService​/UpdateContact
updates contact details for given policy , updates in Ngin or Opus

POST
​/BankDetailsService​/getIFSCDetails
Get Bank Details , providers are Opus Ngin and CRM

POST
​/PremiumCollectionDetailsService​/getPremiumCollectionDetails
retrieves premium collection details for renewal, providers are OPUS and NGIN(Cashiering n billing)

POST
​/PIVCApplication​/SubmitProductQestionnaire
API accepts the Questionare inforamtion submitted by Merchant in Digital PIVC journey, data updates in PAS system based on application Number
Mandate Registration


POST
​/AutomandateLinkService​/sendAutomandateLink
This API is used to sendAutomandateLink

POST
​/GroupIdDetailsForUniversalMandate​/getGroupId
This API is used to get Group Id

POST
​/RegisterAutoMandateService​/registerAutoMandate
This API is used to register Automandate

POST
​/MandateRegistrationService​/getMandateRegister
get Mandate Register

POST
​/GenerateAutoModeReceiptService​/generateAutoModeReceipt
generate AutoModeReceipt , provider is cashiering and Opus




Finance


POST
​/StopPaymentReissueService​/saveStopPaymentReissue
stop Payment Reissue , data saves in SPRP system

POST
​/PaymentAvenuesService​/getPaymentAvenues
Provides branchname based on pincode, city , provider is CRM

POST
​/PaymentReceiptService​/getPaymentReceipt
generates payment receipt for given policy , providers are Opus and Ngin (cashiering)

POST
​/getTransactionDetail​/getTransaction
Get Transaction Details

POST
​/PaymentReissueDetailsService​/getPaymentReissueDetails
Mashup API , API provides to get cheque details from payout module based on policy number. Providers are Ngin , CRM and SPRP

IMPORTANT:

- Use markdown headings (##, ###)
- **CRITICAL**: For sections 2.1 and 2.2, if structured tables exist in source, reproduce them as markdown tables
- **For Purpose and To-be process: PRIORITIZE "PART B" and "PART C (Mandatory) Detailed Requirement" content**
- Extract content based on what's ACTUALLY across ALL processed sheets, regardless of domain
- Adapt language and focus to match the source content type
- If no content found for a subsection after checking ALL sheets, leave it blank

VALIDATION CHECK:

Before finalizing each section, verify that every piece of information can be traced back to the source requirements from the processed Excel sheets (excluding Test Scenarios). For Purpose and To-be process sections, ensure you've prioritized "PART B : (Mandatory) Detailed Requirement" content when available.

OUTPUT FORMAT:
Provide ONLY the markdown sections (## 1.0 Introduction, ### 1.1 Purpose, etc.) with the extracted content. Do not include any of these instructions, validation checks, or processing guidelines in your response.

""",

    "process_requirements": """

You are a Business Analyst expert creating sections 3.0–4.0 of a comprehensive BRD.

PREVIOUS CONTENT:

{previous_content}

SOURCE REQUIREMENTS:

{requirements}

EXCEL FILE PROCESSING INSTRUCTIONS:

**FOR EXCEL FILES (.xlsx/.xls):**
- Process ALL sheets EXCEPT "Test Scenarios" sheet
- Include data from sheets: "Requirement", "Ops Risk Assessment", and any other available sheets
- Extract workflow, process, and business rule information from ALL relevant columns and rows
- Look for step-by-step processes, business rules, and functional requirements across all sheets

CRITICAL INSTRUCTIONS:

- Extract information from ALL available sheets (except Test Scenarios sheet)
- Identify ACTUAL workflows, processes, and business rules from ANY relevant sheet
- Focus on step-by-step logic, conditions, and decision points mentioned across ALL processed sheets
- Adapt to any business domain (training, validation, integration, access control, etc.)
- Do NOT create, assume, or fabricate any content not explicitly present in the source

Create ONLY the following sections with detailed content in markdown:

## 3.0 Process / Data Flow diagram / Figma

Extract DETAILED workflow/process information from ALL processed sheets:

### 3.1 Workflow Description

Create step-by-step process based on what's described across ALL processed sheets:
- What triggers the process or workflow?
- What are the sequential steps or stages?
- What decision points, conditions, or validations occur?
- What are the different paths or outcomes?
- How are errors, exceptions, or edge cases handled?
- What user interactions or system responses are involved?

Format as logical flow:
- Step 1: [Action/Trigger from any relevant sheet]
  - If [condition mentioned in any sheet]: [result/next step]
  - If [alternative condition]: [alternative result]
- Step 2: [Next Action from any relevant sheet]
  - [Continue based on source content from processed sheets]

Look for process indicators across ALL sheets: "workflow", "process", "steps", "sequence", "flow", "journey", "condition", "if", "then", "when", "trigger", "action", "response"

## 4.0 Business / System Requirement

### 4.1 System Requirements

Extract BUSINESS functional requirements from ALL processed sheets:
- Look for detailed requirement sections in the "Requirement" sheet primarily
- Also check other sheets for additional functional requirements
- Extract information from any columns containing requirement descriptions
- Include business rules, validation requirements, and functional specifications

### 4.2 Application / Module Name: [Extract exact application/module name from ANY processed sheet]

Create detailed requirement table based on content from ALL processed sheets:

| **Rule ID** | **Rule Description** | **Expected Result** | **Dependency** |
|-------------|---------------------|-------------------|----------------|
| **4.1.1** | [Extract specific business rule from ANY processed sheet] | [Exact expected behavior mentioned in ANY sheet] | [Technical/system dependencies noted in ANY sheet] |

Focus on extracting from ALL processed sheets:
- Specific business rules, validations, or logic mentioned
- Functional requirements and expected system behaviors
- User access controls, permissions, or restrictions
- Data validation, processing, or transformation rules
- Integration requirements and system interactions

**SPECIFIC FOR EXCEL:** 
- If there's a "Requirement" sheet, prioritize extracting from detailed requirement sections
- Check for cells like "Detailed Requirement", "Business Rule", "Functional Spec", etc.
- Process other sheets for supplementary functional requirements

IMPORTANT:

- Use markdown headings
- Create detailed requirement tables with multiple columns
- Base all content on what's explicitly stated across ALL processed sheets
- Adapt terminology and focus to match the source domain
- Leave blank if no content found after checking ALL relevant sheets

VALIDATION CHECK:

Before finalizing each section, verify that every piece of information can be traced back to the source requirements from the processed Excel sheets (excluding Test Scenarios). Remove any content that cannot be directly attributed to the source documents.

OUTPUT FORMAT:
Provide ONLY the markdown sections (## 3.0, ### 3.1, etc.) with the extracted content. Do not include any of these instructions, validation checks, or processing guidelines in your response.

""",

    "data_communication": """

You are a Business Analyst expert creating section 5.0 of a comprehensive BRD.

PREVIOUS CONTENT:

{previous_content}

SOURCE REQUIREMENTS:

{requirements}

EXCEL FILE PROCESSING INSTRUCTIONS:

**FOR EXCEL FILES (.xlsx/.xls):**
- Process ALL sheets EXCEPT "Test Scenarios" sheet
- Include data from sheets: "Requirement", "Ops Risk Assessment", and any other available sheets
- Extract data requirements, specifications, and communication needs from ALL relevant sheets
- Look for data-related requirements across all processed sheets

CRITICAL INSTRUCTIONS:

- Extract information from ALL available sheets (except Test Scenarios sheet)
- Identify ACTUAL data and communication needs from ANY relevant sheet
- Adapt to any type of data requirements (user data, transaction data, training data, etc.)
- Do NOT create, assume, or fabricate any content not explicitly present in the source

Create ONLY the following section with detailed content in markdown:

## 5.0 MIS / DATA Requirement

### 5.1 Data Specifications

Extract SPECIFIC data requirements from ALL processed sheets:
- What data elements, fields, or attributes are needed?
- What data sources, databases, or systems provide this data?
- What data formats, structures, or schemas are required?
- What data validation, quality, or integrity requirements exist?
- What data processing, transformation, or calculation needs are mentioned?
- Any data retention, archival, or lifecycle requirements

Look across ALL sheets for data-related content in any columns or sections.

### 5.2 Reporting and Analytics Needs

Extract any mentioned across ALL processed sheets:
- What reports, dashboards, or analytics are required?
- What metrics, KPIs, or measurements need to be tracked?
- What data visualization or presentation requirements exist?
- What frequency or scheduling of reports is needed?
- What user roles or audiences need access to reports?
- Any real-time monitoring or alerting requirements
- Which plots/charts are applicable to be built?

Check all processed sheets for reporting requirements, analytics needs, or dashboard specifications.

### 5.3 Data Sources and Destinations

Extract from ALL processed sheets:
- Source systems, databases, or applications providing data
- Target systems, repositories, or destinations for data
- Integration points, APIs, or data exchange mechanisms
- Data flow directions and transformation requirements
- External systems, third-party sources, or partner integrations
- Master data management or reference data needs

Search across ALL sheets for system names, database references, integration points, and data flow information.

## 6.0 Communication Requirement

**EMAIL CONTENT EXTRACTION:**
- Look for email communication patterns in the source requirements including:
  - Email addresses (containing @ symbol)
  - Email headers like "From:", "To:", "Subject:", "Date:"
  - Email greetings like "Hi [Name]", "Hello [Name]", "Dear [Name]"
  - Email signatures with names, titles, phone numbers
  - Reply chains and conversation threads
  - Email closings like "Regards", "Thanks", "Best Regards"
  - Corporate email signatures with company names

**SEARCH FOR EMAIL PATTERNS:**
- @ symbols indicating email addresses
- Phone numbers in signatures
- Corporate titles and designations
- Email thread conversations
- Any communication that looks like email correspondence

**IF email-like content IS FOUND:**
- Extract and include all email communications found in the source
- Preserve the conversation flow and chronological order
- Include email addresses, names, and contact information
- Maintain the original format and structure
- Format as:

### Email Communication Thread

[Extract and preserve the complete email conversation as it appears in the source, maintaining the original format, names, email addresses, signatures, and conversation flow]

**IF NO email-like content is found:**
- State: "No communication requirement content found in source documents"
- Do NOT generate, create, or simulate any email content

**CRITICAL:** Only extract actual email-like content that exists in the source requirements. Never generate, create, assume, or fabricate any email content.

IMPORTANT:

- Use markdown headings
- Preserve any tables in markdown format from ANY processed sheet
- Extract content based on what's ACTUALLY across ALL processed sheets, regardless of domain
- Adapt language and focus to match the source content type
- If no content found for a subsection after checking ALL sheets, leave it blank

VALIDATION CHECK:

Before finalizing each section, verify that every piece of information can be traced back to the source requirements from the processed Excel sheets (excluding Test Scenarios). Remove any content that cannot be directly attributed to the source documents.

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

| **Test ID** | **Test Scenario Name** | **Objective** | **Test Steps** | **Expected Results** | **Type** |
|-------------|---------------|---------------|----------------|---------------------|----------|
| [Extract ID] | [Extract Name] | [Extract Objective] | [Extract Steps] | [Extract Results] | [Extract Type] |

Also, ADDING on to this, generate test scenarios based EXCLUSIVELY on functionality explicitly described in source requirements

**STOP HERE - Do not proceed to Secondary Approach if existing tests are found**

---

**SECONDARY APPROACH - Generate from Functional Requirements:**
**ONLY EXECUTE IF PRIMARY APPROACH YIELDS NO RESULTS**

IF NO existing test scenarios are found in ANY source documents, THEN generate test scenarios based EXCLUSIVELY on functionality explicitly described in source requirements:

| **Test ID** | **Test Name** | **Objective** | **Test Steps** | **Expected Results** | **Type** |
|-------------|---------------|---------------|----------------|---------------------|----------|

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

**CRITICAL EXTRACTION RULE:** 
Extract the EXACT table content from the source documents. Do NOT modify, interpret, or reformat the content.

**SEARCH FOR RISK CONTENT:**
- Look for any sheet named "Ops Risk Assessment", "Risk Evaluation", "Risk Assessment", or similar
- Look for any table or structured data containing risk-related information
- Search for keywords: "risk", "evaluation", "assessment", "impact", "controls"

**EXTRACTION PROCESS:**
1. **IF a risk table/data is found in the source:**
   - Copy the EXACT column headers from the source
   - Copy the EXACT row data from the source
   - Maintain the EXACT table structure and content
   - Convert to clean markdown table format WITHOUT changing any text content
   - Include ALL rows and columns as they appear in the source

2. **EXAMPLE - If source has this table:**
   ```
   | Risk Type | Impact | Mitigation | Status |
   | High Risk | Operational | Control A | Active |
   ```
   
   **OUTPUT EXACTLY:**
   ```
   | Risk Type | Impact | Mitigation | Status |
   |-----------|--------|------------|---------|
   | High Risk | Operational | Control A | Active |
   ```

3. **IF NO risk content found:**
   - State: "No Risk Evaluation content found in source documents"
   - Do NOT create any template or sample content

**FORBIDDEN:**
- Do NOT generate placeholder text like "List down the business risks"
- Do NOT create template structures 
- Do NOT interpret or modify the source content
- Do NOT add explanatory text or instructions in table cells

**VALIDATION:**
Every piece of content in this section must be traceable to the exact source document content.

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
    
    st.write("="*120)
    st.write("📋 COMBINED REQUIREMENTS SENT TO LLM:")
    st.write("="*120)
    
    with st.expander("📄 View Complete Requirements Content", expanded=False):
        st.text_area("Full Content", combined_requirements, height=400)
    
    st.write(f"📊 **Content Statistics:**")
    st.write(f"- Total characters: {len(combined_requirements):,}")
    lines_count = len(combined_requirements.split('\n'))
    words_count = len(combined_requirements.split())
    st.write(f"- Total lines: {lines_count:,}")
    st.write(f"- Total words (approx): {words_count:,}")
    st.write(f"- Number of chunks: {len(req_chunks)}")
    
    st.write(f"📖 **Content Preview (First 2000 characters):**")
    st.code(combined_requirements[:2000] + "..." if len(combined_requirements) > 2000 else combined_requirements)
    
    sections = [line for line in combined_requirements.split('\n') if line.strip().startswith('===')]
    if sections:
        st.write(f"**Document Structure:**")
        for section in sections[:10]:
            st.write(f"- {section.strip()}")
        if len(sections) > 10:
            st.write(f"- ... and {len(sections) - 10} more sections")
    
    st.write("="*120)
    
    previous_content = ""
    final_sections = []
    
    for i, chain in enumerate(chains):
        try:
            st.write(f"\\n🔗 **PROCESSING CHAIN {i+1}/4**")
            st.write(f"{'='*60}")
            
            with st.expander(f"🔍 Chain {i+1} Details - Click to expand", expanded=False):
                
                if i == 0:
                    st.write("**Input to Chain 1 (Introduction & Impact Analysis):**")
                    st.write(f"- Requirements length: {len(combined_requirements):,} characters")
                    st.write("**Requirements Preview:**")
                    st.code(combined_requirements[:1000] + "..." if len(combined_requirements) > 1000 else combined_requirements)
                    
                    st.write("**Template Used:**")
                    st.code(SECTION_TEMPLATES["intro_impact"][:500] + "...")
                    
                    result = chain.run(requirements=combined_requirements)
                else:
                    chain_names = ["", "Process & Requirements", "Data & Communication", "Testing & Final"]
                    st.write(f"**Input to Chain {i+1} ({chain_names[i]}):**")
                    st.write(f"- Previous content length: {len(previous_content):,} characters")
                    st.write(f"- Requirements length: {len(combined_requirements):,} characters")
                    
                    st.write("**Previous Content Preview:**")
                    st.code(previous_content[:800] + "..." if len(previous_content) > 800 else previous_content)
                    
                    st.write("**Requirements Preview:**")
                    st.code(combined_requirements[:800] + "..." if len(combined_requirements) > 800 else combined_requirements)
                    
                    template_keys = ["", "process_requirements", "data_communication", "testing_final"]
                    st.write(f"**Template Used ({template_keys[i]}):**")
                    st.code(SECTION_TEMPLATES[template_keys[i]][:500] + "...")
                    
                    result = chain.run(previous_content=previous_content, requirements=combined_requirements)
                
                st.write(f"**Chain {i+1} Output:**")
                st.write(f"- Response length: {len(result):,} characters")
                result_lines = len(result.split('\n'))
                result_words = len(result.split())
                st.write(f"- Response lines: {result_lines:,}")
                st.write(f"- Response words (approx): {result_words:,}")
                
                output_sections = [line for line in result.split('\n') if line.strip().startswith('##')]
                if output_sections:
                    st.write("**Sections Generated:**")
                    for section in output_sections:
                        st.write(f"- {section.strip()}")
                
                st.write("**Response Preview:**")
                st.code(result[:1000] + "..." if len(result) > 1000 else result)
            
            print(f"\n{'='*60}")
            print(f"CHAIN {i+1} INPUT:")
            print(f"{'='*60}")
            
            if i == 0:
                print("Input to Chain 1 (intro_impact):")
                print(f"Requirements length: {len(combined_requirements)} characters")
                print("First 1000 characters of requirements:")
                print(combined_requirements[:1000] + "..." if len(combined_requirements) > 1000 else combined_requirements)
                
            else:
                print(f"Input to Chain {i+1}:")
                print(f"Previous content length: {len(previous_content)} characters")
                print(f"Requirements length: {len(combined_requirements)} characters")
                print("Previous content (first 500 chars):")
                print(previous_content[:500] + "..." if len(previous_content) > 500 else previous_content)
                print("\nRequirements (first 500 chars):")
                print(combined_requirements[:500] + "..." if len(combined_requirements) > 500 else combined_requirements)
            
            print(f"\nCHAIN {i+1} OUTPUT:")
            print(f"Response length: {len(result)} characters")
            print("First 1000 characters of response:")
            print(result[:1000] + "..." if len(result) > 1000 else result)
            print(f"{'='*60}")
            
            final_sections.append(result)
            previous_content += "\\n\\n" + result
            
            st.write(f"✅ **Completed section group {i+1}/4**")
            st.write(f"📈 **Cumulative content length: {len(previous_content):,} characters**")
            
        except Exception as e:
            print(f"ERROR in chain {i+1}: {str(e)}")
            st.error(f"❌ Error in chain {i+1}: {str(e)}")
            final_sections.append(f"## Error in section group {i+1}\\nError processing this section: {str(e)}")
    
    final_brd = "\\n\\n".join(final_sections)
    
    st.write("\\n" + "="*80)
    st.write("📋 **FINAL BRD GENERATION COMPLETE**")
    st.write("="*80)
    
    with st.expander("📊 Final BRD Statistics & Preview", expanded=True):
        st.write(f"**Final Statistics:**")
        st.write(f"- Total final BRD length: {len(final_brd):,} characters")
        final_lines = len(final_brd.split('\n'))
        final_words = len(final_brd.split())
        st.write(f"- Total lines: {final_lines:,}")
        st.write(f"- Total words (approx): {final_words:,}")
        
        # Show final sections
        final_sections_headers = [line for line in final_brd.split('\n') if line.strip().startswith('##')]
        if final_sections_headers:
            st.write(f"**Generated Sections ({len(final_sections_headers)}):**")
            for section in final_sections_headers:
                st.write(f"- {section.strip()}")
        
        st.write("**Final BRD Preview (first 2000 characters):**")
        st.code(final_brd[:2000] + "..." if len(final_brd) > 2000 else final_brd)
    
    print("\n" + "="*80)
    print("FINAL BRD CONTENT:")
    print("="*80)
    print(f"Total final BRD length: {len(final_brd)} characters")
    print("Final BRD (first 2000 characters):")
    print(final_brd[:2000] + "..." if len(final_brd) > 2000 else final_brd)
    print("="*80)
    
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
        ("    2.1 Impacted Products", "impacted_products"),
        ("    2.2 Applications Impacted", "applications_impacted"), 
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
    def clean_table_cell_value(cell_text):
        if cell_text is None:
            return "-"
        
        str_val = str(cell_text).strip()
        
        if str_val.lower() == 'nan':
            return "-"
        
        if str_val.startswith("Unnamed"):
            return "Insert Column Name"
        
        return str_val
    
    if not table_data or len(table_data) < 2:
        return None
    
    table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
    table.style = 'Table Grid'
    
    for i, cell_text in enumerate(table_data[0]):
        cell = table.rows[0].cells[i]
        cleaned_text = clean_table_cell_value(cell_text)
        cell.text = cleaned_text
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
    
    for row_idx, row_data in enumerate(table_data[1:], 1):
        for col_idx, cell_text in enumerate(row_data):
            if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                cleaned_text = clean_table_cell_value(cell_text)
                table.rows[row_idx].cells[col_idx].text = cleaned_text
    
    return table

def parse_markdown_table(table_text):
    def clean_cell_value(cell_text):
        if cell_text is None:
            return "-"
        
        str_val = str(cell_text).strip()
        
        if str_val.lower() == 'nan':
            return "-"
        
        if str_val.startswith("Unnamed"):
            return "Insert Column Name"
        
        return str_val
    
    lines = [line.strip() for line in table_text.split('\n') if line.strip()]
    
    if len(lines) < 2:
        return None
    
    if len(lines) >= 2 and '---' in lines[1]:
        lines.pop(1)
    
    table_data = []
    max_cols = 0
    
    for line in lines:
        if line.startswith('|') and line.endswith('|'):
            line = line[1:-1]
        
        cells = [clean_cell_value(cell.strip()) for cell in line.split('|')]
        
        while cells and not cells[-1]:
            cells.pop()
        
        if cells:
            table_data.append(cells)
            max_cols = max(max_cols, len(cells))
    
    if not table_data:
        return None
    

    normalized_data = []
    for row in table_data:
        normalized_row = row + ['-'] * (max_cols - len(row))
        normalized_data.append(normalized_row)
    
    while max_cols > 1:
        last_col_has_data = False
        for row in normalized_data:
            if len(row) >= max_cols and row[max_cols-1].strip() and row[max_cols-1] != '-':
                last_col_has_data = True
                break
        
        if last_col_has_data:
            break
        else:
            for row in normalized_data:
                if len(row) >= max_cols:
                    row.pop()
            max_cols -= 1
    
    final_data = []
    for row in normalized_data:
        final_row = row[:max_cols] + ['-'] * max(0, max_cols - len(row))
        final_data.append(final_row)
    
    return final_data if final_data else None

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

def extract_content_from_excel(excel_file, max_rows_per_sheet=70, max_sample_rows=10, visible_only=True):
    def clean_cell_value(cell_text):
        if cell_text is None:
            return "-"
        
        str_val = str(cell_text).strip()
        
        if str_val.lower() == 'nan':
            return "-"
        
        if str_val.startswith("Unnamed"):
            return "Insert Column Name"
        
        return str_val
    
    def extract_horizontal_table(df, start_row_idx, start_col_idx, table_identifier):
        """Extract horizontal table structure starting from a given position"""
        table_data = {
            "table_type": table_identifier,
            "headers": [],
            "data_rows": [],
            "raw_structure": []
        }
        
        try:
            # Look for table headers in the row containing the identifier
            # and subsequent rows
            current_row = start_row_idx
            max_search_rows = min(start_row_idx + 10, len(df))
            
            found_headers = False
            header_row_idx = None
            
            # Search for header row (look for product names or application names)
            for search_row in range(current_row, max_search_rows):
                if search_row >= len(df):
                    break
                    
                row_data = df.iloc[search_row].values
                
                # Check if this row contains table headers
                non_empty_cells = [str(cell).strip() for cell in row_data if pd.notna(cell) and str(cell).strip()]
                
                # Look for product-related headers
                product_indicators = ['ULIP', 'Term', 'Endowment', 'Annuity', 'Health', 'Group', 'All']
                app_indicators = ['OPUS', 'INSTAB', 'NGIN', 'PMAC', 'CRM', 'Cashier', 'Other']
                
                if any(indicator in ' '.join(non_empty_cells).upper() for indicator in product_indicators + app_indicators):
                    header_row_idx = search_row
                    found_headers = True
                    break
            
            if found_headers and header_row_idx is not None:
                # Extract headers
                header_row = df.iloc[header_row_idx]
                headers = []
                header_positions = []
                
                for col_idx, cell_value in enumerate(header_row):
                    if pd.notna(cell_value) and str(cell_value).strip():
                        clean_header = clean_cell_value(cell_value)
                        if clean_header != "-" and not clean_header.startswith("Insert"):
                            headers.append(clean_header)
                            header_positions.append(col_idx)
                
                table_data["headers"] = headers
                
                # Extract data rows (look in next few rows after headers)
                data_start_row = header_row_idx + 1
                max_data_rows = min(data_start_row + 5, len(df))
                
                for data_row_idx in range(data_start_row, max_data_rows):
                    if data_row_idx >= len(df):
                        break
                        
                    data_row = df.iloc[data_row_idx]
                    
                    # Check if row has meaningful data
                    row_values = []
                    has_data = False
                    
                    # Extract values corresponding to header positions
                    for pos in header_positions:
                        if pos < len(data_row):
                            cell_val = clean_cell_value(data_row.iloc[pos])
                            row_values.append(cell_val)
                            if cell_val not in ["-", ""]:
                                has_data = True
                        else:
                            row_values.append("-")
                    
                    if has_data:
                        # Create row description
                        row_description = data_row.iloc[0] if pd.notna(data_row.iloc[0]) else f"Row {data_row_idx + 1}"
                        
                        row_data = {
                            "row_description": clean_cell_value(row_description),
                            "values": dict(zip(headers, row_values))
                        }
                        table_data["data_rows"].append(row_data)
                
                # Create raw structure representation
                if headers and table_data["data_rows"]:
                    table_data["raw_structure"] = {
                        "markdown_table": create_markdown_table(headers, table_data["data_rows"]),
                        "structured_data": table_data["data_rows"]
                    }
            
        except Exception as e:
            table_data["error"] = str(e)
        
        return table_data
    
    def create_markdown_table(headers, data_rows):
        """Create a markdown table representation"""
        if not headers or not data_rows:
            return ""
        
        # Create header row
        header_line = "| " + " | ".join(headers) + " |"
        separator_line = "|" + "|".join([" --- " for _ in headers]) + "|"
        
        # Create data rows
        data_lines = []
        for row in data_rows:
            values = [row["values"].get(header, "-") for header in headers]
            data_line = "| " + " | ".join(values) + " |"
            data_lines.append(data_line)
        
        return "\n".join([header_line, separator_line] + data_lines)
    
    # Initialize JSON structure (same as before)
    result = {
        "metadata": {
            "total_sheets": 0,
            "processing_status": "success",
            "visible_only": visible_only,
            "max_rows_per_sheet": max_rows_per_sheet,
            "max_sample_rows": max_sample_rows
        },
        "priority_content": {
            "part_b": [],
            "part_c": []
        },
        "sheets": [],
        "summary": {
            "part_b_found": False,
            "part_c_found": False,
            "total_rows_processed": 0,
            "total_columns_processed": 0,
            "detailed_requirements_found": False
        }
    }
    
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
                result["metadata"]["processing_status"] = "error"
                result["metadata"]["error"] = "No visible sheets found in the Excel file"
                return json.dumps(result, indent=2)
            
            if not isinstance(excel_data, dict):
                excel_data = {visible_sheets[0]: excel_data}
        else:
            excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        result["metadata"]["total_sheets"] = len(excel_data)
        
        # MODIFIED PART C EXTRACTION - Now supports horizontal tables
        for sheet_name, df in excel_data.items():
            if df.empty:
                continue
                
            for col_idx, col in enumerate(df.columns):
                for row_idx, cell_value in enumerate(df[col]):
                    cell_str = str(cell_value).strip()
                    part_c_patterns = [
                        "PART C : (Mandatory) Detailed Requirement",
                        "PART C (Mandatory) Detailed Requirement", 
                        "PART C : Mandatory Detailed Requirement",
                        "PART C Mandatory Detailed Requirement",
                        "PART C : Detailed Requirement",
                        "PART C Detailed Requirement",
                        "PART C:",
                        "Part C :",
                        "Part C:",
                        "PART C"
                    ]
                    
                    for pattern in part_c_patterns:
                        if pattern.lower() in cell_str.lower():
                            part_c_entry = {
                                "sheet_name": sheet_name,
                                "column": col,
                                "row": row_idx + 2,
                                "header": cell_str,
                                "content": [],
                                "adjacent_content": [],
                                "horizontal_tables": []  # NEW: Store horizontal tables
                            }
                            
                            # Original vertical extraction (keep for compatibility)
                            for next_row in range(row_idx + 1, min(row_idx + 10, len(df))):
                                if next_row < len(df):
                                    next_cell = df.iloc[next_row][col]
                                    if pd.notna(next_cell) and str(next_cell).strip():
                                        part_c_entry["content"].append({
                                            "row": next_row + 2,
                                            "text": str(next_cell).strip()
                                        })
                            
                            # Original adjacent content extraction (keep for compatibility)
                            for adj_col_offset in [-1, 1]:
                                adj_col_index = col_idx + adj_col_offset
                                if 0 <= adj_col_index < len(df.columns):
                                    adj_col = df.columns[adj_col_index]
                                    for adj_row in range(max(0, row_idx-2), min(row_idx + 8, len(df))):
                                        adj_cell = df.iloc[adj_row][adj_col]
                                        if pd.notna(adj_cell) and str(adj_cell).strip():
                                            part_c_entry["adjacent_content"].append({
                                                "column": adj_col,
                                                "row": adj_row + 2,
                                                "text": str(adj_cell).strip()
                                            })
                            
                            # NEW: Extract horizontal tables
                            # Look for "Products Impacted" table
                            for search_row in range(row_idx + 1, min(row_idx + 15, len(df))):
                                if search_row < len(df):
                                    search_cell = df.iloc[search_row][col]
                                    if pd.notna(search_cell) and "Products Impacted" in str(search_cell):
                                        products_table = extract_horizontal_table(df, search_row, col_idx, "Products Impacted")
                                        if products_table["headers"]:
                                            part_c_entry["horizontal_tables"].append(products_table)
                                        break
                            
                            # Look for "Applications Impacted" table
                            for search_row in range(row_idx + 1, min(row_idx + 20, len(df))):
                                if search_row < len(df):
                                    search_cell = df.iloc[search_row][col]
                                    if pd.notna(search_cell) and "Applications Impacted" in str(search_cell):
                                        apps_table = extract_horizontal_table(df, search_row, col_idx, "Applications Impacted")
                                        if apps_table["headers"]:
                                            part_c_entry["horizontal_tables"].append(apps_table)
                                        break
                            
                            result["priority_content"]["part_c"].append(part_c_entry)
                            result["summary"]["part_c_found"] = True
                            break
        
        # Extract Part B content (same logic, can be enhanced similarly)
        for sheet_name, df in excel_data.items():
            if df.empty:
                continue
                
            for col_idx, col in enumerate(df.columns):
                for row_idx, cell_value in enumerate(df[col]):
                    cell_str = str(cell_value).strip()
                    part_b_patterns = [
                        "PART B : (Mandatory) Detailed Requirement",
                        "PART B (Mandatory) Detailed Requirement", 
                        "PART B : Mandatory Detailed Requirement",
                        "PART B Mandatory Detailed Requirement",
                        "PART B : Detailed Requirement",
                        "PART B Detailed Requirement",
                        "PART B:",
                        "Part B :",
                        "Part B:",
                        "PART B"
                    ]
                    
                    for pattern in part_b_patterns:
                        if pattern.lower() in cell_str.lower():
                            part_b_entry = {
                                "sheet_name": sheet_name,
                                "column": col,
                                "row": row_idx + 2,
                                "header": cell_str,
                                "content": [],
                                "adjacent_content": []
                            }
                            
                            # Extract next rows content
                            for next_row in range(row_idx + 1, min(row_idx + 10, len(df))):
                                if next_row < len(df):
                                    next_cell = df.iloc[next_row][col]
                                    if pd.notna(next_cell) and str(next_cell).strip():
                                        part_b_entry["content"].append({
                                            "row": next_row + 2,
                                            "text": str(next_cell).strip()
                                        })
                            
                            # Extract adjacent column content
                            for adj_col_offset in [-1, 1]:
                                adj_col_index = col_idx + adj_col_offset
                                if 0 <= adj_col_index < len(df.columns):
                                    adj_col = df.columns[adj_col_index]
                                    for adj_row in range(max(0, row_idx-2), min(row_idx + 8, len(df))):
                                        adj_cell = df.iloc[adj_row][adj_col]
                                        if pd.notna(adj_cell) and str(adj_cell).strip():
                                            part_b_entry["adjacent_content"].append({
                                                "column": adj_col,
                                                "row": adj_row + 2,
                                                "text": str(adj_cell).strip()
                                            })
                            
                            result["priority_content"]["part_b"].append(part_b_entry)
                            result["summary"]["part_b_found"] = True
                            break
        
        # Rest of the processing (sheets processing) remains the same...
        for sheet_name, df in excel_data.items():
            if df.empty:
                continue
            
            # Limit rows if specified
            original_row_count = len(df)
            if max_rows_per_sheet and len(df) > max_rows_per_sheet:
                df = df.head(max_rows_per_sheet)
            
            sheet_data = {
                "sheet_name": sheet_name,
                "dimensions": {
                    "rows": original_row_count,
                    "columns": len(df.columns),
                    "processed_rows": len(df)
                },
                "columns": {
                    "names": [clean_cell_value(col) for col in df.columns.tolist()],
                    "data_types": {clean_cell_value(col): str(dtype) for col, dtype in df.dtypes.to_dict().items()},
                    "numeric_columns": [clean_cell_value(col) for col in df.select_dtypes(include=['number']).columns.tolist()],
                    "key_columns": []
                },
                "sample_data": [],
                "detailed_requirements": [],
                "data_summary": {
                    "missing_data": {},
                    "unique_value_counts": {}
                }
            }
            
            # Identify key columns
            for col in df.columns:
                col_lower = str(col).lower()
                if any(keyword in col_lower for keyword in ['id', 'name', 'title', 'status', 'type', 'category', 'priority', 'requirement']):
                    sheet_data["columns"]["key_columns"].append(clean_cell_value(col))
            
            # Sample data
            sample_size = min(max_sample_rows, len(df))
            if sample_size > 0:
                display_df = df.head(sample_size)
                
                # Convert to records (list of dictionaries)
                for _, row in display_df.iterrows():
                    row_data = {}
                    for col, val in row.items():
                        cleaned_val = clean_cell_value(val)
                        if len(cleaned_val) > 50:
                            cleaned_val = cleaned_val[:47] + "..."
                        row_data[clean_cell_value(col)] = cleaned_val
                    sheet_data["sample_data"].append(row_data)
            
            # Extract detailed requirements
            for col in df.columns:
                col_str = str(col).lower()
                if any(keyword in col_str for keyword in ['requirement', 'detailed', 'description', 'specification']):
                    req_column = {
                        "column_name": clean_cell_value(col),
                        "requirements": []
                    }
                    
                    for idx, cell_value in enumerate(df[col]):
                        if pd.notna(cell_value) and str(cell_value).strip():
                            cell_text = str(cell_value).strip()
                            if len(cell_text) > 10:
                                req_column["requirements"].append({
                                    "row": idx + 2,
                                    "text": cell_text
                                })
                    
                    if req_column["requirements"]:
                        sheet_data["detailed_requirements"].append(req_column)
                        result["summary"]["detailed_requirements_found"] = True
            
            # Data summary - unique values for key columns
            for col in sheet_data["columns"]["key_columns"][:3]:
                if col in df.columns and df[col].dtype == 'object':
                    unique_vals = df[col].dropna().unique()
                    if len(unique_vals) <= 20:
                        sheet_data["data_summary"]["unique_value_counts"][col] = [clean_cell_value(val) for val in unique_vals[:10]]
                    else:
                        sheet_data["data_summary"]["unique_value_counts"][col] = f"{len(unique_vals)} unique values"
            
            # Missing data summary
            missing_data = df.isnull().sum()
            if missing_data.sum() > 0:
                missing_cols = missing_data[missing_data > 0].head(5)
                sheet_data["data_summary"]["missing_data"] = {clean_cell_value(col): int(count) for col, count in missing_cols.items()}
            
            result["sheets"].append(sheet_data)
            result["summary"]["total_rows_processed"] += len(df)
            result["summary"]["total_columns_processed"] += len(df.columns)
    
    except Exception as e:
        result["metadata"]["processing_status"] = "error"
        result["metadata"]["error"] = str(e)
    
    return json.dumps(result, indent=2, ensure_ascii=False)

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
                                value="2025-01-01-preview",
                                placeholder="2025-01-01-preview")
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
