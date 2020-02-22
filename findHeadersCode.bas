Attribute VB_Name = "findHeadersCode"
'@Folder("QC_Macro")
Option Compare Text
Option Explicit
Dim col_title As Long
Dim col_author As Long
Dim col_author_id As Long
Dim col_identifier As Long
Dim col_associated_identifier As Long
Dim col_issue_date As Long
Dim col_department As Long
Dim col_compound_number As Long
Dim col_protocol_number As Long
Dim col_keywords As Long
Dim col_originating_organization As Long
Dim col_source_database As Long
Dim col_language As Long
Dim col_archive_status As Long
Dim col_alliance_name As Long
Dim col_application_name As Long
Dim col_lnb_author_site As Long
Dim col_lnb_issue_date As Long
Dim col_information_sensitivity As Long
Dim col_personally_identifiable_information As Long
Dim col_primary_or_copy As Long
Dim col_archive_notes As Long
Dim col_storage_site As Long
Dim col_microfilm_location As Long
Dim col_description As Long
Dim col_aadf_tracking_number As Long
Dim col_accession_number As Long
Dim col_archive_location As Long
Dim col_bar_code As Long
Dim col_borrower_id As Long
Dim col_borrower_name As Long
Dim col_container_number As Long
Dim col_create_date As Long
Dim col_creator As Long
Dim col_follow_up_date As Long
Dim col_information_type As Long
Dim col_information_type_description As Long
Dim col_item_details As Long
Dim col_loan_date As Long
Dim col_loan_due_date As Long
Dim col_loan_id As Long
Dim col_loan_return_date As Long
Dim col_loan_status As Long
Dim col_microfilm_number As Long
Dim col_modified_by As Long
Dim col_record_retention_category As Long
Dim col_retention_period_start_date As Long
Dim col_retention_period_start_date_event As Long
Dim col_retention_review_date As Long
Dim col_review_outcome As Long
Dim col_submitter_id As Long
Dim col_has_econtent As Long
Dim col_total_file_size As Long
Dim col_access_level As Long
Dim col_objectid As Long
Dim col_business_unit As Long
Dim col_archive_custodian_group As Long
Dim ws As Worksheet

Sub findHeaders(gridType As String)
    Select Case True
    Case gridType = arrType(0)
        On Error Resume Next
        col_accession_number = rngToprow.Find("Accession Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_title = rngToprow.Find("Title", LookIn:=xlValues, LookAt:=xlPart).Column
        col_record_retention_category = rngToprow.Find("Record Retention Category", LookIn:=xlValues, LookAt:=xlPart).Column
        col_access_level = rngToprow.Find("Access Level", LookIn:=xlValues, LookAt:=xlPart).Column
        col_language = rngToprow.Find("Language", LookIn:=xlValues, LookAt:=xlPart).Column
        col_information_sensitivity = rngToprow.Find("Information Sensitivity", LookIn:=xlValues, LookAt:=xlPart).Column
        col_personally_identifiable_information = rngToprow.Find("Personally Identifiable Information", LookIn:=xlValues, LookAt:=xlPart).Column
        col_archive_status = rngToprow.Find("Archive Status", LookIn:=xlValues, LookAt:=xlPart).Column
        col_author = rngToprow.Find("Author", LookIn:=xlValues, LookAt:=xlPart).Column
        col_author_id = rngToprow.Find("Author ID", LookIn:=xlValues, LookAt:=xlPart).Column
        col_issue_date = rngToprow.Find("Issue Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_department = rngToprow.Find("Department", LookIn:=xlValues, LookAt:=xlPart).Column
        col_originating_organization = rngToprow.Find("Originating Organization", LookIn:=xlValues, LookAt:=xlPart).Column
        col_alliance_name = rngToprow.Find("Alliance Name", LookIn:=xlValues, LookAt:=xlPart).Column
        col_identifier = rngToprow.Find("Identifier", LookIn:=xlValues, LookAt:=xlPart).Column
        col_associated_identifier = rngToprow.Find("Associated Identifier", LookIn:=xlValues, LookAt:=xlPart).Column
        col_compound_number = rngToprow.Find("Compound Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_protocol_number = rngToprow.Find("Protocol Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_keywords = rngToprow.Find("Keywords", LookIn:=xlValues, LookAt:=xlPart).Column
        col_primary_or_copy = rngToprow.Find("Primary Or Copy", LookIn:=xlValues, LookAt:=xlPart).Column
        col_lnb_author_site = rngToprow.Find("LNB Author Site", LookIn:=xlValues, LookAt:=xlPart).Column
        col_lnb_issue_date = rngToprow.Find("LNB Issue Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_storage_site = rngToprow.Find("Storage Site", LookIn:=xlValues, LookAt:=xlPart).Column
        col_information_type = rngToprow.Find("Information Type", LookIn:=xlValues, LookAt:=xlPart).Column
        col_information_type_description = rngToprow.Find("Information Type Description", LookIn:=xlValues, LookAt:=xlPart).Column
        col_container_number = rngToprow.Find("Container Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_archive_location = rngToprow.Find("Archive Location", LookIn:=xlValues, LookAt:=xlPart).Column
        col_bar_code = rngToprow.Find("Bar Code", LookIn:=xlValues, LookAt:=xlPart).Column
        col_microfilm_number = rngToprow.Find("Microfilm Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_archive_notes = rngToprow.Find("Archive Notes", LookIn:=xlValues, LookAt:=xlPart).Column
        col_description = rngToprow.Find("Description", LookIn:=xlValues, LookAt:=xlPart).Column
        col_application_name = rngToprow.Find("Application Name", LookIn:=xlValues, LookAt:=xlPart).Column
        col_microfilm_location = rngToprow.Find("Microfilm Location", LookIn:=xlValues, LookAt:=xlPart).Column
        col_retention_period_start_date = rngToprow.Find("Retention Period Start Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_retention_period_start_date_event = rngToprow.Find("Retention Period Start Date Event", LookIn:=xlValues, LookAt:=xlPart).Column
        col_retention_review_date = rngToprow.Find("Retention Review Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_review_outcome = rngToprow.Find("Review Outcome", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_id = rngToprow.Find("Loan ID", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_date = rngToprow.Find("Loan Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_borrower_name = rngToprow.Find("Borrower Name", LookIn:=xlValues, LookAt:=xlPart).Column
        col_borrower_id = rngToprow.Find("Borrower ID", LookIn:=xlValues, LookAt:=xlPart).Column
        col_item_details = rngToprow.Find("Item Details", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_due_date = rngToprow.Find("Loan Due Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_status = rngToprow.Find("Loan Status", LookIn:=xlValues, LookAt:=xlPart).Column
        col_follow_up_date = rngToprow.Find("Follow-up Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_return_date = rngToprow.Find("Loan Return Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_objectid = rngToprow.Find("ObjectId", LookIn:=xlValues, LookAt:=xlPart).Column
        col_business_unit = rngToprow.Find("Business Unit", LookIn:=xlValues, LookAt:=xlPart).Column
        col_archive_custodian_group = rngToprow.Find("Archive Custodian Group", LookIn:=xlValues, LookAt:=xlPart).Column
        If col_archive_custodian_group Is Nothing Then
            col_archive_custodian_group = rngToprow.Find("Archive Custodain Group", LookIn:=xlValues, LookAt:=xlPart).Column
        End If

    Case gridType = arrType(3)
        col_accession_number = rngToprow.Find("Accession Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_title = rngToprow.Find("Title", LookIn:=xlValues, LookAt:=xlPart).Column
        col_record_retention_category = rngToprow.Find("Record Retention Category", LookIn:=xlValues, LookAt:=xlPart).Column
        col_access_level = rngToprow.Find("Access Level", LookIn:=xlValues, LookAt:=xlPart).Column
        col_language = rngToprow.Find("Language", LookIn:=xlValues, LookAt:=xlPart).Column
        col_information_sensitivity = rngToprow.Find("Information Sensitivity", LookIn:=xlValues, LookAt:=xlPart).Column
        col_personally_identifiable_information = rngToprow.Find("Personally Identifiable Information", LookIn:=xlValues, LookAt:=xlPart).Column
        col_archive_status = rngToprow.Find("Archive Status", LookIn:=xlValues, LookAt:=xlPart).Column
        col_author = rngToprow.Find("Author", LookIn:=xlValues, LookAt:=xlPart).Column
        col_author_id = rngToprow.Find("Author ID", LookIn:=xlValues, LookAt:=xlPart).Column
        col_issue_date = rngToprow.Find("Issue Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_department = rngToprow.Find("Department", LookIn:=xlValues, LookAt:=xlPart).Column
        col_originating_organization = rngToprow.Find("Originating Organization", LookIn:=xlValues, LookAt:=xlPart).Column
        col_alliance_name = rngToprow.Find("Alliance Name", LookIn:=xlValues, LookAt:=xlPart).Column
        col_identifier = rngToprow.Find("Identifier", LookIn:=xlValues, LookAt:=xlPart).Column
        col_associated_identifier = rngToprow.Find("Associated Identifier", LookIn:=xlValues, LookAt:=xlPart).Column
        col_compound_number = rngToprow.Find("Compound Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_protocol_number = rngToprow.Find("Protocol Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_keywords = rngToprow.Find("Keywords", LookIn:=xlValues, LookAt:=xlPart).Column
        col_primary_or_copy = rngToprow.Find("Primary Or Copy", LookIn:=xlValues, LookAt:=xlPart).Column
        col_lnb_author_site = rngToprow.Find("LNB Author Site", LookIn:=xlValues, LookAt:=xlPart).Column
        col_lnb_issue_date = rngToprow.Find("LNB Issue Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_storage_site = rngToprow.Find("Storage Site", LookIn:=xlValues, LookAt:=xlPart).Column
        col_information_type = rngToprow.Find("Information Type", LookIn:=xlValues, LookAt:=xlPart).Column
        col_information_type_description = rngToprow.Find("Information Type Description", LookIn:=xlValues, LookAt:=xlPart).Column
        col_container_number = rngToprow.Find("Container Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_archive_location = rngToprow.Find("Archive Location", LookIn:=xlValues, LookAt:=xlPart).Column
        col_bar_code = rngToprow.Find("Bar Code", LookIn:=xlValues, LookAt:=xlPart).Column
        col_microfilm_number = rngToprow.Find("Microfilm Number", LookIn:=xlValues, LookAt:=xlPart).Column
        col_archive_notes = rngToprow.Find("Archive Notes", LookIn:=xlValues, LookAt:=xlPart).Column
        col_description = rngToprow.Find("Description", LookIn:=xlValues, LookAt:=xlPart).Column
        col_application_name = rngToprow.Find("Application Name", LookIn:=xlValues, LookAt:=xlPart).Column
        col_microfilm_location = rngToprow.Find("Microfilm Location", LookIn:=xlValues, LookAt:=xlPart).Column
        col_retention_period_start_date = rngToprow.Find("Retention Period Start Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_retention_period_start_date_event = rngToprow.Find("Retention Period Start Date Event", LookIn:=xlValues, LookAt:=xlPart).Column
        col_retention_review_date = rngToprow.Find("Retention Review Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_review_outcome = rngToprow.Find("Review Outcome", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_id = rngToprow.Find("Loan ID", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_date = rngToprow.Find("Loan Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_borrower_name = rngToprow.Find("Borrower Name", LookIn:=xlValues, LookAt:=xlPart).Column
        col_borrower_id = rngToprow.Find("Borrower ID", LookIn:=xlValues, LookAt:=xlPart).Column
        col_item_details = rngToprow.Find("Item Details", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_due_date = rngToprow.Find("Loan Due Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_status = rngToprow.Find("Loan Status", LookIn:=xlValues, LookAt:=xlPart).Column
        col_follow_up_date = rngToprow.Find("Follow-up Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_loan_return_date = rngToprow.Find("Loan Return Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_objectid = rngToprow.Find("ObjectId", LookIn:=xlValues, LookAt:=xlPart).Column
        col_business_unit = rngToprow.Find("Business Unit", LookIn:=xlValues, LookAt:=xlPart).Column
        col_archive_custodian_group = rngToprow.Find("Archive Custodian Group", LookIn:=xlValues, LookAt:=xlPart).Column
        If col_archive_custodian_group Is Nothing Then
            col_archive_custodian_group = rngToprow.Find("Archive Custodain Group", LookIn:=xlValues, LookAt:=xlPart).Column
        End If
        col_source_database = rngToprow.Find("Source", LookIn:=xlValues, LookAt:=xlPart).Column
        col_aadf_tracking_number = rngToprow.Find("aadf tracking", LookIn:=xlValues, LookAt:=xlPart).Column
        col_create_date = rngToprow.Find("Create Date", LookIn:=xlValues, LookAt:=xlPart).Column
        col_creator = rngToprow.Find("Creator", LookIn:=xlValues, LookAt:=xlPart).Column
        col_modified_by = rngToprow.Find("Source", LookIn:=xlValues, LookAt:=xlPart).Column
        col_submitter_id = rngToprow.Find("Submitter", LookIn:=xlValues, LookAt:=xlPart).Column
        col_has_econtent = rngToprow.Find("has econtent", LookIn:=xlValues, LookAt:=xlPart).Column
        col_total_file_size = rngToprow.Find("total file size", LookIn:=xlValues, LookAt:=xlPart).Column
    End Select

    On Error GoTo 0
    Exit Sub
End Sub

