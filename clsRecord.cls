VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("QC_Macro")
Option Explicit
' FIELD VARS ================================================================================
' STRINGS ***********************************************************************************
Dim accession_number As String
Dim title As String
Dim record_retention_category As String
Dim access_level As String
Dim language As String
Dim information_sensitivity As String
Dim personally_identifiable_information As String
Dim archive_status As String
Dim author As String
Dim author_id As String
Dim issue_date As String
Dim department As String
Dim originating_organization As String
Dim alliance_name As String
Dim identifier As String
Dim associated_identifier As String
Dim compound_number As String
Dim protocol_number As String
Dim keywords As String
Dim primary_or_copy As String
Dim lnb_author_site As String
Dim lnb_issue_date As String
Dim storage_site As String
Dim information_type As String
Dim information_type_description As String
Dim container_number As String
Dim archive_location As String
Dim bar_code As String
Dim microfilm_number As String
Dim archive_notes As String
Dim description As String
Dim application_name As String
Dim microfilm_location As String
Dim retention_period_start_date As String
Dim retention_period_start_date_event As String
Dim retention_review_date As String
Dim review_outcome As String
Dim loan_id As String
Dim loan_date As String
Dim borrower_name As String
Dim borrower_id As String
Dim item_details As String
Dim loan_due_date As String
Dim loan_status As String
Dim follow_up_date As String
Dim loan_return_date As String
Dim objectid As String
Dim business_unit As String
Dim archive_custodain_group As String
' BOOLEANS ***********************************************************************************
Dim boolRecordRestricted As Boolean
Dim boolAccessRestricted As Boolean
