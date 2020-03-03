Attribute VB_Name = "qcMacroCode"
'@Folder("VBAProject")
Option Explicit
Option Compare Text
Sub qcMacro()
Dim ws As Worksheet, wb As Workbook
Dim r As New clsRecord, e As New clsError, pv As New clsPicklistValues

Dim strDirectory As String, strSourceDataFile As String, gridType As String, errMsg As String
Dim lastRow As Long, rowCounter As Long, i As Long
Dim wbSource As Workbook
Dim rngToprow_pv As Range
Dim arrHeader(3) As Variant

Set ws = ActiveSheet
gridType = gridTypeDetection(ws)
Call findHeaders(gridType)
If gridType = arrType(5) Then Call errHeadersMissing
With ws
    lastRow = .Cells(.Rows.Count, rngToprow.Column).End(xlUp).Row
End With
If lastRow < rngToprow.Row + 1 Then Call errZeroRecordsFound
' ==============================================================================================
' DICTIONARY SETUP =============================================================================
Dim dictRecord As New Dictionary, dictError As New Dictionary, dictPicklistValues As New Dictionary
Dim dict_alliance_names As New Dictionary, dict_application_name As New Dictionary, dict_archive_custodian_group As New Dictionary
Dim dict_archive_status As New Dictionary, dict_business_unit As New Dictionary, dict_followup_status As New Dictionary
Dim dict_information_sensitivity As New Dictionary, dict_lnb_author_site As New Dictionary, dict_organization As New Dictionary
Dim dict_personal_identify_inf As New Dictionary, dict_pier_content_type As New Dictionary, dict_pier_department As New Dictionary
Dim dict_pier_languages As New Dictionary, dict_pier_review_site As New Dictionary, dict_pier_urgency As New Dictionary
Dim dict_piera_archive_site As New Dictionary, dict_piera_item_type As New Dictionary, dict_piera_microfilm_location As New Dictionary
Dim dict_primary_or_copy As New Dictionary, dict_reason_for_rejection As New Dictionary, dict_retention_review_outcome As New Dictionary
Dim dict_retention_start_date_event As New Dictionary
Dim dict_source_database As New Dictionary
Set dictRecord = New Dictionary
Set dictError = New Dictionary
Set dictPicklistValues = New Dictionary
Set dict_alliance_names = New Dictionary
Set dict_application_name = New Dictionary
Set dict_archive_custodian_group = New Dictionary
Set dict_archive_status = New Dictionary
Set dict_business_unit = New Dictionary
Set dict_followup_status = New Dictionary
Set dict_information_sensitivity = New Dictionary
Set dict_lnb_author_site = New Dictionary
Set dict_organization = New Dictionary
Set dict_personal_identify_inf = New Dictionary
Set dict_pier_content_type = New Dictionary
Set dict_pier_department = New Dictionary
Set dict_pier_languages = New Dictionary
Set dict_pier_review_site = New Dictionary
Set dict_pier_urgency = New Dictionary
Set dict_piera_archive_site = New Dictionary
Set dict_piera_item_type = New Dictionary
Set dict_piera_microfilm_location = New Dictionary
Set dict_primary_or_copy = New Dictionary
Set dict_reason_for_rejection = New Dictionary
Set dict_retention_category = New Dictionary
Set dict_retention_review_outcome = New Dictionary
Set dict_retention_start_date_event = New Dictionary
Set dict_source_database = New Dictionary
dictRecord.CompareMode = vbTextCompare
dictError.CompareMode = vbTextCompare
dictPicklistValues.CompareMode = vbTextCompare
dict_alliance_names.CompareMode = vbTextCompare
dict_application_name.CompareMode = vbTextCompare
dict_archive_custodian_group.CompareMode = vbTextCompare
dict_archive_status.CompareMode = vbTextCompare
dict_business_unit.CompareMode = vbTextCompare
dict_followup_status.CompareMode = vbTextCompare
dict_information_sensitivity.CompareMode = vbTextCompare
dict_lnb_author_site.CompareMode = vbTextCompare
dict_organization.CompareMode = vbTextCompare
dict_personal_identify_inf.CompareMode = vbTextCompare
dict_pier_content_type.CompareMode = vbTextCompare
dict_pier_department.CompareMode = vbTextCompare
dict_pier_languages.CompareMode = vbTextCompare
dict_pier_review_site.CompareMode = vbTextCompare
dict_pier_urgency.CompareMode = vbTextCompare
dict_piera_archive_site.CompareMode = vbTextCompare
dict_piera_item_type.CompareMode = vbTextCompare
dict_piera_microfilm_location.CompareMode = vbTextCompare
dict_primary_or_copy.CompareMode = vbTextCompare
dict_reason_for_rejection.CompareMode = vbTextCompare
dict_retention_category.CompareMode = vbTextCompare
dict_retention_review_outcome.CompareMode = vbTextCompare
dict_retention_start_date_event.CompareMode = vbTextCompare
dict_source_database.CompareMode = vbTextCompare

' ====================================================================================================
' ADD PICKLIST VALUES TO MEMORY ======================================================================
strDirectory = "D:\Documents\gitProjects\excel-vba-qcMacro\" ' UPDATE THIS PATH TO LOCAL WORK DIR
strSourceDataFile = "dm_dbo.dictionary.xls"
Dim arrSource(0) As Variant
Set wbSource = Workbooks(strDirectory & strSourceDataFile)
wbSource.Open
With wbSource.ActiveSheet
    lastRow = .Cells(.Rows.Count, .CurrentRegion.Column).End(xlUp).Row ' could possibly go wrong
    If lastRow < 2 Then Call errZeroPicklistValues
    arrHeader(0) = .Range("1:1").Find("pier_property_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    arrHeader(1) = .Range("1:1").Find("pier_property_value", LookIn:=xlValues, LookAt:=xlWhole).Column
    arrHeader(2) = .Range("1:1").Find("pier_value_is_active", LookIn:=xlValues, LookAt:=xlWhole).Column
    arrHeader(3) = .Range("1:1").Find("source", LookIn:=xlValues, LookAt:=xlWhole).Column
    For rowCounter = 2 To lastRow
            Dim target As Range
            Set target = .Cells(rowCounter, arrHeader(0))
            If target.Value <> arrSource(i) Then
                i = i + 1
                ReDim Preserve arrSource(i) As Variant
                arrSource(i) = target.Value
            End If
    Next rowCounter

    For rowCounter = 2 To lastRow
        For i = LBound(arrSource) To UBound(arrSource)
            If .Cells(rowCounter, arrHeader(0)).Value = arrSource(i) Then
                Set pv = New clsPicklistValues
                pv.pier_property_name = .Cells(rowCounter, arrHeader(0))
                pv.pier_property_value = .Cells(rowCounter, arrHeader(1))
                pv.pier_value_is_active = .Cells(rowCounter, arrHeader(2))
                pv.source = .Cells(rowCounter, arrHeader(3))
                If arrSource(i) = "alliance_names" Then dict_alliance_names.Add dict_alliance_names.Count, pv
                If arrSource(i) = "application_name" Then dict_application_name.Add dict_application_name.Count, pv
                If arrSource(i) = "archive_custodian_group" Then dict_archive_custodian_group.Add dict_archive_custodian_group.Count, pv
                If arrSource(i) = "archive_status" Then dict_archive_status.Add dict_archive_status.Count, pv
                If arrSource(i) = "business_unit" Then dict_business_unit.Add dict_business_unit.Count, pv
                If arrSource(i) = "followup_status" Then dict_followup_status.Add dict_followup_status.Count, pv
                If arrSource(i) = "information_sensitivity" Then dict_information_sensitivity.Add dict_information_sensitivity.Count, pv
                If arrSource(i) = "lnb_author_site" Then dict_lnb_author_site.Add dict_lnb_author_site.Count, pv
                If arrSource(i) = "organization" Then dict_organization.Add dict_organization.Count, pv
                If arrSource(i) = "personal_identify_inf" Then dict_personal_identify_inf.Add dict_personal_identify_inf.Count, pv
                If arrSource(i) = "pier_content_type" Then dict_pier_content_type.Add dict_pier_content_type.Count, pv
                If arrSource(i) = "pier_department" Then dict_pier_department.Add dict_pier_department.Count, pv
                If arrSource(i) = "pier_languages" Then dict_pier_languages.Add dict_pier_languages.Count, pv
                If arrSource(i) = "pier_review_site" Then dict_pier_review_site.Add dict_pier_review_site.Count, pv
                If arrSource(i) = "pier_urgency" Then dict_pier_urgency.Add dict_pier_urgency.Count, pv
                If arrSource(i) = "piera_archive_site" Then dict_piera_archive_site.Add dict_piera_archive_site.Count, pv
                If arrSource(i) = "piera_item_type" Then dict_piera_item_type.Add dict_piera_item_type.Count, pv
                If arrSource(i) = "piera_microfilm_location" Then dict_piera_microfilm_location.Add dict_piera_microfilm_location.Count, pv
                If arrSource(i) = "primary_or_copy" Then dict_primary_or_copy.Add dict_primary_or_copy.Count, pv
                If arrSource(i) = "reason_for_rejection" Then dict_reason_for_rejection.Add dict_reason_for_rejection.Count, pv
                If arrSource(i) = "retention_category" Then dict_retention_category.Add dict_retention_category.Count, pv
                If arrSource(i) = "retention_review_outcome" Then dict_retention_review_outcome.Add dict_retention_review_outcome.Count, pv
                If arrSource(i) = "retention_start_date_event" Then dict_retention_start_date_event.Add dict_retention_start_date_event.Count, pv
                If arrSource(i) = "source_database" Then dict_source_database.Add dict_source_database.Count, pv
            End If
        Next i
    Next rowCounter
    Erase arrHeader
End With

' ====================================================================================================
' ADD RECORD DATA TO MEMORY ==========================================================================
For rowCounter = rngToprow.Row + 1 To lastRow
    Set r = New clsRecord
    r.accession_number = ws.Cells(rowCounter, col_accession_number)
    r.title = ws.Cells(rowCounter, col_title)
    r.record_retention_category = ws.Cells(rowCounter, col_record_retention_category)
    r.access_level = ws.Cells(rowCounter, col_access_level)
    r.language = ws.Cells(rowCounter, col_language)
    r.information_sensitivity = ws.Cells(rowCounter, col_information_sensitivity)
    r.personally_identifiable_information = ws.Cells(rowCounter, col_personally_identifiable_information)
    r.archive_status = ws.Cells(rowCounter, col_archive_status)
    r.author = ws.Cells(rowCounter, col_author)
    r.author_id = ws.Cells(rowCounter, col_author_id)
    r.issue_date = ws.Cells(rowCounter, col_issue_date)
    r.department = ws.Cells(rowCounter, col_department)
    r.originating_organization = ws.Cells(rowCounter, col_originating_organization)
    r.alliance_name = ws.Cells(rowCounter, col_alliance_name)
    r.identifier = ws.Cells(rowCounter, col_identifier)
    r.associated_identifier = ws.Cells(rowCounter, col_associated_identifier)
    r.compound_number = ws.Cells(rowCounter, col_compound_number)
    r.protocol_number = ws.Cells(rowCounter, col_protocol_number)
    r.keywords = ws.Cells(rowCounter, col_keywords)
    r.primary_or_copy = ws.Cells(rowCounter, col_primary_or_copy)
    r.lnb_author_site = ws.Cells(rowCounter, col_lnb_author_site)
    r.lnb_issue_date = ws.Cells(rowCounter, col_lnb_issue_date)
    r.storage_site = ws.Cells(rowCounter, col_storage_site)
    r.information_type = ws.Cells(rowCounter, col_information_type)
    r.information_type_description = ws.Cells(rowCounter, col_information_type_description)
    r.container_number = ws.Cells(rowCounter, col_container_number)
    r.archive_location = ws.Cells(rowCounter, col_archive_location)
    r.bar_code = ws.Cells(rowCounter, col_bar_code)
    r.microfilm_number = ws.Cells(rowCounter, col_microfilm_number)
    r.archive_notes = ws.Cells(rowCounter, col_archive_notes)
    r.description = ws.Cells(rowCounter, col_description)
    r.application_name = ws.Cells(rowCounter, col_application_name)
    r.microfilm_location = ws.Cells(rowCounter, col_microfilm_location)
    r.retention_period_start_date = ws.Cells(rowCounter, col_retention_period_start_date)
    r.retention_period_start_date_event = ws.Cells(rowCounter, col_retention_period_start_date_event)
    r.retention_review_date = ws.Cells(rowCounter, col_retention_review_date)
    r.review_outcome = ws.Cells(rowCounter, col_review_outcome)
    r.loan_id = ws.Cells(rowCounter, col_loan_id)
    r.loan_date = ws.Cells(rowCounter, col_loan_date)
    r.borrower_name = ws.Cells(rowCounter, col_borrower_name)
    r.borrower_id = ws.Cells(rowCounter, col_borrower_id)
    r.item_details = ws.Cells(rowCounter, col_item_details)
    r.loan_due_date = ws.Cells(rowCounter, col_loan_due_date)
    r.loan_status = ws.Cells(rowCounter, col_loan_status)
    r.follow_up_date = ws.Cells(rowCounter, col_follow_up_date)
    r.loan_return_date = ws.Cells(rowCounter, col_loan_return_date)
    r.objectid = ws.Cells(rowCounter, col_objectid)
    r.business_unit = ws.Cells(rowCounter, col_business_unit)
    r.archive_custodain_group = ws.Cells(rowCounter, col_archive_custodain_group)
    dictRecord.Add Key:=r.objectid, Item:=r
Next rowCounter
If dictRecord.Count < 1 Then Call errZeroRecordsFound
' ====================================================================================================
' LOOP THROUGH COLLECTION ============================================================================
Dim recordNumber As Long
For recordNumber = 1 To dictRecord.Count ' <= MAIN RECORD FOR LOOP
    With dictRecord.Item(recordNumber)
        ' ACCESS LEVEL - VALUES *NOT* STORED IN dbo.dm_dictionary
        Select Case True
        Case .access_level = "General Access"
            .boolAccessRestricted = False
        Case .access_level <> "General Access"
            .boolAccessRestricted = True
        End Select
        ' AUTHOR
        Select Case True
            Case (InStr(Trim(.author), " ") > 0) & (InStr(Trim(.author), ",") = 0) ' AUTHOR CONTAINS SPACE BUT NO COMMA
            
            Case Trim(.author) = "unknwon" Or Trim(.author) = "unknow" Or Trim(.author) = "unkown" ' UNKNOWN MISPELLED
            
    End With
Next recordNumber
' ====================================================================================================
' ERASE DATA IN MEMORY ===============================================================================
Set dictRecord = Nothing
Set dictError = Nothing
Set dictPicklistValues = Nothing
Erase arrType
Exit Sub
' ************************************** END OF CODE *************************************************

End Sub

Function HasNumber(strData As String) As Boolean
    Dim iCnt As Integer
    For iCnt = 1 To Len(strData)
        If IsNumeric(Mid(strData, iCnt, 1)) Then
            HasNumber = True
            Exit Function
        End If
    Next iCnt
End Function
