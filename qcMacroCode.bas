Attribute VB_Name = "qcMacroCode"
Option Explicit
Option Compare Text
Sub qcMacro()
Dim ws As Worksheet, wb As Workbook
Dim r As New clsRecord, e As New clsError, pv As New clsPicklistValues
Dim dictRecord As New Dictionary, dictError As New Dictionary, dictPicklistValues As New Dictionary
Dim strDirectory As String, strSourceDataFile As String, gridType As String, errMsg As String
Dim lastRow As Long, rowCounter As Long
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
Set dictRecord = New Dictionary
Set dictError = New Dictionary
Set dictPicklistValues = New Dictionary
dictRecord.CompareMode = vbTextCompare
dictError.CompareMode = vbTextCompare
dictPicklistValues.CompareMode = vbTextCompare
' ====================================================================================================
' ADD PICKLIST VALUES TO MEMORY ======================================================================
strDirectory = "D:\Documents\gitProjects\excel-vba-qcMacro\" ' UPDATE THIS PATH TO LOCAL WORK DIR
strSourceDataFile = "dm_dbo.dictionary.xls"
Set wbSource = Workbooks(strDirectory & strSourceDataFile)
wbSource.Open
With wbSource
    lastRow = .Cells(.Rows.Count, rngToprow.Column).End(xlUp).Row
    If lastRow < 2 Then Call errZeroPicklistValues
    arrHeader(0) = .Range("1:1").Find("pier_property_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    arrHeader(1) = .Range("1:1").Find("pier_property_value", LookIn:=xlValues, LookAt:=xlWhole).Column
    arrHeader(2) = .Range("1:1").Find("pier_value_is_active", LookIn:=xlValues, LookAt:=xlWhole).Column
    arrHeader(3) = .Range("1:1").Find("source", LookIn:=xlValues, LookAt:=xlWhole).Column
    For rowCounter = 2 To lastRow
        Set pv = New clsPicklistValues
        pv.pier_property_name = .Cells(rowCounter, arrHeader(0))
        pv.pier_property_value = .Cells(rowCounter, arrHeader(1))
        pv.pier_value_is_active = .Cells(rowCounter, arrHeader(2))
        pv.source = .Cells(rowCounter, arrHeader(3))
'        if dictPickListValues.Exists(pv.pier_property_name)
'        dictPicklistValues.Add Key:=pv, item
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
For recordNumber = 1 To dictRecord.Count
    With dictRecord.Item(recordNumber)
        ' ACCESS LEVEL
        Select Case True
        Case .access_level = "Files Restricted"
            .boolAccessRestricted = True
            
        End Select
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
