Attribute VB_Name = "gridTypeDetectionFunction"
Option Compare Text
Public rngToprow As Range
Public arrType(5) As String
Public Function gridTypeDetection(targetSheet As Worksheet)
Dim a As Double
Dim c As Double
' Forces Recount of Rows/Cols fixing UsedRange bug
a = ActiveSheet.UsedRange.Rows.Count
c = ActiveSheet.UsedRange.Columns.Count
Dim rngFind As Range
Dim rowCount As Long, colCount As Long
Dim firstHeaderValue As String, lastHeaderValue As String
Dim firstHeaderPosition As Long, lastHeaderPosition As Long
' ASSIGN GRID TYPES TO ARRAY =============================================================================
arrType(0) = "Multiedit Data Table"
arrType(1) = "Information Type Grid"
arrType(2) = "Loan Grid"
arrType(3) = "Data Export - Default Columns"
arrType(4) = "Data Export - Custom Columns"
arrType(5) = "N/A"
ReDim valFind(2) As String
valFind(0) = "Accession"
valFind(1) = "Information"
valFind(2) = "Submitter"

' DETERMINE GRID TYPE ====================================================================================
On Error Resume Next
With targetSheet.UsedRange
For loopCounter = 0 To 2
    Set rngFind = .Find(What:=valFind(loopCounter), LookIn:=xlValues, LookAt:=xlPart)
    If Not rngFind Is Nothing Then Exit For
Next loopCounter
Erase valFind
    If rngFind Is Nothing Then GoTo headersNotFound
    Set rngToprow = .Range(Cells(rngFind.Row, rngFind.Column), Cells(rngFind.Row, .Columns.Count))
    If rngToprow Is Nothing Then GoTo headersNotFound
End With
With targetSheet
    rowCount = .Cells(rngToprow.Row, rngToprow.Column).End(xlDown).Row
    colCount = .Cells(rngToprow.Row, rngToprow.Column).End(xlToRight).Column
    firstHeaderPosition = .Cells(rngToprow.Row, rngToprow.Column).End(xlToLeft).Column
    lastHeaderPosition = .Cells(rngToprow.Row, rngToprow.Columns.Count).End(xlToLeft).Column
    firstHeaderValue = Cells(rngToprow.Row, firstHeaderPosition)
    lastHeaderValue = Cells(rngToprow.Row, lastHeaderPosition)
End With
Select Case True
    Case _
    rowCount <= 201 And _
    colCount = 49 And _
    InStr(1, firstHeaderValue, "Accession") > 0 And _
    (InStr(1, lastHeaderValue, "Custodian") > 0 Or InStr(1, lastHeaderValue, "Custodain") > 0)
        gridTypeDetection = arrType(0)
        Exit Function
    Case _
    colCount = 7 And _
    InStr(1, firstHeaderValue, "Information") > 0 And _
    InStr(1, lastHeaderValue, "Microfilm") > 0
        gridTypeDetection = arrType(1)
        Exit Function
    Case True
        gridTypeDetection = arrType(5)
End Select
    'rowCount <= 201 and colCount = 49 and InStr(1,firstHeaderValue,"Accession")>0 and InStr(1,lastHeaderValue,"Custodian")>0

Exit Function
headersNotFound:
MsgBox "Headers Missing!"
End Function
