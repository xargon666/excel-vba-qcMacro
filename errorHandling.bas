Attribute VB_Name = "errorHandling"
'@Folder("QC_Macro")
Option Explicit
Public errMsg As String
Public Sub errHeadersMissing()
    errMsg = MsgBox("Macro Failed: Couldn't find any headers!!", vbCritical)
End Sub

Public Sub errZeroRecordsFound()
    errMsg = MsgBox("Macro Failed: Couldn't find any records!!", vbCritical)
End Sub

