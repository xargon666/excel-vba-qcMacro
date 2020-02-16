Attribute VB_Name = "errorHandling"
Sub errHeadersMissing()
    errMsg = MsgBox("Macro Failed: Couldn't find any headers!!", vbCritical)
End Sub

Sub errZeroRecordsFound()
    errMsg = MsgBox("Macro Failed: Couldn't find any records!!", vbCritical)
End Sub

