Attribute VB_Name = "errorHandling"
'@Folder("VBAProject")
Sub errHeadersMissing()
    errMsg = MsgBox("Macro Failed: Couldn't find any headers!!", vbCritical)
End Sub
Sub errZeroRecordsFound()
    errMsg = MsgBox("Macro Failed: Couldn't find any records!!", vbCritical)
End Sub
Sub errZeroPicklistValues()
    errMsg = MsgBox("Macro Failed: Couldn't any Picklist Values in the dm_dbo.dictionary extract", vbCritical)
End Sub


