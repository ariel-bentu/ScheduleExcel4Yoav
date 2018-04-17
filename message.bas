Attribute VB_Name = "Messages"
Const ROW_MESSAGE_START = 2
Const COL_MESSAGE_ID = 1
Const COL_MESSAGE_MSG = 2


Private myForm As UserForm1


Public Function FormatString(ID As Integer, ParamArray params() As Variant) As String
    Dim settings As Worksheet, template As String
    
    
    Set settings = ThisWorkbook.Sheets("Settings")
    
    template = settings.Cells(ROW_MESSAGE_START - 1 + ID, COL_MESSAGE_MSG)
    For i = 0 To UBound(params)
        template = Replace$(template, "%" + CStr(i + 1), params(i))
    Next
    FormatString = template
    
End Function


Public Function SPrintf(template As String, ParamArray tokens()) As String
    Dim i As Long
    For i = 0 To UBound(tokens)
        template = Replace$(template, "%" + CStr(i + 1), tokens(i))
    Next
    SPrintf = template
End Function

Public Sub HebMsgBox(text As String)
    If myForm Is Nothing Then
       Set myForm = New UserForm1
    End If
    
    myForm.SetText (text)
    myForm.Show
End Sub


