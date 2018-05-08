Attribute VB_Name = "Macros"
Option Explicit

'Version 1.1

Const STRING_SELECT_FACILITIES = 5
Const STRING_SELECT_COURSE = 7
Const SELECT_INSTRUCTOR = 6
Const TITLE_STRING = 9
Const ADD_TEXT_STRING = 8
Const IS_SHARED_STRING = 13

Public CodeBook As Workbook

Sub SelectByForm()
    SelectionForm.InitOnce FormatString(TITLE_STRING), FormatString(ADD_TEXT_STRING), FormatString(IS_SHARED_STRING)
    
    Dim displayShare As Boolean
    Dim multiSelect As Boolean
    
    Dim arr() As String, title As String
    Dim functionPtr As String
    Dim qual As String
    
    functionPtr = "SetCellValue"
    multiSelect = False
    
    If ActiveCell.row = HEADER_ROW Then
        arr = GetCourses()
        title = FormatString(STRING_SELECT_COURSE)
        
    ElseIf ActiveCell.row >= HOURS_START_ROW And ActiveCell.row < HOURS_START_ROW + 32 Then
        arr = GetFacilities(GetParam("Location"))
        title = FormatString(STRING_SELECT_FACILITIES)
        displayShare = True
        functionPtr = "SetFacilityCellValue"
        
    ElseIf ActiveCell.row >= ROW_GUIDE_START And ActiveCell.row < ROW_GUIDE_START + GUIDES_COUNT Then
        qual = ActiveCell.Worksheet.Cells(ActiveCell.row, 1)
        arr = GetInstructors
        arr = GetQualifiedInstructors(arr, qual)
        title = FormatString(SELECT_INSTRUCTOR) + " " + qual
        multiSelect = True
    Else
        Exit Sub
    End If
    
    
    SelectionForm.Load title, functionPtr, ActiveCell.value, arr, displayShare, multiSelect
End Sub


Sub SetCellValue(txtItem As String, share As Boolean, txtAdditionalText As String)
    If Selection.Rows.count > 1 Then
        Selection.Merge False
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
    End If
    ActiveCell.value = ""
    If Len(BTrim(txtItem)) > 0 Then
        If share = True Then
            ActiveCell.value = SHARE_SIGN
        End If
        ActiveCell.value = ActiveCell.value + BTrim(txtItem)
    End If
    
    If (Len(BTrim(txtAdditionalText)) > 0) Then
        If Len(ActiveCell.value) > 0 Then
            ActiveCell.value = ActiveCell.value + vbLf
        Else
            ActiveCell.value = "*"
        End If
        ActiveCell.value = ActiveCell.value + BTrim(txtAdditionalText)
    End If
    
End Sub


Sub SetFacilityCellValue(txtItem As String, share As Boolean, txtAdditionalText As String)
    SetCellValue txtItem, share, txtAdditionalText
    Dim color As Integer
    If Len(txtItem) > 0 Then
        color = FacilityName2Color(txtItem)
        ActiveCell.Interior.ColorIndex = color
    End If
End Sub


Sub syncSchedule()
    UpdateSheet ThisWorkbook.ActiveSheet
End Sub


Sub btnJumpFacility()
    If ActiveWindow.ScrollColumn >= FACILITY_OFFSET Then
        ActiveWindow.ScrollColumn = 1
        UnhideGAP ThisWorkbook.ActiveSheet
    Else
        ActiveWindow.ScrollColumn = FACILITY_OFFSET
        HideGAP ThisWorkbook.ActiveSheet
    End If
End Sub

Function InsertVBComponent(ByVal wb As Workbook, script As String, moduleName As String) As Boolean
    Dim newModule As VBComponent
    'remove the old code
    RemoveVBComponent ActiveWorkbook, moduleName

    Set newModule = wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    newModule.CodeModule.AddFromString script
    newModule.name = moduleName
    InsertVBComponent = True

End Function

Sub RemoveVBComponent(ByVal wb As Workbook, ByVal compName As String)
    On Error Resume Next
    
    With wb.VBProject.VBComponents
        .Remove .Item(compName)
    End With
    On Error GoTo 0

    Set wb = Nothing

End Sub
Function Select_File_Or_Files_Mac() As String
    Dim MyPath As String
    Dim MyScript As String
    On Error Resume Next
    MyPath = MacScript("return (path to documents folder) as String")

    MyScript = _
    "set applescript's text item delimiters to "","" " & vbNewLine & _
               "set theFiles to (choose file " & _
               " with prompt ""Please select a file or files"" default location alias """ & _
               MyPath & """ multiple selections allowed false) as string" & vbNewLine & _
               "set applescript's text item delimiters to """" " & vbNewLine & _
               "return theFiles"

    Select_File_Or_Files_Mac = MacScript(MyScript)
    If Len(Select_File_Or_Files_Mac) > 0 Then
        If Right(Select_File_Or_Files_Mac, 4) <> ".bas" Then
            MsgBox "Only files of type *.bas are allowed"
                Select_File_Or_Files_Mac = ""
        End If
    End If

End Function


Sub UpdateCode()
    Dim strFileToOpen As String
    Dim data As String
    Dim fileStr As String
    
    If Not Application.OperatingSystem Like "*Mac*" Then
        strFileToOpen = Application.GetOpenFilename("*.bas", , "Please choose a code file")
    Else
        strFileToOpen = Select_File_Or_Files_Mac
    End If

    If strFileToOpen <> "" Then
        On Error Resume Next
        Open strFileToOpen For Input As #1
        If Err.Number <> 0 Then
            MsgBox "Error reading code file (*.bas). make sure it has no Hebrew charachters in its path"
            Exit Sub
        End If
        On Error GoTo 0
            Line Input #1, data
            If InStr(data, "Attribute") < 0 Then
                fileStr = fileStr
            End If
            Do Until EOF(1)
                Line Input #1, data
                fileStr = fileStr + vbLf + data
            Loop
        Close #1
        
        Dim moduleName As String
        If InStr(strFileToOpen, "main.bas") > 0 Then
            moduleName = "Main"
        ElseIf InStr(strFileToOpen, "message.bas") > 0 Then
            moduleName = "MasterData"
        ElseIf InStr(strFileToOpen, "masterdata.bas") > 0 Then
            moduleName = "Message"
        End If
        
        If Len(moduleName) > 0 Then
            If InsertVBComponent(ActiveWorkbook, fileStr, moduleName) Then
                MsgBox "Update succeeded!"
                Exit Sub
            End If
            
        End If
    End If
    MsgBox "Update was not completed"
 
End Sub
