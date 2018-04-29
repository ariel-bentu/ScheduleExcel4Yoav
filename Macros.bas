Attribute VB_Name = "Macros"
Option Explicit

'Version 1.0

Const STRING_SELECT_FACILITIES = 5
Const STRING_SELECT_COURSE = 6
Const SELECT_INSTRUCTOR = 7
Const TITLE_STRING = 9
Const ADD_TEXT_STRING = 8
Const IS_SHARED_STRING = 13

Public CodeBook As Workbook

Sub SelectByForm()
    SelectionForm.InitOnce FormatString(TITLE_STRING), FormatString(ADD_TEXT_STRING), FormatString(IS_SHARED_STRING)
    
    Dim displayShare As Boolean
    Dim arr() As String, title As String
    
    If ActiveCell.row = HEADER_ROW Then
        arr = GetCourses()
        title = FormatString(STRING_SELECT_COURSE)
        
    ElseIf ActiveCell.row >= HOURS_START_ROW And ActiveCell.row < HOURS_START_ROW + 32 Then
        arr = GetFacilities(GetParam("Location"))
        title = FormatString(STRING_SELECT_FACILITIES)
        displayShare = True
    ElseIf ActiveCell.row >= ROW_GUIDE_START And ActiveCell.row < ROW_GUIDE_START + GUIDES_COUNT Then
        arr = GetInstructors
        title = FormatString(SELECT_INSTRUCTOR)
    Else
        Exit Sub
    End If
    
    
    SelectionForm.Load title, ActiveCell, arr, displayShare
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
