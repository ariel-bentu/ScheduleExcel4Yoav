Attribute VB_Name = "Macros"
Option Explicit

Const STRING_SELECT_FACILITIES = 5
Const STRING_SELECT_COURSE = 6
Const SELECT_INSTRUCTOR = 7
 Const TITLE_STRING = 9
 Const ADD_TEXT_STRING = 8

Public CodeBook As Workbook

Sub SelectByForm()
    SelectionForm.InitOnce FormatString(TITLE_STRING), FormatString(ADD_TEXT_STRING)
    

    Dim arr() As String, title As String
    
    If ActiveCell.row = HEADER_ROW Then
        arr = GetCourses()
        title = FormatString(STRING_SELECT_COURSE)
        
    ElseIf ActiveCell.row >= HOURS_START_ROW And ActiveCell.row < HOURS_START_ROW + 32 Then
        arr = GetFacilities(GetParam("Location"))
        title = FormatString(STRING_SELECT_FACILITIES)
    ElseIf ActiveCell.row >= ROW_GUIDE_START And ActiveCell.row < ROW_GUIDE_START + GUIDES_COUNT Then
        arr = GetInstructors
        title = FormatString(SELECT_INSTRUCTOR)
    Else
        Exit Sub
    End If
    
    
    SelectionForm.Load title, ActiveCell, arr
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

Sub InsertVBComponent(ByVal wb As Workbook, ByVal CompFileName As String, moduleName As String)
    'Checking whether CompFileName file exists
    If Dir(CompFileName) <> "" Then
        
        'Ignore errors
'        On Error Resume Next
        
        'Inserts component from file
        'wb.VBProject.VBComponents.Import CompFileName
        Dim newModule As VBComponent
        Dim data As String
        Dim fileStr As String
        Open CompFileName For Input As #1
             Line Input #1, data
            If InStr(data, "Attribute") < 0 Then
                fileStr = fileStr
            End If
            Do Until EOF(1)
                Line Input #1, data
                fileStr = fileStr + vbLf + data
            Loop
        Close #1
    
         Set newModule = wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
       newModule.CodeModule.AddFromString fileStr
       newModule.name = moduleName
       
       
        On Error GoTo 0
    End If
    
    Set wb = Nothing

End Sub

Sub RemoveVBComponent(ByVal wb As Workbook, ByVal compName As String)
    On Error Resume Next
    
    'Inserts component from file
    With wb.VBProject.VBComponents
        .Remove .Item(compName)
    End With
    On Error GoTo 0


    Set wb = Nothing

End Sub

Sub UpdateCode()
    
    'Calling InsertVBComponent procedure
    If MsgBox("Are you sure you want to load new code???", vbOKCancel Or vbCritical, "Caution") = vbOK Then
    
     RemoveVBComponent ActiveWorkbook, "Main"
     InsertVBComponent ActiveWorkbook, ActiveWorkbook.Path + "/main.bas", "Main"
     
     RemoveVBComponent ActiveWorkbook, "Message"
     InsertVBComponent ActiveWorkbook, ActiveWorkbook.Path + "/message.bas", "Message"
    
     RemoveVBComponent ActiveWorkbook, "MasterData"
     InsertVBComponent ActiveWorkbook, ActiveWorkbook.Path + "/masterdata.bas", "MasterData"
   End If
End Sub
