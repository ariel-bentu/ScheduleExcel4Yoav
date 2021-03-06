Attribute VB_Name = "MasterData"
'Version 1.2

Type Facility
    name As String
    location As String
    Group As String
    color As Integer
    ID As Integer
End Type

Type Instructor
    name As String
    qualifications() As String
    
End Type


Dim MasterData As Workbook
Private facilities(1000) As Facility
Private FacilitiesCount As Integer
Private courses(1000) As String
Private CoursesCount As Integer

Private instructors(1000) As Instructor
Private InstructorsCount As Integer


Dim masterDataTimeStamp As Single


Sub initMasterData()

    'cache for 60 seconds
    If masterDataTimeStamp + 60 > Timer Then Exit Sub

    Dim s As Worksheet, row As Integer
    If MasterData Is Nothing Then
    
    
        On Error Resume Next
        Set MasterData = Workbooks.Open(ActiveWorkbook.Path & "/masterdata.xlsx", , True)
       If Err.Number <> 0 Then
            HebMsgBox FormatString(12, Err.Description)
            Exit Sub
        End If
         
        row = 2
        Set s = MasterData.Sheets("Facilities")
        If Err.Number <> 0 Then
            HebMsgBox FormatString(12, Err.Description)
            Exit Sub
        End If
        
        FacilitiesCount = 0
        While BTrim(s.Cells(row, 2)) <> ""
            FacilitiesCount = FacilitiesCount + 1
            With facilities(FacilitiesCount)
                .location = BTrim(s.Cells(row, 1).value)
                .name = BTrim(s.Cells(row, 2).value)
                .Group = BTrim(s.Cells(row, 3).value)
                .color = s.Cells(row, 4).Interior.ColorIndex
            End With
            row = row + 1
        Wend
        
        
        row = 2
        CoursesCount = 0
        Set s = MasterData.Sheets("Courses")
        While BTrim(s.Cells(row, 1)) <> ""
            CoursesCount = CoursesCount + 1
            courses(CoursesCount) = BTrim(s.Cells(row, 1).value)
            row = row + 1
        Wend
        
        
        row = 2
        InstructorsCount = 0
        Set s = MasterData.Sheets("Instructors")
        While BTrim(s.Cells(row, 1)) <> ""
            InstructorsCount = InstructorsCount + 1
            instructors(InstructorsCount).name = BTrim(s.Cells(row, 1).value)
            instructors(InstructorsCount).qualifications = Split(s.Cells(row, 2).value, ",")
            
            row = row + 1
        Wend
        
    
        
        masterDataTimeStamp = Timer
        MasterData.Close False
    End If
    
End Sub

Public Function BTrim(txt As String) As String
    txt = Trim(txt)
    While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Wend
    Dim c As String
    'leading
    c = Left(txt, 1)
    While c = vbCr Or c = vbLf Or c = vbTab
        txt = Right(txt, Len(txt) - 1)
        c = Left(txt, 1)
    Wend
    
    'trailing
    c = Right(txt, 1)
    While c = vbCr Or c = vbLf Or c = vbTab
        txt = Right(txt, Len(txt) - 1)
        c = Right(txt, 1)
    Wend
    
    BTrim = txt
End Function





Public Function FacilityID2Name(ID As Integer) As String
    If ID < 1 Then
        FacilityID2Name = ""
        Exit Function
    End If
    initMasterData
    FacilityID2Name = facilities(ID).name
End Function

Public Function FacilityID2Color(ID As Integer) As Integer
    If ID < 1 Then
        FacilityID2Color = 0
        Exit Function
    End If
    initMasterData
    FacilityID2Color = facilities(ID).color
End Function


Public Function FacilityName2Color(name As String) As Integer
    Dim ID As String
    ID = FacilityName2ID(name)
    If ID > -1 Then
        FacilityName2Color = facilities(ID).color
    End If
End Function


Public Function Facility2GroupName(ID As Integer) As String
    If ID < 1 Then
        Facility2GroupName = ""
        Exit Function
    End If
    initMasterData
    Facility2GroupName = facilities(ID).Group
End Function


Public Function GetInstructors() As Variant
   initMasterData
    Dim i As Integer
    Dim res() As String
    ReDim res(1 To InstructorsCount)
    
    For i = 1 To InstructorsCount
            res(i) = instructors(i).name
    Next
    
    GetInstructors = res
End Function

Public Function InstructorHasQualifications(name As String, qualification As String) As Boolean
    Dim i As Integer, j As Integer
    InstructorHasQualifications = False
    
    For i = 1 To InstructorsCount
            If instructors(i).name = BTrim(name) Then
                For j = LBound(instructors(i).qualifications) To UBound(instructors(i).qualifications)
                    If InStr(instructors(i).qualifications(j), qualification) > 0 Then
                        InstructorHasQualifications = True
                        Exit Function
                    End If
                Next
                Exit Function
            End If
    Next
    
End Function

Public Function Instructor2ID(name As String) As Integer
    Dim i As Integer
    For i = 1 To InstructorsCount
            If instructors(i).name = BTrim(name) Then
                Instructor2ID = i
                Exit Function
            End If
    Next
    Instructor2ID = 0
End Function

Public Function GetFacilities(location As String) As Variant

    initMasterData
    Dim i As Integer
    Dim alloc As Boolean
    alloc = False
    Dim res() As String
    
    For i = 1 To FacilitiesCount
        If facilities(i).location = location Then
            If Not alloc Then
                ReDim res(1)
                alloc = True
            Else
                ReDim Preserve res(UBound(res) + 1)
            End If
            res(UBound(res)) = facilities(i).name

        End If
    Next
    
    GetFacilities = res
End Function


Public Function GetCourses() As Variant

    initMasterData
    Dim i As Integer
    Dim alloc As Boolean
    alloc = False
    Dim res() As String
    ReDim res(1 To CoursesCount)
    
    For i = 1 To CoursesCount
            res(i) = courses(i)
        
    Next
    
    GetCourses = res
End Function

Public Function FacilityName2ID(name As String) As Integer
    initMasterData
    Dim i As Integer
    Dim pos As Integer
    'take only first line:
    pos = InStr(name, vbLf)
    If (pos > 0) Then
        name = Left(name, pos - 1)
    End If

    name = BTrim(name)
    For i = 1 To FacilitiesCount
        If facilities(i).name = name Then
            'check if it is really a facility and not lunch etc.
            If facilities(i).location = "" Then
                FacilityName2ID = 0
            Else
                FacilityName2ID = i
            End If
            Exit Function
        End If
    
    Next
    FacilityName2ID = -1
End Function

Public Function CourseID2Name(ID As Integer) As String
    If ID < 1 Then
        CourseID2Name = ""
        Exit Function
    End If
    initMasterData
     CourseID2Name = courses(ID)
End Function

Public Function CourseName2ID(name As String) As Integer
   initMasterData
   Dim i As Integer
   name = BTrim(name)
   row = 1
    For i = 1 To CoursesCount
        If courses(i) = name Then
                CourseName2ID = i
                Exit Function
        End If
    Next
    CourseName2ID = -1
End Function

