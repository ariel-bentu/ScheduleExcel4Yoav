Attribute VB_Name = "MasterData"
'MASTER DATA

Type Facility
    Name As String
    Location As String
    Group As String
    ID As Integer
End Type


Dim MasterData As Workbook
Private facilities(200) As Facility
Private FacilitiesCount As Integer
Private courses(200) As String
Private CoursesCount As Integer

Dim masterDataTimeStamp As Single


Sub initMasterData()
    'cache for 60 seconds
    If masterDataTimeStamp + 60 > Timer Then Exit Sub

    Dim s As Worksheet, row As Integer
    If MasterData Is Nothing Then
    
        Application.ScreenUpdating = False
        On Error Resume Next
   
    
        Set MasterData = Workbooks.Open(ActiveWorkbook.Path & "/yoav-masterdata.xlsx", , True)
        MasterData.Windows(1).Visible = False
        Application.ScreenUpdating = True
        row = 2
        Set s = MasterData.Sheets("Facilities")
        While s.Cells(row, 2) <> ""
            With facilities(row - 1)
                .Location = s.Cells(row, 1).Value
                .Name = s.Cells(row, 2).Value
                .Group = s.Cells(row, 3).Value
            End With
            row = row + 1
        Wend
        FacilitiesCount = row - 1
        
        
       row = 1
        Set s = MasterData.Sheets("Courses")
        While s.Cells(row, 1) <> ""
            courses(row) = s.Cells(row, 1).Value
            row = row + 1
        Wend
        CoursesCount = row
    End If
    masterDataTimeStamp = Timer
    MasterData.Close
    
End Sub

Public Function FacilityID2Name(ID As Integer) As String
    If ID < 1 Then
        FacilityID2Name = ""
        Exit Function
    End If
    initMasterData
    FacilityID2Name = facilities(ID).Name
End Function

Public Function FacilityName2ID(Name As String) As Integer
    initMasterData
    Dim i As Integer
    Dim pos As Integer
    'take only first line:
    pos = InStr(Name, vbLf)
    If (pos > 0) Then
        Name = Left(Name, pos - 1)
    End If

    Name = Trim(Name)
    For i = 1 To FacilitiesCount
        If facilities(i).Name = Name Then
            'check if it is really a facility and not lunch etc.
            If facilities(i).Location = "" Then
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

Public Function CourseName2ID(Name As String) As Integer
   initMasterData
   Dim i As Integer
   Name = Trim(Name)
   row = 1
    For i = 1 To CoursesCount
        If courses(i) = Name Then
                CourseName2ID = i
                Exit Function
        End If
    Next
    CourseName2ID = -1
End Function

