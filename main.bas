Attribute VB_Name = "Main"
Option Explicit

Public Const HEADER_ROW = 1
Public Const HOURS_START_ROW = 6
Public Const FACILITY_OFFSET = 100
Public Const ROW_SUMMARY_HOURS = 55
Public Const ROW_SUMMARY_INSTRUCTORS = 56

Public Const ROW_GUIDE_START = 40
Public Const GUIDES_COUNT = 9

Public Type TimeSlot
    length As Integer
    StartSlot As Integer
    ID As Integer
    SlotTitle As Variant
    Color As Integer
End Type


Public Type total
    key As String
    Value As Single
End Type

Public Type TotalList
    Count As Integer
    Items(20) As total
End Type

Public Type Guide
    Group As String
    name As String
    Color As Integer
End Type

Public Type Item
    ID As Integer
    name As String
    TimeSlots(1 To 32) As TimeSlot
    SlotCount As Integer
    Guides(10) As Guide
    GuideCount As Integer
    Totals(2) As TotalList
End Type



Public Type List
    Count As Integer
    Items() As Item
End Type


Sub UpdateSheet(ByRef theSheet As Worksheet)
    Dim courses As List
    Dim facilities As List
    Dim errStr As String
    
    errStr = UI2Courses(courses, theSheet)
    If errStr <> "" Then
        HebMsgBox errStr
        Exit Sub
    End If
    
    Dim totalsFacilities As TotalList
    CalculateFacilitiesTotals courses
    CalculateInstructorsTotals courses
    
    Totals2UI courses, theSheet
    
    errStr = Courses2Facilities(courses, facilities)
    If errStr <> "" Then
        HebMsgBox errStr
        Exit Sub
    End If
      
    Facilities2UI facilities, theSheet
End Sub

Sub CalculateInstructorsTotals(ByRef courses As List)
    'each group has one or more guides comma seperated
    'for each guide, find the
    
    
    
    Dim i As Integer, j As Integer
    Dim slot As Integer, index As Integer
    For i = 1 To courses.Count
        With courses.Items(i)
            For j = 1 To .GuideCount
                For slot = 1 To .SlotCount
                    If .Guides(j).Color = .TimeSlots(slot).Color Then
                        index = getTotalIndex(.Guides(j).name, .Totals(2))
                        If index < 0 Then
                            .Totals(2).Count = .Totals(2).Count + 1
                            index = .Totals(2).Count
                            .Totals(2).Items(index).key = .Guides(j).name
                        End If
                   
                        .Totals(2).Items(index).Value = .Totals(2).Items(index).Value + (.TimeSlots(slot).length / 2)
                   End If
                Next
            Next
        End With
    Next
    
    
End Sub


Sub CalculateFacilitiesTotals(ByRef courses As List)
    Dim groupName As String
    Dim i, j, index As Integer
    
    For i = 1 To courses.Count
        For j = 1 To courses.Items(i).SlotCount
            With courses.Items(i)
                groupName = Facility2GroupName(.TimeSlots(j).ID)
                index = getTotalIndex(groupName, .Totals(1))
                If index < 0 Then
                    .Totals(1).Count = .Totals(1).Count + 1
                    index = .Totals(1).Count
                    .Totals(1).Items(index).key = groupName
                End If
                
                .Totals(1).Items(index).Value = .Totals(1).Items(index).Value + .TimeSlots(j).length / 2 'each one is half hour
                
            End With
        Next
    Next
End Sub

Sub Totals2UI(ByRef theList As List, ByRef curr As Worksheet)
    Dim i As Integer, col As Integer, length As Integer
    Dim r As Range
    'Clean totals
    col = 2
    curr.Range(curr.Cells(ROW_SUMMARY_HOURS, col), curr.Cells(ROW_SUMMARY_INSTRUCTORS, FACILITY_OFFSET)).Clear
    
   For i = 1 To theList.Count
        length = getHowManyColToSkip(curr, col)
        Set r = getRange(curr, ROW_SUMMARY_HOURS, col, length, False)
        r.Cells(1, 1).Value = getTotalString(theList.Items(i).Totals(1))
        
 
        r.Merge False
        r.HorizontalAlignment = xlRight
        r.VerticalAlignment = xlTop
        'r.Interior.ColorIndex = curr.Cells(ROW_SUMMARY_HOURS, 1).Interior.ColorIndex
        
        Set r = getRange(curr, ROW_SUMMARY_INSTRUCTORS, col, length, False)
        r.Cells(1, 1).Value = getTotalString(theList.Items(i).Totals(2))
        r.Merge False
        r.HorizontalAlignment = xlRight
        r.VerticalAlignment = xlTop
        
        col = col + length
    Next
End Sub

Function getTotalString(total As TotalList) As String
    Dim i As Integer
    For i = 1 To total.Count
        
        getTotalString = getTotalString + IIf(Len(getTotalString) > 0, vbLf, "") + total.Items(i).key + ": " + CStr(total.Items(i).Value)
    Next

End Function

Function getRange(ByRef curr As Worksheet, row As Integer, col As Integer, length As Integer, isVertical As Boolean) As Range

    Set getRange = curr.Range(curr.Cells(row, col), curr.Cells(row + _
         IIf(isVertical, length - 1, 0), col + IIf(Not isVertical, length - 1, 0)))
End Function

Function getHowManyColToSkip(curr As Worksheet, col As Integer) As Integer
    Dim r As Range
    Set r = curr.Cells(HEADER_ROW, col).MergeArea
    getHowManyColToSkip = r.Columns.Count
End Function

Function getTotalIndex(key As String, Totals As TotalList) As Integer
    Dim i As Integer
    For i = 1 To Totals.Count
        If Totals.Items(i).key = key Then
            getTotalIndex = i
            Exit Function
        End If
    Next
    getTotalIndex = -1
End Function

Sub MyReDim(ByRef l As List)

    Dim newSize As Integer
    On Error Resume Next
    newSize = UBound(l.Items) + 1
    If Err.Number <> 0 Then
        newSize = 1
    End If
    ReDim Preserve l.Items(newSize)

End Sub


Public Sub UnhideGAP(ByRef curr As Worksheet)
        Dim r As Range
        'Set r = Range(curr.Cells(, 10), curr.Cells(, FACILITY_OFFSET - 1))
        'r.EntireColumn.Hidden = False
End Sub
Public Sub HideGAP(ByRef curr As Worksheet)
        
        Dim r As Range
        Exit Sub
        'find first empty col
        
        Dim col As Integer
        Dim ma As Range
        col = 1
        Set ma = curr.Cells(HEADER_ROW, col).MergeArea
            While ma.Cells(1, 1).Value <> ""
                col = col + ma.Columns.Count
                 Set ma = curr.Cells(HEADER_ROW, col).MergeArea
            Wend
        
        Set r = Range(curr.Cells(, col + 1), curr.Cells(, FACILITY_OFFSET - 1))
        r.EntireColumn.Hidden = True

End Sub

Function UI2Courses(ByRef courses As List, ByRef curr As Worksheet) As String
    Dim col, row As Integer
    Dim headerVal As String, facilityName As String
    Dim ma As Range
    Dim SkipCols As Integer
    
   For col = 2 To 100
        headerVal = curr.Cells(HEADER_ROW, col).Value
        'Debug.Print headerVal
        
        Set ma = curr.Cells(HEADER_ROW, col).MergeArea
        If ma Is Nothing Then
            SkipCols = 0
        Else
            SkipCols = ma.Columns.Count
        End If
        If headerVal = "" Then
            Exit For
        Else
            'Debug.Print "Header: " + headerVal
            MyReDim courses
            
            courses.Count = courses.Count + 1
            courses.Items(courses.Count).ID = CourseName2ID(headerVal)
            If courses.Items(courses.Count).ID < 0 Then
                UI2Courses = FormatString(3, headerVal)
                Exit Function
            End If
            courses.Items(courses.Count).name = headerVal
             
            
           
            For row = HOURS_START_ROW To HOURS_START_ROW + 31
                Set ma = curr.Cells(row, col).MergeArea
                facilityName = Trim(ma.Cells(1, 1))
                'Debug.Print "row: " + CStr(row) + "   col:" + CStr(col) + " Name:" + facilityName
                If (facilityName <> "" And Left(facilityName, 1) <> "*") Then
                    'add new slot to course
                    Dim slot As TimeSlot
                    slot.ID = FacilityName2ID(facilityName)
                    slot.SlotTitle = facilityName
                    If (slot.ID = -1) Then
                        UI2Courses = FormatString(1, facilityName)
                        Exit Function
                    End If
                    If slot.ID > 0 Then
                        slot.length = ma.Rows.Count
                        slot.Color = ma.Interior.ColorIndex
                        
                        slot.StartSlot = row - HOURS_START_ROW
                        With courses.Items(courses.Count)
                            .SlotCount = .SlotCount + 1
                            .TimeSlots(.SlotCount) = slot
                            
                        End With
                        'Debug.Print "Facility " + facilityName + " (" + CStr(slot.ID) + ") :" + CStr(ma.Rows.Count)
                    End If
                    'HebMsgBox FormatString(1, ma.Cells(1, 1), CStr(ma.Cells.Count))
    
 
                    row = row + ma.Rows.Count - 1
                End If
  
            Next
            
            Dim name As String
            Dim names() As String
            Dim n As Integer
            
            'extract Guides
            For row = ROW_GUIDE_START To ROW_GUIDE_START + GUIDES_COUNT
                Set ma = curr.Cells(row, col).MergeArea
                name = ma.Cells(1, 1)
                If name <> "" Then
                    names = Split(name, ",")
                    For n = LBound(names) To UBound(names)
                        ' verify the Guide is qualified - todo
                        With courses.Items(courses.Count)
                            .GuideCount = .GuideCount + 1
                            .Guides(.GuideCount).name = Trim(names(n))
                            
                            
                            
                            If Not InstructorHasQualifications(.Guides(.GuideCount).name, Trim(curr.Cells(row, 1))) Then
                                If Instructor2ID(.Guides(.GuideCount).name) = 0 Then
                                    'guide does not exits
                                    HebMsgBox FormatString(10, .Guides(.GuideCount).name)
                                Else
                                    HebMsgBox FormatString(11, .Guides(.GuideCount).name, curr.Cells(row, 1))
                                End If
                            End If
                            
                            .Guides(.GuideCount).Color = curr.Cells(row, 1).Interior.ColorIndex
                        End With
                    Next
                End If
            Next
            
            
        End If
        col = col + SkipCols - 1
    Next
    
End Function

 

Sub Facilities2UI(ByRef facilities As List, curr As Worksheet)
    Dim col As Integer

   'cleanup facility
    Dim facilityRange As Range
    With curr
        Set facilityRange = .Range(.Cells(1, FACILITY_OFFSET), .Cells(100, FACILITY_OFFSET + 100))
    End With
    facilityRange.Clear
    facilityRange.UnMerge
    
    
    'put all facilities
    Dim facHeaders() As String
    Dim location As String
    location = GetParam("Location")
    facHeaders = GetFacilities(location)
    
    
    For col = 1 To UBound(facHeaders)
             facilityRange.Cells(HEADER_ROW, col).Value = facHeaders(col)
    Next
    
    
    Dim slotInx As Integer
    Dim timeSlotRange As Range
    Dim i As Integer, j As Integer
    'print to sheet the facility
    For i = 1 To facilities.Count
    
          For j = 1 To UBound(facHeaders)
            If facHeaders(j) = facilities.Items(i).name Then
                col = j
                Exit For
            End If
        Next
    
        
        'facilityRange.Cells(HEADER_ROW, col).Value = facilities.Items(col).Name
        
        For slotInx = 1 To facilities.Items(i).SlotCount
            With facilities.Items(i).TimeSlots(slotInx)
                Set timeSlotRange = facilityRange.Range( _
                    facilityRange.Parent.Cells(HOURS_START_ROW + .StartSlot, col), _
                    facilityRange.Parent.Cells(HOURS_START_ROW + .StartSlot + .length - 1, col))

                timeSlotRange.Cells(1, 1).Value = CourseID2Name(.ID)
                'timeSlotRange.Select
                timeSlotRange.Merge False
                timeSlotRange.HorizontalAlignment = xlCenter
                timeSlotRange.VerticalAlignment = xlCenter
                timeSlotRange.Interior.ColorIndex = .Color
                
            End With
            
        Next
    Next
 
End Sub



Function Courses2Facilities(ByRef courses As List, ByRef facilities As List) As String
    Dim i As Integer, j As Integer, facilityIndex, k As Integer
    Dim newSlot As TimeSlot
    For i = 1 To courses.Count
        For j = 1 To courses.Items(i).SlotCount
            facilityIndex = getFacilityIndex(facilities, courses.Items(i).TimeSlots(j).ID)
            If facilityIndex = 0 Then 'not found
                MyReDim facilities
                facilities.Count = facilities.Count + 1
                facilityIndex = facilities.Count
            End If
             
             With facilities.Items(facilityIndex)
             
                'makes sure no conflicting slots
                newSlot = courses.Items(i).TimeSlots(j)
                For k = 1 To .SlotCount
                    With .TimeSlots(k)
                        's = .TimeSlots(k)
                        If (newSlot.StartSlot >= .StartSlot And newSlot.StartSlot <= .StartSlot + .length) Or _
                           (newSlot.StartSlot + newSlot.length >= .StartSlot And newSlot.StartSlot + newSlot.length <= .StartSlot + .length) Then
                            'conflict
                             Courses2Facilities = FormatString(2, courses.Items(i).name, FacilityID2Name(newSlot.ID), CourseID2Name(.ID))
                             Exit Function
                        End If
                    End With
                Next
             
             
                 .ID = courses.Items(i).TimeSlots(j).ID
                 .name = FacilityID2Name(courses.Items(i).TimeSlots(j).ID)
                 .SlotCount = .SlotCount + 1
                 .TimeSlots(.SlotCount) = courses.Items(i).TimeSlots(j)
                 'fix id to be course id instead of facility id
                 .TimeSlots(.SlotCount).ID = courses.Items(i).ID
             End With
        
        Next
    Next
    
    
End Function

Function getHour(index As Integer) As String
    getHour = ""
End Function

Function getFacilityIndex(facilities As List, facilityID As Integer) As Integer
    Dim i As Integer
    For i = 1 To facilities.Count
        If facilities.Items(i).ID = facilityID Then
            getFacilityIndex = i
            Exit Function
        End If
    Next
    getFacilityIndex = 0
        
End Function



