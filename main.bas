Attribute VB_Name = "Main"
Option Explicit

Const HEADER_ROW = 1
Const HOURS_START_ROW = 4
Public Const FACILITY_OFFSET = 20
Const ROW_SUMMARY_HOURS = 50
Const ROW_SUMMARY_GUIDES = 51



Public Type TimeSlot
    Length As Integer
    StartSlot As Integer
    ID As Integer
    SlotTitle As Variant
    Color As Integer
End Type

Public Type Item
    ID As Integer
    Name As String
    TimeSlots(1 To 32) As TimeSlot
    SlotCount As Integer
End Type


Public Type List
    Count As Integer
    Items(20) As Item
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
    
    errStr = Courses2Facilities(courses, facilities)
    If errStr <> "" Then
        HebMsgBox errStr
        Exit Sub
    End If
      
    Facilities2UI facilities, theSheet
End Sub

Function UI2Courses(ByRef courses As List, ByRef curr As Worksheet) As String
    Dim col, row As Integer
    Dim headerVal As String, facilityName As String
    Dim ma As Range
    Dim SkipCols As Integer
    
   For col = 2 To 100
        headerVal = curr.Cells(HEADER_ROW, col).Value
        Debug.Print headerVal
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
            
            courses.Count = courses.Count + 1
            courses.Items(courses.Count).ID = CourseName2ID(headerVal)
            If courses.Items(courses.Count).ID < 0 Then
                UI2Courses = FormatString(3, headerVal)
                Exit Function
            End If
            courses.Items(courses.Count).Name = headerVal
             
            
           
            For row = HOURS_START_ROW To HOURS_START_ROW + 32
                Set ma = curr.Cells(row, col).MergeArea
                facilityName = ma.Cells(1, 1)
                'Debug.Print "row: " + CStr(row) + "   col:" + CStr(col) + " Name:" + facilityName
                If (facilityName <> "") Then
                    'add new slot to course
                    Dim slot As TimeSlot
                    slot.ID = FacilityName2ID(facilityName)
                    slot.SlotTitle = facilityName
                    If (slot.ID = -1) Then
                        UI2Courses = FormatString(1, facilityName)
                        Exit Function
                    End If
                    If slot.ID > 0 Then
                        slot.Length = ma.Rows.Count
                        slot.Color = ma.Interior.ColorIndex
                        
                        slot.StartSlot = row - HOURS_START_ROW
                        With courses.Items(courses.Count)
                            .SlotCount = .SlotCount + 1
                            .TimeSlots(.SlotCount) = slot
                            
                        End With
                        Debug.Print "Facility " + facilityName + " (" + CStr(slot.ID) + ") :" + CStr(ma.Rows.Count)
                    End If
                    'HebMsgBox FormatString(1, ma.Cells(1, 1), CStr(ma.Cells.Count))
    
 
                    row = row + ma.Rows.Count - 1
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
    
    Dim slotInx As Integer
    Dim timeSlotRange As Range
    'print to sheet the facility
    For col = 1 To facilities.Count
        facilityRange.Cells(HEADER_ROW, col).Value = facilities.Items(col).Name
        
        For slotInx = 1 To facilities.Items(col).SlotCount
            With facilities.Items(col).TimeSlots(slotInx)
                Set timeSlotRange = facilityRange.Range(facilityRange.Parent.Cells(HOURS_START_ROW + .StartSlot, col), facilityRange.Parent.Cells(HOURS_START_ROW + .StartSlot + .Length - 1, col))
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
                facilities.Count = facilities.Count + 1
                facilityIndex = facilities.Count
            End If
             
             With facilities.Items(facilityIndex)
             
                'makes sure no conflicting slots
                newSlot = courses.Items(i).TimeSlots(j)
                For k = 1 To .SlotCount
                    With .TimeSlots(k)
                        's = .TimeSlots(k)
                        If (newSlot.StartSlot >= .StartSlot And newSlot.StartSlot <= .StartSlot + .Length) Or _
                           (newSlot.StartSlot + newSlot.Length >= .StartSlot And newSlot.StartSlot + newSlot.Length <= .StartSlot + .Length) Then
                            'conflict
                             Courses2Facilities = FormatString(2, courses.Items(i).Name, FacilityID2Name(newSlot.ID), CourseID2Name(.ID))
                             Exit Function
                        End If
                    End With
                Next
             
             
                 .ID = courses.Items(i).TimeSlots(j).ID
                 .Name = FacilityID2Name(courses.Items(i).TimeSlots(j).ID)
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

