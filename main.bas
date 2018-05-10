Attribute VB_Name = "Main"
Option Explicit
'Version 1.3


Public Const HEADER_ROW = 1
Public Const HOURS_START_ROW = 6
Public Const FACILITY_OFFSET = 100
Public Const ROW_SUMMARY_HOURS = 55
Public Const ROW_SUMMARY_INSTRUCTORS = 56

Public Const ROW_GUIDE_START = 40
Public Const GUIDES_COUNT = 11

Public Const SHARE_SIGN = "+"

Public Type TimeSlot
    length As Integer
    StartSlot As Integer
    ID As Integer
    SlotTitle As Variant
    Shared As Boolean
    color As Integer
End Type


Public Type total
    key As String
    value As Single
End Type

Public Type TotalList
    count As Integer
    Items(20) As total
End Type

Public Type Guide
    Group As String
    name As String
    color As Integer
    additionalColor As Integer
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



Public Type list
    count As Integer
    Items() As Item
End Type


Sub UpdateSheet(ByRef theSheet As Worksheet)
    Dim courses As list
    Dim facilities As list
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
    End If
      
    Facilities2UI facilities, theSheet
End Sub

Function GetQualifiedInstructors(in_list As Variant, qualification As String) As Variant
    Dim i As Integer, count As Integer
    Dim defQ As String
    Dim list() As String
    list = in_list
    
    defQ = GetParam("Default Q")
    
    For i = LBound(list) To UBound(list)
        If InStr(defQ, qualification) > 0 Or InstructorHasQualifications(list(i), qualification) Then
            count = count + 1
        Else
            list(i) = ""
        End If
    Next
    
    Dim res() As String
    ReDim res(1 To count + 1)
    count = 0
    
    For i = LBound(list) To UBound(list)
        If list(i) <> "" Then
            count = count + 1
            res(count) = list(i)
        End If
    Next
    
   GetQualifiedInstructors = res
    
    
End Function

Sub CalculateInstructorsTotals(ByRef courses As list)
    
    Dim i As Integer, j As Integer
    Dim slot As Integer, index As Integer
    For i = 1 To courses.count
        With courses.Items(i)
            For j = 1 To .GuideCount
                For slot = 1 To .SlotCount
                    If .Guides(j).color = .TimeSlots(slot).color Or _
                        .Guides(j).additionalColor = .TimeSlots(slot).color Then
                        index = getTotalIndex(.Guides(j).name, .Totals(2))
                        If index < 0 Then
                            .Totals(2).count = .Totals(2).count + 1
                            index = .Totals(2).count
                            .Totals(2).Items(index).key = .Guides(j).name
                        End If
                   
                        .Totals(2).Items(index).value = .Totals(2).Items(index).value + (.TimeSlots(slot).length / 2)
                   End If
                Next
            Next
        End With
    Next
    
    
End Sub


Sub CalculateFacilitiesTotals(ByRef courses As list)
    Dim groupName As String
    Dim i, j, index As Integer
    
    For i = 1 To courses.count
        For j = 1 To courses.Items(i).SlotCount
            With courses.Items(i)
                groupName = Facility2GroupName(.TimeSlots(j).ID)
                index = getTotalIndex(groupName, .Totals(1))
                If index < 0 Then
                    .Totals(1).count = .Totals(1).count + 1
                    index = .Totals(1).count
                    .Totals(1).Items(index).key = groupName
                End If
                
                .Totals(1).Items(index).value = .Totals(1).Items(index).value + .TimeSlots(j).length / 2 'each one is half hour
                
            End With
        Next
    Next
End Sub

Sub Totals2UI(ByRef theList As list, ByRef curr As Worksheet)
    Dim i As Integer, col As Integer, length As Integer
    Dim r As Range
    On Error Resume Next
    Application.ScreenUpdating = False
    'Clean totals
    col = 2
    curr.Range(curr.Cells(ROW_SUMMARY_HOURS, col), curr.Cells(ROW_SUMMARY_INSTRUCTORS, FACILITY_OFFSET)).Clear
    
   For i = 1 To theList.count
        length = getHowManyColToSkip(curr, col)

        Set r = getRange(curr, ROW_SUMMARY_HOURS, col, length, False)
        r.Cells(1, 1).value = getTotalString(theList.Items(i).Totals(1))
        
 
        r.Merge False
        r.HorizontalAlignment = xlRight
        r.VerticalAlignment = xlTop
        'r.Interior.ColorIndex = curr.Cells(ROW_SUMMARY_HOURS, 1).Interior.ColorIndex
        
        Set r = getRange(curr, ROW_SUMMARY_INSTRUCTORS, col, length, False)
        r.Cells(1, 1).value = getTotalString(theList.Items(i).Totals(2))
        r.Merge False
        r.HorizontalAlignment = xlRight
        r.VerticalAlignment = xlTop
        
        col = col + length
    Next
    
    Application.ScreenUpdating = True

End Sub

Function getTotalString(total As TotalList) As String
    Dim i As Integer
    For i = 1 To total.count
        
        getTotalString = getTotalString + IIf(Len(getTotalString) > 0, vbLf, "") + total.Items(i).key + ": " + CStr(total.Items(i).value)
    Next

End Function

Function getRange(ByRef curr As Worksheet, row As Integer, col As Integer, length As Integer, isVertical As Boolean) As Range

    Set getRange = curr.Range(curr.Cells(row, col), curr.Cells(row + _
         IIf(isVertical, length - 1, 0), col + IIf(Not isVertical, length - 1, 0)))
End Function

Function getHowManyColToSkip(curr As Worksheet, col As Integer) As Integer
    Dim r As Range
    Set r = curr.Cells(HEADER_ROW, col).MergeArea
    getHowManyColToSkip = r.Columns.count
End Function

Function getTotalIndex(key As String, Totals As TotalList) As Integer
    Dim i As Integer
    For i = 1 To Totals.count
        If Totals.Items(i).key = key Then
            getTotalIndex = i
            Exit Function
        End If
    Next
    getTotalIndex = -1
End Function

Sub MyReDim(ByRef l As list)

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
            While ma.Cells(1, 1).value <> ""
                col = col + ma.Columns.count
                 Set ma = curr.Cells(HEADER_ROW, col).MergeArea
            Wend
        
        Set r = Range(curr.Cells(, col + 1), curr.Cells(, FACILITY_OFFSET - 1))
        r.EntireColumn.Hidden = True

End Sub

Function UI2Courses(ByRef courses As list, ByRef curr As Worksheet) As String
    Dim col, row As Integer
    Dim headerVal As String, facilityName As String
    Dim ma As Range
    Dim SkipCols As Integer
    Dim isShared As Boolean
   For col = 2 To 100
        headerVal = curr.Cells(HEADER_ROW, col).value
        'Debug.Print headerVal
        
        Set ma = curr.Cells(HEADER_ROW, col).MergeArea
        If ma Is Nothing Then
            SkipCols = 0
        Else
            SkipCols = ma.Columns.count
        End If
        If headerVal = "" Then
            Exit For
        Else
            'Debug.Print "Header: " + headerVal
            MyReDim courses
            
            courses.count = courses.count + 1
            courses.Items(courses.count).ID = CourseName2ID(headerVal)
            If courses.Items(courses.count).ID < 0 Then
                UI2Courses = FormatString(3, headerVal)
                Exit Function
            End If
            courses.Items(courses.count).name = headerVal
             
            
           
            For row = HOURS_START_ROW To HOURS_START_ROW + 31
                Set ma = curr.Cells(row, col).MergeArea
                If ma.row = row Then
                   facilityName = BTrim(ma.Cells(1, 1))
                   If Left(facilityName, 1) = SHARE_SIGN Then
                       facilityName = Right(facilityName, Len(facilityName) - 1)
                       isShared = True
                   Else
                       isShared = False
                   End If
                   
                  If (facilityName <> "" And Left(facilityName, 1) <> "*") Then
                       'add new slot to course
                       Dim slot As TimeSlot
                  
                       slot.ID = FacilityName2ID(facilityName)
                       slot.SlotTitle = facilityName
                       slot.Shared = isShared
                       If (slot.ID = -1) Then
                           UI2Courses = FormatString(1, facilityName, headerVal, CStr(row))
                           Exit Function
                       End If
                       If slot.ID > 0 Then
                           slot.length = ma.Rows.count
                            If IsNull(ma.Interior.ColorIndex) Then
                                ma.UnMerge
                                ma.Merge False
                            End If
                            If Not IsNull(ma.Interior.ColorIndex) Then
                                slot.color = ma.Interior.ColorIndex
                            Else
                                slot.color = ma.Interior.color
                            End If
                            
                            slot.StartSlot = row - HOURS_START_ROW
                           With courses.Items(courses.count)
                               .SlotCount = .SlotCount + 1
                               .TimeSlots(.SlotCount) = slot
                               
                           End With
                           'Debug.Print "Facility " + facilityName + " (" + CStr(slot.ID) + ") :" + CStr(ma.Rows.Count)
                       End If
                       'HebMsgBox FormatString(1, ma.Cells(1, 1), CStr(ma.Cells.Count))
       
    
                       row = row + ma.Rows.count - 1
                   End If
                End If
            Next
            
            Dim name As String
            Dim names() As String
            Dim N As Integer
            
            'extract Guides
            Dim defQ As String
            defQ = GetParam("Default Q")
            Dim errStr As String
            For row = ROW_GUIDE_START To ROW_GUIDE_START + GUIDES_COUNT
                Set ma = curr.Cells(row, col).MergeArea
                name = ma.Cells(1, 1)
                If name <> "" Then
                    names = Split(name, ",")
                    For N = LBound(names) To UBound(names)
                        ' verify the Guide is qualified
                        With courses.Items(courses.count)
                            .GuideCount = .GuideCount + 1
                            .Guides(.GuideCount).name = BTrim(names(N))
                            
                            
                            
                            If Not InStr(defQ, BTrim(curr.Cells(row, 1))) > 0 And _
                                Not InstructorHasQualifications(.Guides(.GuideCount).name, BTrim(curr.Cells(row, 1))) Then
                                If Instructor2ID(.Guides(.GuideCount).name) = 0 Then
                                    'guide does not exits
                                    errStr = addErr(errStr, FormatString(10, .Guides(.GuideCount).name))
                                Else
                                    errStr = addErr(errStr, FormatString(11, .Guides(.GuideCount).name, curr.Cells(row, 1)))
                                End If
                            End If
                            
                            .Guides(.GuideCount).color = curr.Cells(row, 1).Interior.ColorIndex
                            .Guides(.GuideCount).additionalColor = getAdditionalColor(.Guides(.GuideCount).color)
                        End With
                    Next
                End If
            Next
            
             
            
        End If
        col = col + SkipCols - 1
    Next
    
    If Len(errStr) > 0 Then
      HebMsgBox errStr
    End If
    
    
End Function

Function getAdditionalColor(color As Integer) As Integer
    Dim addColor As String
    addColor = GetParam("AdditionalColor" + CStr(color))
    If Len(addColor) > 0 Then
        getAdditionalColor = CInt(addColor)
        Exit Function
    End If
    getAdditionalColor = -1
End Function

 Function addErr(ByRef str As String, newErr As String)
    If InStr(str, newErr) <= 0 Then
        addErr = str + IIf(Len(str) > 0, vbLf, "") + newErr
    Else
        addErr = str
    End If
 End Function

Sub Facilities2UI(ByRef facilities As list, curr As Worksheet)
    Dim col As Integer
    On Error Resume Next
    Application.ScreenUpdating = False

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
        facilityRange.Cells(HEADER_ROW, col).value = facHeaders(col)
    Next
    
    
    Dim slotInx As Integer
    Dim timeSlotRange As Range
    Dim i As Integer, j As Integer
    'print to sheet the facility
    For i = 1 To facilities.count
        col = -1
        For j = 1 To UBound(facHeaders)
            If facHeaders(j) = facilities.Items(i).name Then
                col = j
                Exit For
            End If
        Next
    
        If col = -1 Then
            HebMsgBox FormatString(15, facilities.Items(i).name, location)
            Exit Sub
        End If
        
        'facilityRange.Cells(HEADER_ROW, col).Value = facilities.Items(col).Name
        
        For slotInx = 1 To facilities.Items(i).SlotCount
            With facilities.Items(i).TimeSlots(slotInx)
                Set timeSlotRange = facilityRange.Range( _
                    facilityRange.Parent.Cells(HOURS_START_ROW + .StartSlot, col), _
                    facilityRange.Parent.Cells(HOURS_START_ROW + .StartSlot + .length - 1, col))
                addFacility2Cell timeSlotRange.Cells(1, 1), CourseID2Name(.ID)
                 
                'timeSlotRange.Select
                 formatRange timeSlotRange, .color
                
            End With
            
        Next
    Next
   Application.ScreenUpdating = True


End Sub


Sub addFacility2Cell(cell As Range, value As String)
    If Len(cell.value) > 0 Then
        cell.value = cell.value + vbLf + "+" + vbLf + value
        cell.AddComment FormatString(14)
    Else
        cell.value = value
    End If
End Sub

Sub formatRange(r As Range, color As Integer)
    On Error Resume Next
    Application.DisplayAlerts = False
    Dim isOverlapRange As Boolean
    With r
        isOverlapRange = isOverlap(r)
        
        
        If Not isOverlapRange Then
            .Merge False
        End If
         .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = color
        .WrapText = True
        If isOverlapRange Then
            .Cells(1, 1).Borders.ColorIndex = 3
            .Cells(1, 1).Borders.Weight = xlThick
        Else
            .Borders.ColorIndex = 0
            .Borders.Weight = xlThin
        End If
        .Borders.LineStyle = xlContinuous
        
        
        If InStr(.Cells(1, 1), "+") > 0 Then
            .Font.Size = 6
        End If
        
    End With
     
     
    Application.DisplayAlerts = True
    
End Sub

Function isOverlap(r As Range) As Boolean
    Dim i As Integer, j As Integer, count As Integer
    
    For i = 1 To r.Rows.count
        For j = 1 To r.Columns.count
            If BTrim(r.Cells(i, j).value) <> "" Then
                count = count + 1
            End If
        Next
    Next
    isOverlap = (count > 1)
End Function

Function Courses2Facilities(ByRef courses As list, ByRef facilities As list) As String
    Dim i As Integer, j As Integer, facilityIndex, k As Integer
    Dim newSlot As TimeSlot
    Dim errStr As String
    
    For i = 1 To courses.count
        For j = 1 To courses.Items(i).SlotCount
            facilityIndex = getFacilityIndex(facilities, courses.Items(i).TimeSlots(j).ID)
            If facilityIndex = 0 Then 'not found
                MyReDim facilities
                facilities.count = facilities.count + 1
                facilityIndex = facilities.count
            End If
             
             With facilities.Items(facilityIndex)
             
                'makes sure no conflicting slots
                newSlot = courses.Items(i).TimeSlots(j)
                For k = 1 To .SlotCount
                    With .TimeSlots(k)
                        's = .TimeSlots(k)
                        If ((newSlot.StartSlot >= .StartSlot And newSlot.StartSlot < .StartSlot + .length) Or _
                           (newSlot.StartSlot + newSlot.length > .StartSlot And newSlot.StartSlot + newSlot.length <= .StartSlot + .length)) Then
                            If newSlot.StartSlot = .StartSlot And newSlot.length = .length And (newSlot.Shared Or .Shared) Then
                                'same start time and same length and one is defined shared. so it is OK
                            Else
                                'conflic
                                errStr = addErr(errStr, FormatString(2, courses.Items(i).name, FacilityID2Name(newSlot.ID), CourseID2Name(.ID)))
                            End If
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
    
    Courses2Facilities = errStr
    
End Function

Function getHour(index As Integer) As String
    getHour = ""
End Function

Function getFacilityIndex(facilities As list, facilityID As Integer) As Integer
    Dim i As Integer
    For i = 1 To facilities.count
        If facilities.Items(i).ID = facilityID Then
            getFacilityIndex = i
            Exit Function
        End If
    Next
    getFacilityIndex = 0
        
End Function





