VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectionForm 
   Caption         =   "מסך בחירה"
   ClientHeight    =   6450
   ClientLeft      =   40
   ClientTop       =   400
   ClientWidth     =   8040
   OleObjectBlob   =   "SelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_currCell As Range
Private m_items() As String
Private m_functionPtr As String
Private m_multiSelect As Boolean

'Version 1.1

Sub Clear()
    lsValues.Clear
End Sub

Public Sub InitOnce(txtAdd As String, title As String, isSharedLbl As String)
    Me.Caption = "Select..."
    lblAdditionalText = txtAdd
    chkShare.Caption = isSharedLbl
End Sub

Public Sub Load(title As String, functionPtr As String, value As String, values As Variant, displayShare As Boolean, multiSelect As Boolean)
    Clear
    SetTitle title
    m_functionPtr = functionPtr
    m_multiSelect = multiSelect

    lsValues.ColumnCount = 2
    lsValues.ColumnWidths = "90;10"
    
    lsValues.multiSelect = IIf(multiSelect, fmMultiSelectMulti, fmMultiSelectSingle)
    
    m_items = values
    chkShare.Visible = displayShare
    
    SetValue value
    FilterItems
    Me.Show
End Sub

 Sub SetTitle(txt As String)
     lblItem = txt
End Sub


Sub FilterItems()

    lsValues.Clear
    
    Dim val() As String
    If m_multiSelect And Len(txtItem.text) > 0 Then
     val = Split(BTrim(txtItem.text), ",")
    Else
        ReDim val(1 To 1)
        val(1) = BTrim(txtItem.text)
    End If
    
    For i = LBound(m_items) To UBound(m_items)
        For j = LBound(val) To UBound(val)
            If (InStr(BTrim(m_items(i)), BTrim(val(j))) = 1) Then
                lsValues.AddItem m_items(i)
                If BTrim(m_items(i)) = BTrim(val(j)) Then
                    'exact match - select the item
                    lsValues.Selected(lsValues.ListCount - 1) = True
                End If
                Exit For
            End If
        Next
    Next
 
End Sub

 Sub SetValue(value As String)
    Dim pos As Integer
    Dim val As String
    val = value
    
    If Left(val, 1) = SHARE_SIGN Then
        chkShare.value = True
        val = Right(val, Len(val) - 1)
    Else
        chkShare.value = False
    End If
    
    If Left(val, 1) = "*" Then
        txtItem = ""
        txtAdditionalText = Right(val, Len(val) - 1)
    Else
        
        pos = InStr(val, vbLf)
        If pos > 0 Then
            txtItem = BTrim(Left(val, pos - 1))
            txtAdditionalText = BTrim(Right(val, Len(val) - pos))
        Else
            txtItem = BTrim(val)
            txtAdditionalText = ""
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim val As String
    
    If m_multiSelect Then
        For i = 0 To lsValues.ListCount - 1
            If lsValues.Selected(i) Then
                val = val + IIf(Len(val) > 0, ", ", "") + lsValues.list(i)
            End If
        Next
    Else
        val = BTrim(txtItem.text)
    End If

    
    Application.Run m_functionPtr, val, chkShare.value, BTrim(txtAdditionalText)
   
    Me.Hide
End Sub



Private Sub lsValues_Click()
    txtItem = lsValues.list(lsValues.ListIndex)
End Sub

Private Sub txtItem_Change()
 FilterItems
End Sub

