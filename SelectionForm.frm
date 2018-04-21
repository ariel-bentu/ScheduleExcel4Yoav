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


Sub Clear()
    lsValues.Clear
End Sub

Public Sub InitOnce(txtAdd As String, title As String)
    Me.Caption = "Select..."
    lblAdditionalText = txtAdd
End Sub

Public Sub Load(title As String, targetCell As Range, values As Variant)
    Clear
    SetTitle title
    
    m_items = values
    
    
    SetCellTarget targetCell
    FilterItems
    Me.Show
End Sub

 Sub SetTitle(txt As String)
     lblItem = txt
End Sub


Sub FilterItems()

    lsValues.Clear
    For i = LBound(m_items) To UBound(m_items)
        If (InStr(Trim(m_items(i)), Trim(txtItem.text)) = 1) Then
            lsValues.AddItem m_items(i)
        End If
    Next
 
End Sub

 Sub SetCellTarget(cell As Range)
    Set m_currCell = cell
    Dim pos As Integer
    
    If Left(cell.Value, 1) = "*" Then
        txtItem = ""
        txtAdditionalText = Right(cell.Value, Len(cell.Value) - 1)
    Else
        
        pos = InStr(cell.Value, vbLf)
        If pos > 0 Then
            txtItem = Trim(Left(cell.Value, pos - 1))
            txtAdditionalText = Trim(Right(cell.Value, Len(cell.Value) - pos))
        Else
            txtItem = Trim(cell.Value)
            txtAdditionalText = ""
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Selection.Rows.Count > 1 Then
        Selection.Merge False
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
    End If

    m_currCell.Value = IIf(Len(Trim(txtItem)) > 0, Trim(txtItem), "")
    If (Len(Trim(txtAdditionalText)) > 0) Then
        If Len(m_currCell.Value) > 0 Then
            m_currCell.Value = m_currCell.Value + vbLf
        Else
            m_currCell.Value = "*"
        End If
        m_currCell.Value = m_currCell.Value + Trim(txtAdditionalText)
    End If
   
    Me.Hide
End Sub



Private Sub lsValues_Click()
    txtItem = lsValues.List(lsValues.ListIndex)
End Sub

Private Sub txtItem_Change()
 FilterItems
End Sub
