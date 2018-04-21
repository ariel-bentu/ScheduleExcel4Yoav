VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Message Box"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   500
   ClientWidth     =   12880
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdOK_Click()
   Me.Hide
End Sub

Public Sub SetText(text As String)
    txtMessage.Value = text
End Sub


Private Sub UserForm_Activate()
    Me.Width = 600
    Me.RightToLeft = True
    Me.Caption = "Message"

   txtMessage.Height = 40
    txtMessage.Width = 500
    txtMessage.Font.Size = 14
    
    txtMessage.TextAlign = fmTextAlignRight
    txtMessage.BorderStyle = fmBorderStyleNone
End Sub

