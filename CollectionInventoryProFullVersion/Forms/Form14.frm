VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4425
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5970
   ControlBox      =   0   'False
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form14.frx":5C12
   ScaleHeight     =   4425
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
frmLogin.Show vbModal

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
frmLogin.Show vbModal

End Sub

