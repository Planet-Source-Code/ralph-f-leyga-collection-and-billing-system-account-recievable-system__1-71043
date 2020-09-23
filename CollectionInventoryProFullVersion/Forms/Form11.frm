VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Region"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1920
      Picture         =   "Form11.frx":5C12
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      Picture         =   "Form11.frx":608B
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Region:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
If Text1.Text <> "" Then
Set rs = New ADODB.Recordset
rs.Open "Select * from tregion ", db, 3, 3
With rs
        .AddNew
        .Fields("region") = Text1.Text
        .Update
End With
Form8.Timer1.Enabled = True
MsgBox "Region is save!", vbInformation
Unload Me
Else
MsgBox "Input a text!", vbExclamation
End If
Exit Sub
err:
MsgBox "Invalid Data!"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

