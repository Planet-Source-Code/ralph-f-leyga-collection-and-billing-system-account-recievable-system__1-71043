VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Month And Year"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      Picture         =   "Form9.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      Picture         =   "Form9.frx":608B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Month and Year:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
If Text1.Text <> "" Then
Set rs = New ADODB.Recordset
rs.Open "Select * from tmonthof", db, 3, 3
With rs
            .AddNew
            .Fields("monthof") = Text1.Text
            .Update
End With
Form8.Timer1.Enabled = True
MsgBox "New year and month save!", vbInformation
Unload Me
Else
MsgBox "Input the Month and Year Correctly", vbExclamation
End If
Exit Sub
err:
MsgBox "Invalid Data!"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

