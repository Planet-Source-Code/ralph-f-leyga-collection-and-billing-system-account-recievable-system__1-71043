VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modify Branch"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      Picture         =   "Form7.frx":5C12
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      Picture         =   "Form7.frx":608B
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
If Text1.Text <> "" Then

Set rs = New ADODB.Recordset

rs.Open "Select * from tbranch where branchid=" & Form8.Text1.Text & "", db, 3, 3

With rs
            
            .Fields("branch") = Text1.Text
            
            .Update
End With

Form8.Timer1.Enabled = True

MsgBox "Branch is added!", vbInformation

Form8.Text1.Text = ""

Unload Me

Else

MsgBox "Input the correct branch!", vbExclamation

End If

Exit Sub
err:
MsgBox "Invalid Data!"
End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

Set rs = New ADODB.Recordset

rs.Open "Select * from tbranch where branchid=" & Form8.Text1.Text & "", db, 3, 3

Text1.Text = rs!branch

Set rs = Nothing

End Sub
