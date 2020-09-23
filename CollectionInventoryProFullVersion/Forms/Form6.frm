VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Branch"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      Picture         =   "Form6.frx":5C12
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      Picture         =   "Form6.frx":608B
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo err

If Text1.Text <> "" Then

Set rs = New ADODB.Recordset

rs.Open "Select * from tbranch", db, 3, 3

With rs

            .AddNew
            
            .Fields("branch") = Text1.Text
            
            .Update
End With

Form8.Timer1.Enabled = True

MsgBox "New branch is added!", vbInformation

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

