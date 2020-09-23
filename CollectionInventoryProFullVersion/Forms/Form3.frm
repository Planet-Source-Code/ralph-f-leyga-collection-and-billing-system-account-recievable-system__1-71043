VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modify Employee"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tlastname 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox tfirstname 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox tmiddlename 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox tage 
      Height          =   285
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox tgender 
      Height          =   315
      ItemData        =   "Form3.frx":5C12
      Left            =   1200
      List            =   "Form3.frx":5C1C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox taddress 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   2280
      Width           =   3855
   End
   Begin VB.ComboBox tposition 
      Height          =   315
      ItemData        =   "Form3.frx":5C26
      Left            =   1200
      List            =   "Form3.frx":5C2D
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox tstatus 
      Height          =   315
      ItemData        =   "Form3.frx":5C37
      Left            =   1200
      List            =   "Form3.frx":5C3E
      TabIndex        =   7
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox tarea 
      Height          =   315
      ItemData        =   "Form3.frx":5C48
      Left            =   1200
      List            =   "Form3.frx":5C4F
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.ComboBox tregion 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ComboBox tbranch 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5160
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label1 
      Caption         =   "Last name:"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "First name:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Middle name:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Age:"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Gender:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Position:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Area:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Region:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Branch:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   5160
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5160
      Y1              =   4560
      Y2              =   4560
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If tlastname.Text <> "" And tfirstname.Text <> "" And tmiddlename.Text <> "" And tage.Text <> "" And tgender.Text <> "" And taddress.Text <> "" And tposition.Text <> "" And tarea.Text <> "" And tregion.Text <> "" And tbranch.Text <> "" Then

Set rs = New ADODB.Recordset
rs.Open "Select * from temployee where id=" & Form1.Text3.Text & "", db, 3, 3
With rs

        .Fields("lastname") = tlastname.Text
        .Fields("firstname") = tfirstname.Text
        .Fields("middlename") = tmiddlename.Text
        .Fields("age") = tage.Text
        .Fields("gender") = tgender.Text
        .Fields("address") = taddress.Text
        .Fields("position") = tposition.Text
        .Fields("status") = tstatus.Text
        .Fields("area") = tarea.Text
        .Fields("region") = tregion.Text
        .Fields("branch") = tbranch.Text
        .Update
End With
MsgBox "Employee is save!", vbInformation
Form1.Timer1.Enabled = True
Form1.Text3.Text = ""
Set rs = Nothing
Unload Me
Else
MsgBox "All fields are required!", vbExclamation
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
cbo
Set rs = New ADODB.Recordset
rs.Open "Select * from temployee where id=" & Form1.Text3.Text & "", db, 3, 3
        tlastname.Text = rs.Fields("lastname")
        tfirstname.Text = rs.Fields("firstname")
        tmiddlename.Text = rs.Fields("middlename")
        tage.Text = rs.Fields("age")
        tgender.Text = rs.Fields("gender")
        taddress.Text = rs.Fields("address")
        tposition.Text = rs.Fields("position")
        tstatus.Text = rs.Fields("status")
        tarea.Text = rs.Fields("area")
        tregion.Text = rs.Fields("region")
        tbranch.Text = rs.Fields("branch")
Set rs = Nothing
End Sub
Public Sub cbo()
Set rs = New ADODB.Recordset
rs.Open "Select * from tregion order by region", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rs.EOF
        tregion.AddItem rs!region
        rs.MoveNext
    Loop
    Set rs = Nothing
    
Set rs = New ADODB.Recordset
rs.Open "Select * from tbranch order by branch", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rs.EOF
        tbranch.AddItem rs!branch
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub
