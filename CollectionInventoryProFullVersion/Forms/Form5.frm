VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modify Collection"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      Picture         =   "Form5.frx":5C12
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6600
      Picture         =   "Form5.frx":608B
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox tid 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox tgrossincentive 
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox tlesscashband 
      Height          =   285
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox tnetIncentive 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox treleasedincentive 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox tforReleased 
      Height          =   285
      Left            =   6600
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox tATMoffline 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox tmonthof 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee ID:"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Gross Incentive:"
      Height          =   195
      Left            =   2640
      TabIndex        =   14
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Less Cash Bond:"
      Height          =   195
      Left            =   5160
      TabIndex        =   13
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Net Incentive:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Released Incentive:"
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "For Released:"
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ATM Offline:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Month of:"
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   1560
      Width           =   705
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If tgrossincentive.Text <> "" And tid.Text <> "" And tnetIncentive.Text <> "" And treleasedincentive.Text <> "" And tforReleased.Text <> "" And tATMoffline.Text <> "" And tmonthof.Text <> "" Then
Set rs = New ADODB.Recordset

rs.Open "Select * from tcash where recordnumber=" & Form1.Text11.Text & "", db, 3, 3


With rs
        .Fields("id") = tid.Text
        
        .Fields("gross_incentive") = tgrossincentive.Text
        
        .Fields("less_cash_bond") = tlesscashband.Text
        
        .Fields("Net_incentive") = tnetIncentive.Text
        
        .Fields("release_incentive") = treleasedincentive.Text
        
        .Fields("for_release") = tforReleased.Text
        
        .Fields("atm_offline") = tATMoffline.Text
        
        .Fields("monthof") = tmonthof.Text
        
        .Update
        
End With

Form1.Timer1.Enabled = True

Form1.Text11.Text = ""

MsgBox "New Collection is Added!", vbInformation

Form1.Text11.Text = ""

Unload Me

Set rs = Nothing

Else

MsgBox "All fields are required!", vbExclamation

End If

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

cbo

Set rs = New ADODB.Recordset

rs.Open "Select * from tcash where recordnumber=" & Form1.Text11.Text & "", db, 3, 3

tid.Text = rs!id

tgrossincentive.Text = rs!gross_incentive

tlesscashband.Text = rs!less_cash_bond

tnetIncentive.Text = rs!net_incentive

treleasedincentive.Text = rs!release_incentive

tforReleased.Text = rs!for_release

tATMoffline.Text = rs!atm_offline

tmonthof.Text = rs!monthof


End Sub

Public Sub cbo()

Set rs = New ADODB.Recordset

tid.Clear

rs.Open "Select * from temployee order by id", db, 3, 3

    Do Until rs.EOF
    
        tid.AddItem rs!id
        
        rs.MoveNext
        
    Loop
    
    Set rs = Nothing
    
Set rs = New ADODB.Recordset

tmonthof.Clear

rs.Open "Select * from tmonthof order by monthof", db, 3, 3

    Do Until rs.EOF
    
        tmonthof.AddItem rs!monthof
        
        rs.MoveNext
        
    Loop
    
    Set rs = Nothing
    

End Sub
