VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      Picture         =   "Form8.frx":5C12
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8281
      _Version        =   393216
      TabOrientation  =   2
      Style           =   1
      TabHeight       =   706
      TabCaption(0)   =   "Branch"
      TabPicture(0)   =   "Form8.frx":D634
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ImageList1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Timer1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Month and Year List"
      TabPicture(1)   =   "Form8.frx":D650
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ImageList2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Region"
      TabPicture(2)   =   "Form8.frx":D66C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command19"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command18"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "ImageList3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.TextBox Text3 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   -70965
         TabIndex        =   16
         Top             =   4245
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   -70725
         Top             =   4365
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form8.frx":D688
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   -70125
         TabIndex        =   15
         Top             =   4365
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   -71205
         Top             =   4365
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form8.frx":12E7A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3555
         TabIndex        =   13
         Top             =   4245
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4995
         Top             =   4485
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4035
         Top             =   4365
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form8.frx":19114
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   2520
         Picture         =   "Form8.frx":1E906
         TabIndex        =   9
         Top             =   4005
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Modify"
         Height          =   375
         Left            =   1560
         Picture         =   "Form8.frx":26328
         TabIndex        =   8
         Top             =   4005
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         Caption         =   "New"
         Height          =   375
         Left            =   555
         Picture         =   "Form8.frx":2DD4A
         TabIndex        =   7
         Top             =   4005
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -72480
         Picture         =   "Form8.frx":3576C
         TabIndex        =   6
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Modify"
         Height          =   375
         Left            =   -73440
         Picture         =   "Form8.frx":3D18E
         TabIndex        =   5
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton Command16 
         Caption         =   "New"
         Height          =   375
         Left            =   -74445
         Picture         =   "Form8.frx":3D607
         TabIndex        =   4
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton Command17 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -72480
         Picture         =   "Form8.frx":3DA80
         TabIndex        =   3
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Modify"
         Height          =   375
         Left            =   -73440
         Picture         =   "Form8.frx":3DEF9
         TabIndex        =   2
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton Command19 
         Caption         =   "New"
         Height          =   375
         Left            =   -74445
         Picture         =   "Form8.frx":3E372
         TabIndex        =   1
         Top             =   4080
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   3855
         Left            =   -74400
         TabIndex        =   10
         Top             =   120
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ColHdrIcons     =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Month ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Month and Year"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3855
         Left            =   600
         TabIndex        =   11
         Top             =   120
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Branch ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Branch"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   3855
         Left            =   -74400
         TabIndex        =   12
         Top             =   120
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList3"
         SmallIcons      =   "ImageList3"
         ColHdrIcons     =   "ImageList3"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Region ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Region"
            Object.Width           =   5292
         EndProperty
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command11_Click()

On Error GoTo error

If Text1.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text1.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from tbranch where branchid=" & Text1.Text & "", db, 3, 3

MsgBox "Data is remove.", vbInformation

Set rs = Nothing

Timer1.Enabled = True

Text1.Text = ""

End If

Else

MsgBox "No Information Selected!", vbExclamation

End If

Exit Sub

error:

        MsgBox "No Active Record!", vbExclamation


End Sub

Private Sub Command12_Click()
If Text1.Text <> "" Then
Form7.Show vbModal
Else
MsgBox "Select a information", vbExclamation
End If
End Sub

Private Sub Command13_Click()
Form6.Show vbModal
End Sub

Private Sub Command14_Click()

On Error GoTo error

If Text2.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text2.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from tmonthof where monthofid=" & Text2.Text & "", db, 3, 3

MsgBox "Data is remove.", vbInformation

Set rs = Nothing

Timer1.Enabled = True

Text1.Text = ""

End If

Else

MsgBox "No Information Selected!", vbExclamation

End If

Exit Sub

error:

        MsgBox "No Active Record!", vbExclamation



End Sub

Private Sub Command15_Click()
If Text2.Text <> "" Then
Form10.Show vbModal
Else
MsgBox "Select a Information!", vbExclamation
End If
End Sub

Private Sub Command16_Click()
Form9.Show vbModal
End Sub

Private Sub Command17_Click()

On Error GoTo error

If Text3.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text3.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from tregion where regionid=" & Text3.Text & "", db, 3, 3

MsgBox "Data is remove.", vbInformation

Set rs = Nothing

Timer1.Enabled = True

Text3.Text = ""

End If

Else

MsgBox "No Information Selected!", vbExclamation

End If

Exit Sub

error:

        MsgBox "No Active Record!", vbExclamation



End Sub

Private Sub Command18_Click()
If Text3.Text <> "" Then
Form12.Show vbModal
Else
MsgBox "Select a information!", vbExclamation
End If
End Sub

Private Sub Command19_Click()
Form11.Show vbModal
End Sub

Private Sub Form_Load()
branchlist
monthoflist
regionlist
End Sub

Public Sub branchlist()
On Error Resume Next
ListView3.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from tbranch order by branch"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView3.ListItems.Add , , !branchid, 1, 1
            
            ListView3.ListItems(ListView3.ListItems.Count).SubItems(1) = !branch
            
            .MoveNext
            
            Loop
            
                
                
        .Close
        
    End With
    

  Set rs = Nothing
End Sub

Private Sub ListView3_Click()

On Error GoTo err

Text1.Text = ListView3.SelectedItem.Text

Exit Sub

err:

    MsgBox "No Record!", vbExclamation
    
End Sub

Private Sub ListView4_Click()

On Error GoTo err

Text2.Text = ListView4.SelectedItem.Text

Exit Sub

err:

    MsgBox "No Record!", vbExclamation
    
End Sub

Private Sub ListView5_Click()
On Error GoTo err

Text3.Text = ListView5.SelectedItem.Text

Exit Sub

err:
        MsgBox "No Record!", vbExclamation

End Sub

Private Sub Timer1_Timer()

branchlist

monthoflist

regionlist

Timer1.Enabled = False

End Sub

Public Sub monthoflist()
On Error Resume Next
ListView4.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from tmonthof order by monthof"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView4.ListItems.Add , , !monthofid, 1, 1
            
            ListView4.ListItems(ListView4.ListItems.Count).SubItems(1) = !monthof
            
            .MoveNext
            
            Loop
            
                
                
        .Close
        
    End With
    
  Set rs = Nothing
  
End Sub

Public Sub regionlist()
On Error Resume Next
ListView5.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from tregion order by region"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView5.ListItems.Add , , !regionid, 1, 1
            
            ListView5.ListItems(ListView5.ListItems.Count).SubItems(1) = !region
            
            .MoveNext
            
            Loop
            
                
                
        .Close
        
    End With
    
  Set rs = Nothing
  
End Sub
