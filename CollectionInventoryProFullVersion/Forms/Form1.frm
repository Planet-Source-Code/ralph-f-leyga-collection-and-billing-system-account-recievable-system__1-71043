VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collection Invetory Pro Version 1.0"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Close"
      Height          =   375
      Left            =   10680
      Picture         =   "Form1.frx":5C12
      TabIndex        =   41
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7680
      Top             =   120
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Settings"
      Height          =   375
      Left            =   9480
      Picture         =   "Form1.frx":608B
      TabIndex        =   40
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton btnAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   8400
      Picture         =   "Form1.frx":6504
      TabIndex        =   19
      Top             =   7200
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   11520
      Top             =   6600
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
            Picture         =   "Form1.frx":697D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5280
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11400
      Top             =   6600
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
            Picture         =   "Form1.frx":8FF7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "Employee Information"
      TabPicture(0)   =   "Form1.frx":EC19
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Collections"
      TabPicture(1)   =   "Form1.frx":EC35
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ListView2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command6"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command7"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Combo2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text11"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Command10"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Command8"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.CommandButton Command8 
         Caption         =   "&Print"
         Height          =   375
         Left            =   -71640
         Picture         =   "Form1.frx":EC51
         TabIndex        =   42
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   375
         Left            =   -64320
         TabIndex        =   39
         Top             =   5280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -63840
         TabIndex        =   38
         Top             =   5175
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -65160
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   6015
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67440
         TabIndex        =   36
         Top             =   6000
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":F0CA
         Left            =   -69720
         List            =   "Form1.frx":F0E3
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   6015
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -72720
         Picture         =   "Form1.frx":F145
         TabIndex        =   18
         Top             =   6000
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Modify"
         Height          =   375
         Left            =   -73680
         Picture         =   "Form1.frx":F5BE
         TabIndex        =   17
         Top             =   6015
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "New"
         Height          =   375
         Left            =   -74640
         Picture         =   "Form1.frx":FA37
         TabIndex        =   16
         Top             =   6015
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Information Panel"
         Height          =   5295
         Left            =   -66480
         TabIndex        =   14
         Top             =   495
         Width           =   2535
         Begin VB.CommandButton Command14 
            Caption         =   "View"
            Height          =   255
            Left            =   1080
            TabIndex        =   47
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command13 
            Caption         =   "::"
            Height          =   375
            Left            =   1920
            TabIndex        =   46
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox tmonthof 
            Height          =   285
            Left            =   240
            TabIndex        =   45
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton Command12 
            Caption         =   "::"
            Height          =   375
            Left            =   1920
            TabIndex        =   44
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox tregion 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   43
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   4920
            Width           =   1935
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   4200
            Width           =   1935
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   3480
            Width           =   1935
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total For Release:"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   4680
            Width           =   1320
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Total Released Incentive:"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   3960
            Width           =   1845
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Total Net Incentive:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   3240
            Width           =   1440
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total Less Cash Bond:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   2520
            Width           =   1590
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total Gross Incentive:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   1800
            Width           =   1590
         End
         Begin VB.Label Label4 
            Caption         =   "Month of:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Region:"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5295
         Left            =   -74640
         TabIndex        =   13
         Top             =   615
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   9340
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Record Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Area"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Branch"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Position"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Gross Incentive"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Less: Cash Bond"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Net Incentive"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Released Incentive"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "For Release"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ATM Offline"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   3120
         Picture         =   "Form1.frx":FEB0
         TabIndex        =   5
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   1920
         Picture         =   "Form1.frx":10329
         TabIndex        =   4
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&New Employee"
         Height          =   375
         Left            =   360
         Picture         =   "Form1.frx":107A2
         TabIndex        =   3
         Top             =   6000
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4455
         Left            =   360
         TabIndex        =   2
         Top             =   615
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7858
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Employee ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Last name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "First name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Middle name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Age"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Gender"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Address"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Position"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Area"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Region"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Branch"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   360
         TabIndex        =   1
         Top             =   5175
         Width           =   10695
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   8760
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Search"
            Height          =   375
            Left            =   10200
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3600
            TabIndex        =   9
            Top             =   240
            Width           =   2655
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form1.frx":10C1B
            Left            =   960
            List            =   "Form1.frx":10C31
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Total Employee:"
            Height          =   255
            Left            =   6360
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Enter text:"
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Sea 
            Caption         =   "Search for:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label12 
         Caption         =   "Total Info:"
         Height          =   255
         Left            =   -66000
         TabIndex        =   34
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Enter Text:"
         Height          =   255
         Left            =   -68280
         TabIndex        =   33
         Top             =   6015
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Search for:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   32
         Top             =   6015
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnAbout_Click()
Form13.Show vbModal
End Sub

Private Sub Combo1_Click()

If Combo1.Text <> "" Then

Text1.Enabled = True

End If

End Sub

Private Sub Combo2_Click()

If Combo2.Text <> "" Then
Text4.Enabled = True
End If

End Sub

Private Sub Command1_Click()

Form2.Show vbModal

End Sub

Private Sub Command10_Click()

On Error Resume Next

ListView2.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from qcash where " & Combo2.Text & " like '" & Text4.Text & "%' and region='" & tregion.Text & "' and monthof='" & tmonthof.Text & "' order by name asc"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView2.ListItems.Add , , !recordnumber, 1, 1
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(1) = !Area
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(2) = !branch
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(3) = !Position
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(4) = !Status
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(5) = !Name
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(6) = !gross_incentive
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(7) = !less_cash_bond
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(8) = !net_incentive
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(9) = !release_incentive
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(10) = !for_release
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(11) = !atm_offline
            
            .MoveNext
            
            Loop
            
                Text2.Text = rs.RecordCount
                
        .Close
        
    End With
    

  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  Text5.Text = "0"
  
  Text6.Text = "0"
  
  Text7.Text = "0"
  
  Text8.Text = "0"
  
  Text9.Text = "0"
  
  Text10.Text = "0"
  
  rs.Open "Select count(recordnumber) as totalrecord from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text5.Text = rs!totalrecord
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(Gross_Incentive) as totalgross from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text6.Text = rs!totalgross
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(less_cash_bond) as totalband from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text7.Text = rs!totalband
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(net_incentive) as totalnet from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text8.Text = rs!totalnet
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(release_incentive) as totalreleased from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text9.Text = rs!totalreleased
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(for_release) as totalfor from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text10.Text = rs!totalfor
  
  Set rs = Nothing
  
End Sub

Private Sub Command11_Click()
Form8.Show vbModal
End Sub

Private Sub Command12_Click()
Form15.Show vbModal
End Sub

Private Sub Command13_Click()
Form16.Show vbModal
End Sub

Private Sub Command14_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()

If Text3.Text <> "" Then

Form3.Show vbModal

Else

MsgBox "Select a information", vbExclamation

End If

End Sub

Private Sub Command3_Click()

On Error GoTo error

If Text3.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text3.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from temployee where id=" & Text3.Text & "", db, 3, 3

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

Private Sub Command4_Click()

dbase
On Error Resume Next
ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from temployee where " & Combo1.Text & " like '" & Text1.Text & "%' order by lastname"
        
        .Open criteria, db, 3, 3
        
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !id, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !lastname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !firstname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !middlename
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !age
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = !gender
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = !address
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = !Position
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = !Status
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = !Area
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = !region
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = !branch
            
            .MoveNext
            
            Loop
            
        .Close
        
    End With
    
  Set rs = Nothing
  
End Sub

Private Sub Command5_Click()

Form4.Show vbModal

End Sub

Private Sub Command6_Click()

If Text11.Text <> "" Then


Form5.Show vbModal

Else

MsgBox "Select Collection to Edit!", vbExclamation

End If

End Sub

Private Sub Command7_Click()

On Error GoTo error

If Text11.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text11.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from tcash where recordnumber=" & Text11.Text & "", db, 3, 3

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

Private Sub Command8_Click()
Set rs = New ADODB.Recordset
rs.Open "Select * from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof.Text & "' order by name asc", db, 3, 3
Set DataReport1.DataSource = rs
DataReport1.Sections("Section2").Controls("region").Caption = tregion.Text
DataReport1.Sections("Section2").Controls("monthof").Caption = tmonthof.Text
DataReport1.Sections("Section5").Controls("totalgi").Caption = Text6.Text
DataReport1.Sections("Section5").Controls("totallcb").Caption = Text7.Text
DataReport1.Sections("Section5").Controls("totalni").Caption = Text8.Text
DataReport1.Sections("Section5").Controls("totalri").Caption = Text9.Text
DataReport1.Sections("Section5").Controls("totalfr").Caption = Text10.Text
DataReport1.Show vbModal

End Sub

Private Sub Command9_Click()

Unload Me

End Sub

Private Sub Form_Load()

dbase

employeelist

cbo

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub ListView1_Click()
On Error GoTo error
Text3.Text = ListView1.SelectedItem.Text
Exit Sub
error:
MsgBox "No Record!", vbExclamation
End Sub

Private Sub ListView2_Click()
On Error GoTo err
Text11.Text = ListView2.SelectedItem.Text
Exit Sub
err:
MsgBox "No Active Record!", vbExclamation
End Sub

Private Sub Text1_Change()

Command4_Click

End Sub

Private Sub Text12_Change()

collectlist

End Sub

Private Sub Text4_Change()

Command10_Click

End Sub

Private Sub Timer1_Timer()

'Form_Load
employeelist
collectlist
cbo
Timer1.Enabled = False

  Set rs = Nothing
    Set rs = New ADODB.Recordset
  rs.Open "Select * from temployee", db, 3, 3
  Text2.Text = rs.RecordCount
  Set rs = Nothing
End Sub

Public Sub employeelist()
On Error Resume Next
ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from temployee order by lastname"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !id, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !lastname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !firstname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !middlename
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !age
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = !gender
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = !address
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = !Position
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = !Status
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = !Area
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = !region
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = !branch
            
            .MoveNext
            
            Loop
            
                
                
        .Close
        
    End With
    
End Sub
Public Sub cbo()

'Set rs = New ADODB.Recordset

'tregion.Clear

'rs.Open "Select * from tregion order by region", db, 3, 3

  '  Do Until rs.EOF
    
    '    tregion.AddItem rs!region
        
    '    rs.MoveNext
        
  '  Loop
    
   ' Set rs = Nothing

'Set rs = New ADODB.Recordset

'tmonthof.Clear

'rs.Open "Select * from tmonthof order by monthof", db, 3, 3

 '   Do Until rs.EOF
    
   '     tmonthof.AddItem rs!monthof
        
    '    rs.MoveNext
        
    'Loop
    
   ' Set rs = Nothing
    
   'Set rs = Nothing
   
    Set rs = New ADODB.Recordset
    
  rs.Open "Select * from temployee", db, 3, 3
  
  Text2.Text = rs.RecordCount
  
  Set rs = Nothing
  
End Sub

Public Sub collectlist()

On Error Resume Next

ListView2.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof.Text & "' order by name asc"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView2.ListItems.Add , , !recordnumber, 1, 1
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(1) = !Area
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(2) = !branch
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(3) = !Position
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(4) = !Status
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(5) = !Name
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(6) = !gross_incentive
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(7) = !less_cash_bond
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(8) = !net_incentive
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(9) = !release_incentive
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(10) = !for_release
            
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(11) = !atm_offline
            
            .MoveNext
            
            Loop
            
                Text2.Text = rs.RecordCount
                
        .Close
        
    End With
    

  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  Text5.Text = "0"
  
  Text6.Text = "0"
  
  Text7.Text = "0"
  
  Text8.Text = "0"
  
  Text9.Text = "0"
  
  Text10.Text = "0"
  
  rs.Open "Select count(recordnumber) as totalrecord from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text5.Text = rs!totalrecord
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(Gross_Incentive) as totalgross from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text6.Text = rs!totalgross
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(less_cash_bond) as totalband from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text7.Text = rs!totalband
  
  Set rs = Nothing

  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(net_incentive) as totalnet from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text8.Text = rs!totalnet
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(release_incentive) as totalreleased from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text9.Text = rs!totalreleased
  
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  
  rs.Open "Select sum(for_release) as totalfor from qcash where region='" & tregion.Text & "' and monthof='" & tmonthof & "'", db, 3, 3
  
  Text10.Text = rs!totalfor
  
  Set rs = Nothing
  
End Sub

Private Sub Timer2_Timer()
  Form14.Show vbModal
  Timer2.Enabled = False
End Sub



Private Sub tmonthof_Change()

collectlist

End Sub

Private Sub tregion_Change()

collectlist

End Sub
