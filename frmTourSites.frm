VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTourSites 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TourSites"
   ClientHeight    =   9645
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmTourSites.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid TourSites 
      Height          =   5175
      Left            =   2880
      TabIndex        =   16
      Top             =   3360
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9128
      _Version        =   393216
      BackColor       =   8454016
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "ToursiteName"
         Caption         =   "ToursiteName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TourSiteCode"
         Caption         =   "TourSiteCode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "DistanceinKm"
         Caption         =   "DistanceinKm"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Chargesperhead"
         Caption         =   "Chargesperhead"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "VehicleEntranceFee"
         Caption         =   "VehicleEntranceFee"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Height          =   2175
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   5295
      Begin VB.TextBox txtToursiteName 
         DataField       =   "ToursiteName"
         Height          =   285
         Left            =   1725
         TabIndex        =   10
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtTourSiteCode 
         DataField       =   "TourSiteCode"
         Height          =   285
         Left            =   1725
         TabIndex        =   9
         Top             =   615
         Width           =   1320
      End
      Begin VB.TextBox txtDistanceinKm 
         DataField       =   "DistanceinKm"
         Height          =   285
         Left            =   1725
         TabIndex        =   8
         Top             =   1005
         Width           =   3375
      End
      Begin VB.TextBox txtChargesperhead 
         DataField       =   "Chargesperhead"
         Height          =   285
         Left            =   1725
         TabIndex        =   7
         Top             =   1380
         Width           =   1320
      End
      Begin VB.TextBox txtVehicleEntranceFee 
         DataField       =   "VehicleEntranceFee"
         Height          =   285
         Left            =   1725
         TabIndex        =   6
         Top             =   1755
         Width           =   1320
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ToursiteName:"
         Height          =   255
         Index           =   0
         Left            =   -120
         TabIndex        =   15
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TourSiteCode:"
         Height          =   255
         Index           =   1
         Left            =   -120
         TabIndex        =   14
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DistanceinKm:"
         Height          =   255
         Index           =   2
         Left            =   -120
         TabIndex        =   13
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chargesperhead:"
         Height          =   255
         Index           =   3
         Left            =   -120
         TabIndex        =   12
         Top             =   1425
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VehicleEntranceFee:"
         Height          =   255
         Index           =   4
         Left            =   -120
         TabIndex        =   11
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      Top             =   2520
      Width           =   5295
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Add Site"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Delete"
         Height          =   300
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Update"
         Height          =   300
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Close"
         Height          =   300
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTourSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Call Enabletext
Call gencode
Me.txtToursiteName.SetFocus
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Tour Sites Operations ,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub

Private Sub cmdUpdate_Click()
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
With openrecordset("select * from TourSites")

If Me.txtToursiteName.Text = "" Then
MsgBox "Enter ToursiteName", vbCritical
txtToursiteName.SetFocus
Else

If Me.txtTourSiteCode.Text = "" Then
MsgBox "Enter TourSiteCode", vbCritical
txtTourSiteCode.SetFocus
Else

If Me.txtDistanceinKm.Text = "" Then
MsgBox "Enter DistanceinKm", vbCritical
txtDistanceinKm.SetFocus
Else

If Me.txtChargesperhead.Text = "" Then
MsgBox "Enter Chargesperhead", vbCritical
txtChargesperhead.SetFocus
Else

If Me.txtVehicleEntranceFee.Text = "" Then
MsgBox "Enter VehicleEntranceFee", vbCritical
txtVehicleEntranceFee.SetFocus
Else

.AddNew
!ToursiteName = Me.txtToursiteName.Text
!TourSiteCode = Me.txtTourSiteCode
!DistanceinKm = Me.txtDistanceinKm.Text
!Chargesperhead = Me.txtChargesperhead.Text
!VehicleEntranceFee = Me.txtVehicleEntranceFee.Text
.Update
.Requery
MsgBox "Tour Site Added Successfully", vbCritical
dbttoursites.Show
Call clear
Call Disabletext
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
End If
End If
End If
End If
End If
End With
End Sub
Private Sub clear()
Me.txtToursiteName.Text = ""
Me.txtTourSiteCode.Text = ""
Me.txtDistanceinKm.Text = ""
Me.txtChargesperhead.Text = ""
Me.txtVehicleEntranceFee.Text = ""
End Sub

Private Sub Disabletext()
    On Error Resume Next
    Dim cntrl As Control
    For Each cntrl In Me
    If TypeOf cntrl Is TextBox Then
    cntrl.Enabled = False
    End If
    Next cntrl
End Sub
Private Sub Enabletext()
    On Error Resume Next
    Dim cntrl As Control
    For Each cntrl In Me
    If TypeOf cntrl Is TextBox Then
    cntrl.Enabled = True
    End If
    Next cntrl
End Sub
Private Sub Enablecombo()
    On Error Resume Next
    Dim cntrl As Control
    For Each cntrl In Me
    If TypeOf cntrl Is ComboBox Then
    cntrl.Enabled = True
    End If
    Next cntrl

End Sub
Private Sub Disablecombo()
    On Error Resume Next
    Dim cntrl As Control
    For Each cntrl In Me
    If TypeOf cntrl Is ComboBox Then
    cntrl.Enabled = False
    End If
    Next cntrl
End Sub

Private Sub Form_Load()
Set TourSites.DataSource = openrecordset("select * from TourSites")
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Call Disabletext
End Sub
Private Sub gencode()
Dim nam As Long
Dim adm As String
Dim s As String
Dim mth, yearpart As String
nam = openrecordset("select * from TourSites").RecordCount + 1
mth = Format(Now, "mmmm - yyyy")
yearpart = Right(mth, 2)
s = "TSC"
adm = (s & "/" & yearpart & "/" & nam)
Me.txtTourSiteCode.Text = adm
End Sub



