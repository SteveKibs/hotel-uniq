VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmhouses 
   BackColor       =   &H0080FF80&
   Caption         =   $"Frmhouses.frx":0000
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   13665
   WindowState     =   2  'Maximized
   Begin VB.TextBox jhn 
      Height          =   285
      Left            =   6840
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtme 
      Height          =   285
      Left            =   5400
      TabIndex        =   28
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtHouseNo 
      Height          =   285
      Left            =   3960
      TabIndex        =   27
      Top             =   1680
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid houselists 
      Height          =   4575
      Left            =   0
      TabIndex        =   25
      Top             =   4920
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   8070
      _Version        =   393216
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "HouseName"
         Caption         =   "HouseName"
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
         DataField       =   "NoofRooms"
         Caption         =   "NoofRooms"
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
         DataField       =   "TypeofRooms"
         Caption         =   "TypeofRooms"
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
         DataField       =   "Occupants"
         Caption         =   "Occupants"
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
         DataField       =   "BookingAmount"
         Caption         =   "BookingAmount"
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
      BeginProperty Column05 
         DataField       =   "Entertainment"
         Caption         =   "Entertainment"
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
      BeginProperty Column06 
         DataField       =   "OtherServices"
         Caption         =   "OtherServices"
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
            ColumnWidth     =   1739.906
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
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   8280
      Top             =   720
   End
   Begin VB.ComboBox CBOSERVICES 
      Height          =   315
      Left            =   7560
      TabIndex        =   23
      Text            =   "Select other Services Offered"
      Top             =   3480
      Width           =   2535
   End
   Begin VB.ComboBox CBOENTERTAINMENT 
      Height          =   315
      Left            =   5880
      TabIndex        =   22
      Text            =   "Select Type of Entertainment available"
      Top             =   3120
      Width           =   3135
   End
   Begin VB.ComboBox CBOOCCUPANTS 
      Height          =   315
      Left            =   5880
      TabIndex        =   21
      Text            =   "Select No of Occupants"
      Top             =   2400
      Width           =   3135
   End
   Begin VB.ComboBox cbohousetype 
      Height          =   315
      Left            =   5880
      TabIndex        =   20
      Text            =   "Choose House Type"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   12240
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   975
      Left            =   1440
      TabIndex        =   14
      Top             =   3840
      Width           =   9975
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Add New House"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Delete House"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Cancel Operation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.TextBox txtOtherServices 
      DataField       =   "OtherServices"
      Height          =   285
      Left            =   3975
      TabIndex        =   13
      Top             =   3525
      Width           =   3360
   End
   Begin VB.TextBox txtEntertainment 
      DataField       =   "Entertainment"
      Height          =   285
      Left            =   3975
      TabIndex        =   11
      Top             =   3150
      Width           =   1800
   End
   Begin VB.TextBox txtBookingAmount 
      DataField       =   "BookingAmount"
      Height          =   285
      Left            =   3975
      TabIndex        =   9
      Top             =   2760
      Width           =   1800
   End
   Begin VB.TextBox txtOccupants 
      DataField       =   "Occupants"
      Height          =   285
      Left            =   3975
      TabIndex        =   7
      Top             =   2385
      Width           =   1800
   End
   Begin VB.TextBox txtTypeofRooms 
      DataField       =   "TypeofRooms"
      Height          =   285
      Left            =   3975
      TabIndex        =   5
      Top             =   2010
      Width           =   1800
   End
   Begin VB.TextBox txtNoofRooms 
      DataField       =   "NoofRooms"
      Height          =   285
      Left            =   3975
      TabIndex        =   3
      Top             =   1260
      Width           =   3120
   End
   Begin VB.TextBox txtHouseName 
      DataField       =   "HouseName"
      Height          =   285
      Left            =   3975
      TabIndex        =   1
      Top             =   885
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "HouseNo"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JUBELEE BEACH HOTEL  FURNISHED HOUSES"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   600
      TabIndex        =   24
      Top             =   120
      Width           =   10815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OtherServices:"
      Height          =   255
      Index           =   6
      Left            =   2130
      TabIndex        =   12
      Top             =   3570
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entertainment:"
      Height          =   255
      Index           =   5
      Left            =   2130
      TabIndex        =   10
      Top             =   3195
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BookingAmount:"
      Height          =   255
      Index           =   4
      Left            =   2130
      TabIndex        =   8
      Top             =   2805
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupants:"
      Height          =   255
      Index           =   3
      Left            =   2130
      TabIndex        =   6
      Top             =   2430
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TypeofRooms:"
      Height          =   255
      Index           =   2
      Left            =   2130
      TabIndex        =   4
      Top             =   2055
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NoofRooms:"
      Height          =   255
      Index           =   1
      Left            =   2130
      TabIndex        =   2
      Top             =   1305
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HouseName:"
      Height          =   255
      Index           =   0
      Left            =   2130
      TabIndex        =   0
      Top             =   930
      Width           =   1815
   End
End
Attribute VB_Name = "Frmhouses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBOENTERTAINMENT_Click()
txtEntertainment.Text = CBOENTERTAINMENT.Text
End Sub
Private Sub cbohousetype_Click()
txtTypeofRooms.Text = cbohousetype.Text
End Sub
Private Sub CBOOCCUPANTS_Click()
txtOccupants.Text = CBOOCCUPANTS.Text
End Sub
Private Sub CBOSERVICES_Click()
txtOtherServices.Text = CBOSERVICES.Text
End Sub

Private Sub cmdAdd_Click()
Me.cmdadd.Enabled = False
Me.cmdUpdate.Enabled = True
Me.cmdDelete.Enabled = False
Me.txtHouseName.SetFocus
With openrecordset("select * from Houses")
If .BOF = True Then
txtme.Text = 0
Me.txtHouseNo.Text = 0
Else
.MoveLast
txtme.Text = !Houseno
End If
End With
End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Houses Operations,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub
Private Sub cmdUpdate_Click()
Me.cmdadd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = True

'With openrecordset("select * from Houses")
If Me.txtHouseNo.Text = "" Then
MsgBox "Enter HouseName", vbCritical
txtHouseNo.SetFocus
End If

If Me.txtHouseName.Text = "" Then
MsgBox "Enter HouseName", vbCritical
txtHouseName.SetFocus
End If

If Me.txtNoofRooms.Text = "" Then
MsgBox "Enter NoofRooms", vbCritical
txtNoofRooms.SetFocus
End If

If Me.txtTypeofRooms.Text = "" Then
MsgBox "Enter TypeofRooms", vbCritical
txtTypeofRooms.SetFocus
End If

If Me.txtOccupants.Text = "" Then
MsgBox "Enter Occupants", vbCritical
txtOccupants.SetFocus
End If

If Me.txtBookingAmount.Text = "" Then
MsgBox "Enter BookingAmount", vbCritical
txtBookingAmount.SetFocus
End If

If Me.txtEntertainment.Text = "" Then
MsgBox "Enter Entertainment", vbCritical
txtEntertainment.SetFocus
End If

If Me.txtOtherServices.Text = "" Then
MsgBox "Enter OtherServices", vbCritical
txtOtherServices.SetFocus
Else

Call gen
Call clear
MsgBox "House added succesfull"
Me.cmdadd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = True
End If
'End With
End Sub

Private Sub clear()
Me.txtHouseName.Text = ""
Me.txtNoofRooms.Text = ""
Me.txtTypeofRooms.Text = ""
Me.txtOccupants.Text = ""
Me.txtBookingAmount.Text = ""
Me.txtEntertainment.Text = ""
Me.txtOtherServices.Text = ""
End Sub

Private Sub gen()
Dim f, n As Long
With openrecordset("select * from Houses")
n = Me.txtNoofRooms.Text + Me.txtme
.MoveLast
While !Houseno < n
If !Houseno = n Then
MsgBox "Rooms Generated Succesfully", vbCritical
Else
.MoveLast
f = !Houseno + 1
.AddNew
!HouseName = Me.txtHouseName.Text
!Houseno = f
!NoofRooms = Me.txtNoofRooms.Text
!TypeofRooms = Me.txtTypeofRooms.Text
!Occupants = Me.txtOccupants.Text
!BookingAmount = Me.txtBookingAmount.Text
!Entertainment = Me.txtEntertainment.Text
!OtherServices = Me.txtOtherServices.Text
.Update
.MoveLast
End If
Wend
End With
End Sub

Private Sub Form_Load()
Me.cmdadd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = True
Set houselists.DataSource = openrecordset("select * from Houses")
Me.CBOENTERTAINMENT.AddItem "TV set"
Me.CBOENTERTAINMENT.AddItem "Sound System"
Me.CBOENTERTAINMENT.AddItem "Computer Games"
Me.CBOENTERTAINMENT.AddItem "Internet Services"
Me.cbohousetype.AddItem "Single Room"
Me.cbohousetype.AddItem "Double Room"
Me.cbohousetype.AddItem "Family House"
Me.CBOSERVICES.AddItem "Cold Drinks"
Me.CBOSERVICES.AddItem "Massage"
Me.CBOSERVICES.AddItem "Meat Roasting"
Me.CBOSERVICES.AddItem "Swiming Pool"
Me.CBOOCCUPANTS.AddItem "1"
Me.CBOOCCUPANTS.AddItem "2"
Me.CBOOCCUPANTS.AddItem "3"
Me.CBOOCCUPANTS.AddItem "4"
Me.CBOOCCUPANTS.AddItem "5"
Me.CBOOCCUPANTS.AddItem "6"
End Sub

Private Sub Timer2_Timer()
If Me.Label1.ForeColor = vbBlue Then
Label1.ForeColor = vbRed
Else
Label1.ForeColor = vbBlue
End If
End Sub

Private Sub txtNoofRooms_Change()
Dim g, k As Double
Dim w As Double
g = Val(Me.txtNoofRooms.Text)
k = Val(Me.txtme)
n = g + k
Me.jhn.Text = "" & (n)

End Sub
