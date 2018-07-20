VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmtransport 
   BackColor       =   &H0080FF80&
   Caption         =   "                                                   TRANSPORT SYSTSEM"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Frmtransport.frx":0000
   ScaleHeight     =   9570
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid Triplist 
      Height          =   4575
      Left            =   2520
      TabIndex        =   22
      Top             =   4560
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   8454016
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "CarBusPlateNo"
         Caption         =   "CarBusPlateNo"
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
         DataField       =   "Cartype"
         Caption         =   "Cartype"
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
         DataField       =   "DepatureTime"
         Caption         =   "DepatureTime"
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
         DataField       =   "Destination"
         Caption         =   "Destination"
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
         DataField       =   "DriverIDNo"
         Caption         =   "DriverIDNo"
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
         DataField       =   "DriverNames"
         Caption         =   "DriverNames"
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
         DataField       =   "Noofkms"
         Caption         =   "Noofkms"
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
      BeginProperty Column07 
         DataField       =   "NoofTourists"
         Caption         =   "NoofTourists"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Height          =   3375
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   6855
      Begin VB.ComboBox cbocar 
         Height          =   315
         Left            =   4440
         TabIndex        =   23
         Text            =   "Select CarType"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtCarBusPlateNo 
         BackColor       =   &H00FFFFFF&
         DataField       =   "CarBusPlateNo"
         Height          =   285
         Left            =   2685
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtDepatureTime 
         DataField       =   "DepatureTime"
         Height          =   285
         Left            =   2685
         TabIndex        =   12
         Top             =   1005
         Width           =   1320
      End
      Begin VB.TextBox txtDestination 
         DataField       =   "Destination"
         Height          =   285
         Left            =   2685
         TabIndex        =   11
         Top             =   1380
         Width           =   3375
      End
      Begin VB.TextBox txtDriverIDNo 
         DataField       =   "DriverIDNo"
         Height          =   285
         Left            =   2685
         TabIndex        =   10
         Top             =   1755
         Width           =   660
      End
      Begin VB.TextBox txtDriverNames 
         DataField       =   "DriverNames"
         Height          =   285
         Left            =   2685
         TabIndex        =   9
         Top             =   2145
         Width           =   3375
      End
      Begin VB.TextBox txtNoofkms 
         DataField       =   "Noofkms"
         Height          =   285
         Left            =   2685
         TabIndex        =   8
         Top             =   2520
         Width           =   660
      End
      Begin VB.TextBox txtNoofTourists 
         DataField       =   "NoofTourists"
         Height          =   285
         Left            =   2685
         TabIndex        =   7
         Top             =   2895
         Width           =   660
      End
      Begin VB.TextBox txtCartype 
         DataField       =   "Cartype"
         Height          =   285
         Left            =   2685
         TabIndex        =   6
         Top             =   615
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CarBusPlateNo:"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   21
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cartype:"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   20
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DepatureTime:"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   19
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination:"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   18
         Top             =   1425
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DriverIDNo:"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   17
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DriverNames:"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   16
         Top             =   2190
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Noofkms:"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   15
         Top             =   2565
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoofTourists:"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   14
         Top             =   2940
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   3600
      Width           =   6855
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
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
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
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Add New Trip"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Frmtransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbocar_Click()
Me.txtCartype.Text = cbocar.Text
End Sub

Private Sub cmdAdd_Click()
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
Me.txtCarBusPlateNo.SetFocus
End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Transport Operations ,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub

Private Sub cmdUpdate_Click()
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True

With openrecordset("select * from Transport")

If Me.txtCarBusPlateNo.Text = "" Then
MsgBox "Enter CarBusPlateNo", vbCritical
Me.txtCarBusPlateNo.SetFocus
End If

If Me.txtCartype.Text = "" Then
MsgBox "Enter Cartype", vbCritical
Me.txtCartype.SetFocus
End If

If Me.txtDepatureTime.Text = "" Then
MsgBox "Enter DepatureTime", vbCritical
Me.txtDepatureTime.SetFocus
End If

If Me.txtDestination.Text = "" Then
MsgBox "Enter Destination", vbCritical
Me.txtDestination.SetFocus
End If

If Me.txtDriverIDNo.Text = "" Then
MsgBox "Enter DriverIDNo", vbCritical
Me.txtDriverIDNo.SetFocus
End If

If Me.txtDriverNames.Text = "" Then
MsgBox "Enter DriverNames", vbCritical
Me.txtDriverNames.SetFocus
End If

If Me.txtNoofkms.Text = "" Then
MsgBox "Enter Noofkms", vbCritical
Me.txtNoofkms.SetFocus
End If

If Me.txtNoofTourists.Text = "" Then
MsgBox "Enter NoofTourists", vbCritical
Me.txtNoofTourists.SetFocus
Else

.AddNew
!CarBusPlateNo = Me.txtCarBusPlateNo
!Cartype = Me.txtCartype
!DepatureTime = Me.txtDepatureTime
!Destination = Me.txtDestination
'!DriverIDNo = Me.txtDriverIDNo
!DriverNames = Me.txtDriverNames
'!Noofkms = Me.txtNoofkms
!NoofTourists = Me.txtNoofTourists
.Update
.Requery
MsgBox "Printing Gate Pass Now", vbCritical
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Call clear
End If
End With

End Sub
Private Sub clear()
Me.txtCarBusPlateNo.Text = ""
Me.txtCartype.Text = ""
Me.txtDepatureTime.Text = ""
Me.txtDestination.Text = ""
Me.txtDriverIDNo.Text = ""
Me.txtDriverNames.Text = ""
Me.txtNoofkms.Text = ""
Me.txtNoofTourists.Text = ""

End Sub

Private Sub Form_Load()
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False

cbocar.AddItem "52 SEATER BUS"
cbocar.AddItem "14 SEATER RANGE ROVER"
cbocar.AddItem "10 SEATER TOYOTA"
cbocar.AddItem "6 SEATER RANGE"
cbocar.AddItem "4 SEATER TOYOTA RANGE"
Set Triplist.DataSource = openrecordset("select * from Transport")
End Sub
