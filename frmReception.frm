VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReception 
   BackColor       =   &H00FF00FF&
   Caption         =   "                                                                                 JUBELEE BEACH HOTEL RECEPTION OFFICE"
   ClientHeight    =   11010
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid TouristsList 
      Bindings        =   "frmReception.frx":0000
      Height          =   4575
      Left            =   240
      TabIndex        =   41
      Top             =   6600
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16711935
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
      DataMember      =   "Reception"
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "FirstName"
         Caption         =   "FirstName"
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
         DataField       =   "LastName"
         Caption         =   "LastName"
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
         DataField       =   "MiddleName"
         Caption         =   "MiddleName"
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
         DataField       =   "MobileNo"
         Caption         =   "MobileNo"
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
         DataField       =   "NationalID"
         Caption         =   "NationalID"
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
         DataField       =   "MaritalStatus"
         Caption         =   "MaritalStatus"
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
         DataField       =   "Nationality"
         Caption         =   "Nationality"
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
         DataField       =   "Sex"
         Caption         =   "Sex"
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
      BeginProperty Column08 
         DataField       =   "NoofDays"
         Caption         =   "NoofDays"
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
      BeginProperty Column09 
         DataField       =   "PassportNo"
         Caption         =   "PassportNo"
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
      BeginProperty Column10 
         DataField       =   "ReceptionFees"
         Caption         =   "ReceptionFees"
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
      BeginProperty Column11 
         DataField       =   "ServedBy"
         Caption         =   "ServedBy"
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
      BeginProperty Column12 
         DataField       =   "TypeofVisit"
         Caption         =   "TypeofVisit"
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
      BeginProperty Column13 
         DataField       =   "DateofBirth"
         Caption         =   "DateofBirth"
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
      BeginProperty Column14 
         DataField       =   "CustomerNo"
         Caption         =   "CustomerNo"
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
      BeginProperty Column15 
         DataField       =   "CityState"
         Caption         =   "CityState"
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
      BeginProperty Column16 
         DataField       =   "AreaCode"
         Caption         =   "AreaCode"
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
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF00FF&
      Height          =   2295
      Left            =   240
      TabIndex        =   30
      Top             =   4080
      Width           =   6735
      Begin VB.ComboBox cbovisit 
         Height          =   315
         Left            =   4080
         TabIndex        =   45
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtNoofDays 
         DataField       =   "NoofDays"
         Height          =   285
         Left            =   1845
         TabIndex        =   35
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtPassportNo 
         DataField       =   "PassportNo"
         Height          =   285
         Left            =   1845
         TabIndex        =   34
         Top             =   615
         Width           =   1320
      End
      Begin VB.TextBox txtReceptionFees 
         DataField       =   "ReceptionFees"
         Height          =   285
         Left            =   1845
         TabIndex        =   33
         Top             =   990
         Width           =   1320
      End
      Begin VB.TextBox txtServedBy 
         DataField       =   "ServedBy"
         Height          =   285
         Left            =   1845
         TabIndex        =   32
         Top             =   1380
         Width           =   3375
      End
      Begin VB.TextBox txtTypeofVisit 
         DataField       =   "TypeofVisit"
         Height          =   285
         Left            =   1845
         TabIndex        =   31
         Top             =   1755
         Width           =   2295
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoofDays:"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   40
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PassportNo:"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   39
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReceptionFees:"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   38
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ServedBy:"
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   37
         Top             =   1425
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TypeofVisit:"
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   36
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF00FF&
      Height          =   2415
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   11895
      Begin VB.ComboBox cbosex 
         Height          =   315
         Left            =   4200
         TabIndex        =   44
         Text            =   "Choose Sex"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox cbocountry 
         Height          =   315
         Left            =   4200
         TabIndex        =   43
         Text            =   "Choose Country"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox cbostatus 
         Height          =   315
         Left            =   4080
         TabIndex        =   42
         Text            =   "Choose Marital Status"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtMobileNo 
         DataField       =   "MobileNo"
         Height          =   285
         Left            =   2445
         TabIndex        =   20
         Top             =   480
         Width           =   1320
      End
      Begin VB.TextBox txtNationalID 
         DataField       =   "NationalID"
         Height          =   285
         Left            =   2445
         TabIndex        =   19
         Top             =   855
         Width           =   1320
      End
      Begin VB.TextBox txtMaritalStatus 
         DataField       =   "MaritalStatus"
         Height          =   285
         Left            =   2445
         TabIndex        =   18
         Top             =   1245
         Width           =   1695
      End
      Begin VB.TextBox txtNationality 
         DataField       =   "Nationality"
         Height          =   285
         Left            =   2445
         TabIndex        =   17
         Top             =   1620
         Width           =   1815
      End
      Begin VB.TextBox txtSex 
         DataField       =   "Sex"
         Height          =   285
         Left            =   2445
         TabIndex        =   16
         Top             =   1995
         Width           =   1815
      End
      Begin VB.TextBox txtDateofBirth 
         DataField       =   "DateofBirth"
         Height          =   285
         Left            =   8205
         TabIndex        =   15
         Top             =   675
         Width           =   1320
      End
      Begin VB.TextBox txtCustomerNo 
         DataField       =   "CustomerNo"
         Height          =   285
         Left            =   8205
         TabIndex        =   14
         Top             =   1065
         Width           =   3375
      End
      Begin VB.TextBox txtCityState 
         DataField       =   "CityState"
         Height          =   285
         Left            =   8205
         TabIndex        =   13
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtAreaCode 
         DataField       =   "AreaCode"
         Height          =   285
         Left            =   8205
         TabIndex        =   12
         Top             =   1815
         Width           =   3375
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MobileNo:"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   29
         Top             =   525
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NationalID:"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   28
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MaritalStatus:"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   27
         Top             =   1290
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality:"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   26
         Top             =   1665
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   25
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DateofBirth:"
         Height          =   255
         Index           =   13
         Left            =   6360
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CustomerNo:"
         Height          =   255
         Index           =   14
         Left            =   6360
         TabIndex        =   23
         Top             =   1110
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CityState:"
         Height          =   255
         Index           =   15
         Left            =   6360
         TabIndex        =   22
         Top             =   1485
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AreaCode:"
         Height          =   255
         Index           =   16
         Left            =   6360
         TabIndex        =   21
         Top             =   1860
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF00FF&
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtFirstName 
         DataField       =   "FirstName"
         Height          =   285
         Left            =   1365
         TabIndex        =   7
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox txtLastName 
         DataField       =   "LastName"
         Height          =   285
         Left            =   1365
         TabIndex        =   6
         Top             =   495
         Width           =   3375
      End
      Begin VB.TextBox txtMiddleName 
         DataField       =   "MiddleName"
         Height          =   285
         Left            =   1365
         TabIndex        =   5
         Top             =   885
         Width           =   3375
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FirstName:"
         Height          =   255
         Index           =   0
         Left            =   -480
         TabIndex        =   10
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LastName:"
         Height          =   255
         Index           =   1
         Left            =   -480
         TabIndex        =   9
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MiddleName:"
         Height          =   255
         Index           =   2
         Left            =   -480
         TabIndex        =   8
         Top             =   930
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Height          =   2295
      Left            =   6960
      TabIndex        =   0
      Top             =   4080
      Width           =   5415
      Begin VB.CommandButton CMDCANCEL 
         BackColor       =   &H0000FFFF&
         Caption         =   "CANCEL OPERATION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H0000FFFF&
         Caption         =   "&CLOSE RECEPTION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H0000FFFF&
         Caption         =   "&UPDATE DETAILS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FFFF&
         Caption         =   "&NEW CUSTOMER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public constr As String
'Public newflag As Boolean
Dim WithEvents receptionRS As Recordset
Attribute receptionRS.VB_VarHelpID = -1
'Dim WithEvents stdRS As Recordset
'Dim WithEvents classesRS As Recordset
'Dim WithEvents invoiceRS As Recordset
'Dim WithEvents feesRS As Recordset
Private Sub cbocountry_Click()
txtNationality.Text = Me.cbocountry.Text
End Sub
Private Sub cbosex_Click()
txtSex.Text = cbosex.Text
End Sub
Private Sub cbostatus_Click()
txtMaritalStatus.Text = Me.cbostatus.Text
End Sub
Private Sub cbovisit_Click()
txtTypeofVisit.Text = cbovisit.Text
End Sub
Private Sub cmdAdd_Click()
Call gencusn
Call Enabletext
Call Enablecombo
Me.CMDCANCEL.Enabled = True
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
Me.txtFirstName.SetFocus
End Sub

Private Sub CMDCANCEL_Click()
Call clear
Call Disablecombo
Call Disabletext
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.CMDCANCEL.Enabled = False
End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Reception Operations ,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub
Private Sub cmdUpdate_Click()
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False

If Me.txtFirstName.Text = "" Then
MsgBox "Enter Customer FirstName", vbCritical
txtFirstName.SetFocus
Else

If Me.txtMiddleName.Text = "" Then
MsgBox "Enter Customer MiddleName", vbCritical
txtMiddleName.SetFocus
Else

If Me.txtLastName.Text = "" Then
MsgBox "Enter Customer LastName", vbCritical
txtLastName.SetFocus
Else

If Me.txtMobileNo.Text = "" Then
MsgBox "Enter Customer MobileNo", vbCritical
txtMobileNo.SetFocus
Else

If Me.txtNationalID.Text = "" Then
MsgBox "Enter Customer NationalID", vbCritical
txtNationalID.SetFocus
Else

If Me.txtMaritalStatus.Text = "" Then
MsgBox "Enter Customer MaritalStatus", vbCritical
txtMaritalStatus.SetFocus
Else

If Me.txtNationality.Text = "" Then
MsgBox "Enter Customer Nationality", vbCritical
txtNationality.SetFocus
Else

If Me.txtSex.Text = "" Then
MsgBox "Enter Customer Sex", vbCritical
txtSex.SetFocus
Else

If Me.txtNoofDays.Text = "" Then
MsgBox "Enter Customer NoofDays", vbCritical
txtNoofDays.SetFocus
Else

If Me.txtPassportNo.Text = "" Then
MsgBox "Enter Customer PassportNo", vbCritical
txtPassportNo.SetFocus
Else

If Me.txtReceptionFees.Text = "" Then
MsgBox "Enter Customer ReceptionFees", vbCritical
txtReceptionFees.SetFocus
Else

If Me.txtServedBy.Text = "" Then
MsgBox "Enter Customer ServedBy", vbCritical
txtServedBy.SetFocus
Else

If Me.txtTypeofVisit.Text = "" Then
MsgBox "Enter Customer TypeofVisit", vbCritical
txtTypeofVisit.SetFocus
Else

If Me.txtDateofBirth.Text = "" Then
MsgBox "Enter Customer DateofBirth", vbCritical
txtDateofBirth.SetFocus
Else

If Me.txtCustomerNo.Text = "" Then
MsgBox "Enter Customer No", vbCritical
txtCustomerNo.SetFocus
Else

If Me.txtCityState.Text = "" Then
MsgBox "Enter Customer CityState", vbCritical
txtCityState.SetFocus
Else

If Me.txtAreaCode.Text = "" Then
MsgBox "Enter Customer AreaCode", vbCritical
txtAreaCode.SetFocus
Else

receptionRS.AddNew
receptionRS!FirstName = Me.txtFirstName.Text
receptionRS!MiddleName = Me.txtMiddleName.Text
receptionRS!LastName = Me.txtLastName.Text
receptionRS!MobileNo = Me.txtMobileNo.Text
'receptionRS!NationalID = Me.txtNationalID.Text
receptionRS!MaritalStatus = Me.txtMaritalStatus.Text
receptionRS!Nationality = Me.txtNationality.Text
receptionRS!Sex = Me.txtSex.Text
receptionRS!NoofDays = Me.txtNoofDays.Text
receptionRS!PassportNo = Me.txtPassportNo.Text
receptionRS!ReceptionFees = Me.txtReceptionFees.Text
receptionRS!ServedBy = Me.txtServedBy.Text
receptionRS!TypeofVisit = Me.txtTypeofVisit.Text
receptionRS!DateofBirth = Me.txtDateofBirth.Text
receptionRS!CustomerNo = Me.txtCustomerNo.Text
receptionRS!CityState = Me.txtCityState.Text
receptionRS!AreaCode = Me.txtAreaCode.Text
receptionRS.Update
receptionRS.Requery
MsgBox "Tourist Regristraticon Succesful", vbCritical
frmtag.Show
Call clear
Call Disabletext
Call Disablecombo
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.CMDCANCEL.Enabled = False
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
'
End Sub
Private Sub clear()
Me.txtFirstName.Text = ""
Me.txtMiddleName.Text = ""
Me.txtLastName.Text = ""
Me.txtMobileNo.Text = ""
Me.txtNationalID.Text = ""
Me.txtMaritalStatus.Text = ""
Me.txtNationality.Text = ""
Me.txtSex.Text = ""
Me.txtNoofDays.Text = ""
Me.txtPassportNo.Text = ""
Me.txtReceptionFees.Text = ""
Me.txtServedBy.Text = ""
Me.txtTypeofVisit.Text = ""
Me.txtDateofBirth.Text = ""
Me.txtCustomerNo.Text = ""
Me.txtCityState.Text = ""
Me.txtAreaCode.Text = ""
End Sub



Private Sub Form_Load()
Dim db As Connection
Dim cntrl As Control
constr = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open constr

Set receptionRS = Nothing

db.Close
db.Open constr

Set receptionRS = New Recordset
receptionRS.Open "select * from reception", db, adOpenStatic, adLockOptimistic

Call Disabletext
Call Disablecombo
Me.CMDCANCEL.Enabled = False
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cbosex.AddItem "Male"
Me.cbosex.AddItem "Female"
Me.cbostatus.AddItem "Married"
Me.cbostatus.AddItem "Divorced"
Me.cbostatus.AddItem "Single"
Me.cbostatus.AddItem "Windowed"
Me.cbocountry.AddItem "Kenya"
Me.cbocountry.AddItem "Uganda"
Me.cbocountry.AddItem "Tanzania"
Me.cbocountry.AddItem "Burundi"
Me.cbocountry.AddItem "Rwanda"
Me.cbocountry.AddItem "Congo"
Me.cbocountry.AddItem "South Africa"
Me.cbocountry.AddItem "Zambia"
Me.cbocountry.AddItem "Lesotho"
Me.cbocountry.AddItem "Mozambique"
Me.cbocountry.AddItem "Sudan"
Me.cbocountry.AddItem "Somalia"
Me.cbocountry.AddItem "Eqypt"
Me.cbocountry.AddItem "Iran"
Me.cbocountry.AddItem "Irack"
Me.cbocountry.AddItem "India"
Me.cbocountry.AddItem "Japan"
Me.cbocountry.AddItem "Korea"
Me.cbocountry.AddItem "Califonia"
Me.cbocountry.AddItem "Los Angles"
Me.cbocountry.AddItem "Colombia"
Me.cbocountry.AddItem "Nigeria"
Me.cbocountry.AddItem "Seria"
Me.cbocountry.AddItem "Pakistan"
Me.cbocountry.AddItem "Saudi Arabia"
Me.cbocountry.AddItem "U.S.A"
Me.cbocountry.AddItem "Germany"
Me.cbocountry.AddItem "Huwai"
Me.cbocountry.AddItem "Cambodia"
Me.cbocountry.AddItem "Vietnam"
cbovisit.AddItem "Family Holiday"
cbovisit.AddItem "Company Research Tour"
cbovisit.AddItem "Individual Holiday"
cbovisit.AddItem "School Tour"
cbovisit.AddItem "Touring Africa"
'Set TouristsList.DataSource = OpenRecordset("select * from Reception")
End Sub
Private Sub gencusn()
Dim nam As Long
Dim adm As String
Dim s As String
Dim mth, yearpart As String
If receptionRS.BOF = True Then
Me.txtCustomerNo.Text = "JBHCN / 13 / 1"
Else
receptionRS.MoveFirst
receptionRS.MoveLast
nam = receptionRS.RecordCount + 1
mth = Format(Now, "mmmm - yyyy")
yearpart = Right(mth, 2)
s = "JBHCN"
adm = (s & "/" & yearpart & "/" & nam)
Me.txtCustomerNo.Text = adm
End If
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
