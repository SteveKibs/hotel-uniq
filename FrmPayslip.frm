VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPayslip 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Height          =   2535
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   10695
      Begin VB.TextBox txtWorkerCategory 
         DataField       =   "WorkerCategory"
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtDepartment 
         DataField       =   "Department"
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtBasicSalary 
         DataField       =   "BasicSalary"
         Height          =   285
         Left            =   1935
         TabIndex        =   18
         Top             =   990
         Width           =   1320
      End
      Begin VB.TextBox txtAllowances 
         DataField       =   "Allowances"
         Height          =   285
         Left            =   1965
         TabIndex        =   17
         Top             =   1365
         Width           =   1320
      End
      Begin VB.TextBox txtDeductions 
         DataField       =   "Deductions"
         Height          =   285
         Left            =   1965
         TabIndex        =   16
         Top             =   1740
         Width           =   1320
      End
      Begin VB.TextBox txtNetSalary 
         DataField       =   "NetSalary"
         Height          =   285
         Left            =   1965
         TabIndex        =   15
         Top             =   2130
         Width           =   1320
      End
      Begin VB.TextBox txtDated 
         DataField       =   "Dated"
         Height          =   285
         Left            =   6885
         TabIndex        =   14
         Top             =   225
         Width           =   1320
      End
      Begin VB.TextBox txtPayslipNo 
         DataField       =   "PayslipNo"
         Height          =   285
         Left            =   6885
         TabIndex        =   13
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtPrintedBy 
         DataField       =   "PrintedBy"
         Height          =   285
         Left            =   6885
         TabIndex        =   12
         Top             =   990
         Width           =   3375
      End
      Begin VB.TextBox txtTime 
         DataField       =   "Time"
         Height          =   285
         Left            =   6885
         TabIndex        =   11
         Top             =   1365
         Width           =   1320
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WorkerCategory:"
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BasicSalary:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allowances:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   1410
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deductions:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   1785
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NetSalary:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   2175
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dated:"
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   23
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PayslipNo:"
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   22
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PrintedBy:"
         Height          =   255
         Index           =   8
         Left            =   5040
         TabIndex        =   21
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   20
         Top             =   1410
         Width           =   1815
      End
   End
   Begin MSDataGridLib.DataGrid Paysliplist 
      Bindings        =   "FrmPayslip.frx":0000
      Height          =   5055
      Left            =   0
      TabIndex        =   9
      Top             =   4680
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   8916
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
      DataMember      =   "Payslip"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "WorkerNo"
         Caption         =   "WorkerNo"
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
         DataField       =   "Department"
         Caption         =   "Department"
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
         DataField       =   "BasicSalary"
         Caption         =   "BasicSalary"
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
         DataField       =   "Allowances"
         Caption         =   "Allowances"
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
         DataField       =   "Deductions"
         Caption         =   "Deductions"
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
         DataField       =   "NetSalary"
         Caption         =   "NetSalary"
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
         DataField       =   "Dated"
         Caption         =   "Dated"
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
         DataField       =   "PayslipNo"
         Caption         =   "PayslipNo"
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
         DataField       =   "PrintedBy"
         Caption         =   "PrintedBy"
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
         DataField       =   "Time"
         Caption         =   "Time"
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
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF00FF&
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   10695
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
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
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Show Payslip History "
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Cancel Operations"
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
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Generate Payslip"
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
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdhide 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Hide Payslip History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox txtWorkerNo 
      DataField       =   "WorkerNo"
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT EMPLOYEES PAYSLIPS"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2520
      TabIndex        =   31
      Top             =   120
      Width           =   11055
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WorkerNo:"
      Height          =   255
      Index           =   0
      Left            =   -720
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "FrmPayslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public constr As String
'Public newflag As Boolean
Dim WithEvents payslipRS As Recordset
Attribute payslipRS.VB_VarHelpID = -1
Dim WithEvents workersRS As Recordset
Attribute workersRS.VB_VarHelpID = -1

Private Sub cmdAdd_Click()
Call Enabletext
Call payslipgen
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
cmdDelete.Enabled = True
Dim h As String
Dim FOUND As Boolean
h = InputBox("Enter Worker's Employement No")
Dim db As Connection
Dim cntrl As Control
constr = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open constr

Set workersRS = Nothing

db.Close
db.Open constr

Set workersRS = New Recordset
workersRS.Open "select * from workers", db, adOpenStatic, adLockOptimistic

workersRS.MoveFirst
 While workersRS.EOF = False And FOUND = False
If h = workersRS!WorkerNo Then
Me.txtWorkerNo.Text = workersRS!WorkerNo
txtDepartment.Text = workersRS!Department
txtBasicSalary.Text = workersRS!BasicSalary
txtWorkerCategory.Text = workersRS!WorkerCategory
txtTime.Text = Time & ""
txtDated.Text = Date & ""
FrmDeductions.txtdeds = Me.txtWorkerCategory
FrmAllowances.txtAllowances.Text = Me.txtWorkerCategory
FOUND = True
Exit Sub
End If
workersRS.MoveNext

If workersRS.EOF = True And FOUND = False Then
MsgBox "NOT A VALID WORKER NO"
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Exit Sub
End If
Wend

End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Payslip Operations ,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub

Private Sub cmdDelete_Click()
Call clear
Call Disabletext
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
cmdDelete.Enabled = False
End Sub

Private Sub cmdhide_Click()
Paysliplist.Enabled = False
Paysliplist.Visible = False
Me.cmdhide.Enabled = False
Me.cmdhide.Visible = False
cmdRefresh.Enabled = True
cmdRefresh.Visible = True
End Sub

Private Sub cmdRefresh_Click()
Paysliplist.Enabled = True
Paysliplist.Visible = True
Me.cmdhide.Enabled = True
Me.cmdhide.Visible = True
cmdRefresh.Enabled = False
cmdRefresh.Visible = False
End Sub
Private Sub clear()
Me.txtWorkerNo.Text = ""
Me.txtAllowances.Text = ""
Me.txtBasicSalary.Text = ""
Me.txtDated.Text = ""
Me.txtDeductions.Text = ""
Me.txtDepartment.Text = ""
Me.txtNetSalary.Text = ""
Me.txtNetSalary.Text = ""
Me.txtPayslipNo.Text = ""
Me.txtPrintedBy.Text = ""
Me.txtTime.Text = ""
Me.txtWorkerCategory.Text = ""
End Sub
Private Sub cmdUpdate_Click()

Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
cmdDelete.Enabled = False

If Me.txtWorkerNo.Text = "" Then
MsgBox "Enter Customer FirstName", vbCritical
txtWorkerNo.SetFocus
Else

If Me.txtAllowances.Text = "" Then
MsgBox "Enter Allowances", vbCritical
txtAllowances.SetFocus
Else

If Me.txtBasicSalary.Text = "" Then
MsgBox "Enter Customer FirstName", vbCritical
txtBasicSalary.SetFocus
Else

If Me.txtDated.Text = "" Then
MsgBox "Enter Dated", vbCritical
txtDated.SetFocus
Else

If Me.txtDeductions.Text = "" Then
MsgBox "Enter Customer Deductions", vbCritical
txtDeductions.SetFocus
Else

If Me.txtDepartment.Text = "" Then
MsgBox "Enter Customer Department", vbCritical
txtDepartment.SetFocus
Else

If Me.txtNetSalary.Text = "" Then
MsgBox "Enter Customer NetSalary", vbCritical
txtNetSalary.SetFocus
Else



If Me.txtPayslipNo.Text = "" Then
MsgBox "Enter Customer FirstName", vbCritical
txtPayslipNo.SetFocus
Else

If Me.txtPrintedBy.Text = "" Then
MsgBox "Enter Customer FirstName", vbCritical
txtPrintedBy.SetFocus
Else

If Me.txtTime.Text = "" Then
MsgBox "Enter Customer FirstName", vbCritical
txtTime.SetFocus
Else
payslipRS.AddNew
payslipRS!WorkerNo = Me.txtWorkerNo
payslipRS!Allowances = Me.txtAllowances
payslipRS!BasicSalary = Me.txtBasicSalary
payslipRS!Dated = Me.txtDated
payslipRS!Deductions = Me.txtDeductions
payslipRS!Department = Me.txtDepartment
payslipRS!NetSalary = Me.txtNetSalary
'!NetSalary = Me.txtNetSalary
payslipRS!PayslipNo = Me.txtPayslipNo
payslipRS!PrintedBy = Me.txtPrintedBy
payslipRS!Time = Me.txtTime
payslipRS.Update
payslipRS.Requery
MsgBox "Workers Payslip will print in three seconds", vbCritical
FRMPAYSLIPPS.Show
FRMPAYSLIPPS.Label1.Caption = Me.txtWorkerCategory
FRMPAYSLIPPS.Label2.Caption = Me.txtWorkerNo
FRMPAYSLIPPS.Label4.Caption = Me.txtBasicSalary
FRMPAYSLIPPS.Label9.Caption = Me.txtDated
FRMPAYSLIPPS.Label6.Caption = Me.txtDeductions
FRMPAYSLIPPS.Label3.Caption = Me.txtDepartment
FRMPAYSLIPPS.Label7.Caption = Me.txtNetSalary
FRMPAYSLIPPS.Label8.Caption = Me.txtPayslipNo
FRMPAYSLIPPS.Label11.Caption = Me.txtPrintedBy
FRMPAYSLIPPS.Label10.Caption = Me.txtTime
FRMPAYSLIPPS.Label5.Caption = Me.txtAllowances
Call clear
Call Disabletext
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
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


End Sub

Private Sub Form_Load()
Dim db As Connection
Dim cntrl As Control
constr = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open constr

Set payslipRS = Nothing

db.Close
db.Open constr

Set payslipRS = New Recordset
payslipRS.Open "select * from payslip", db, adOpenStatic, adLockOptimistic

Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
cmdDelete.Enabled = False
Me.txtDated.Text = Date & ""
'Set Paysliplist.DataSource = OpenRecordset("select * from Payslip")
Paysliplist.Enabled = False
Paysliplist.Visible = False
Me.cmdhide.Enabled = False
Me.cmdhide.Visible = False
cmdRefresh.Enabled = True
cmdRefresh.Visible = True
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

Private Sub txtallowances_Change()
Dim f As Double
Dim g As Double
Dim j As Double
Dim l As Double
f = Val(Me.txtAllowances.Text)
g = Val(Me.txtBasicSalary.Text)
j = Val(Me.txtDeductions.Text)
l = Val(Me.txtNetSalary.Text)
l = (g - j) + f
Me.txtNetSalary.Text = "" & (l)
End Sub

Private Sub txtDeductions_Change()
Dim f As Double
Dim g As Double
Dim j As Double
Dim l As Double
f = Val(Me.txtAllowances.Text)
g = Val(Me.txtBasicSalary.Text)
j = Val(Me.txtDeductions.Text)
l = Val(Me.txtNetSalary.Text)
l = (g - j) + f
Me.txtNetSalary.Text = "" & (l)
End Sub

Private Sub txtNetSalary_Click()
Dim f As Double
Dim g As Double
Dim j As Double
Dim l As Double
f = Val(Me.txtAllowances.Text)
g = Val(Me.txtBasicSalary.Text)
j = Val(Me.txtDeductions.Text)
l = Val(Me.txtNetSalary.Text)
l = (g - j) + f
Me.txtNetSalary.Text = "" & (l)
End Sub
Private Sub payslipgen()
Dim nam As Long
Dim adm As String
Dim s As String
Dim mth, yearpart As String
If payslipRS.BOF = True Then
Me.txtPayslipNo.Text = "20154/13/2013"
Else
payslipRS.MoveFirst
payslipRS.MoveLast
nam = payslipRS.RecordCount + 1
mth = Format(Now, "mmmm - yyyy")
yearpart = Right(mth, 2)
s = "20154"
adm = (s & "/" & yearpart & "/" & nam)
Me.txtPayslipNo.Text = adm
End If
End Sub

