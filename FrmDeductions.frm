VERSION 5.00
Begin VB.Form FrmDeductions 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Height          =   2775
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   5775
      Begin VB.TextBox txttake 
         Height          =   285
         Left            =   3840
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtdeds 
         Height          =   285
         Left            =   3720
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox cbocategory 
         Height          =   315
         Left            =   3450
         TabIndex        =   18
         Text            =   "Select Worker Category"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtTotalDeduction 
         DataField       =   "TotalDeduction"
         Height          =   285
         Left            =   1605
         TabIndex        =   17
         Top             =   2265
         Width           =   1320
      End
      Begin VB.TextBox txtPAYE 
         DataField       =   "PAYE"
         Height          =   285
         Left            =   1605
         TabIndex        =   15
         Top             =   1875
         Width           =   1320
      End
      Begin VB.TextBox txtNSSF 
         DataField       =   "NSSF"
         Height          =   285
         Left            =   1605
         TabIndex        =   13
         Top             =   1500
         Width           =   1320
      End
      Begin VB.TextBox txtNHIF 
         DataField       =   "NHIF"
         Height          =   285
         Left            =   1605
         TabIndex        =   11
         Top             =   1125
         Width           =   1320
      End
      Begin VB.TextBox txtINCOMETAX 
         DataField       =   "INCOMETAX"
         Height          =   285
         Left            =   1605
         TabIndex        =   9
         Top             =   735
         Width           =   1320
      End
      Begin VB.TextBox txtWorkerCategory 
         DataField       =   "WorkerCategory"
         Height          =   285
         Left            =   1605
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TotalDeduction:"
         Height          =   255
         Index           =   5
         Left            =   -240
         TabIndex        =   16
         Top             =   2310
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAYE:"
         Height          =   255
         Index           =   4
         Left            =   -240
         TabIndex        =   14
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NSSF:"
         Height          =   255
         Index           =   3
         Left            =   -240
         TabIndex        =   12
         Top             =   1545
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NHIF:"
         Height          =   255
         Index           =   2
         Left            =   -240
         TabIndex        =   10
         Top             =   1170
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INCOMETAX:"
         Height          =   255
         Index           =   1
         Left            =   -240
         TabIndex        =   8
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WorkerCategory:"
         Height          =   255
         Index           =   0
         Left            =   -240
         TabIndex        =   6
         Top             =   405
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF00FF&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   3600
      Width           =   5775
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Add Allowances"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1815
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1335
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
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
         Height          =   300
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JUBELEE BEACH HOTEL  EMPLOYEE DEDUCTIONS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   0
      TabIndex        =   19
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "FrmDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public constr As String
'Public newflag As Boolean
Dim WithEvents deductionsRS As Recordset
Attribute deductionsRS.VB_VarHelpID = -1
Private Sub cbocategory_Click()
Me.txtWorkerCategory.Text = cbocategory.Text
End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Deduction Operations ,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If

End Sub
Private Sub clear()
Me.txtINCOMETAX.Text = ""
Me.txtNHIF.Text = ""
Me.txtNSSF.Text = ""
Me.txtPAYE.Text = ""
Me.txtTotalDeduction.Text = ""
Me.txtWorkerCategory.Text = ""
End Sub
Private Sub cmdUpdate_Click()

If Me.txtINCOMETAX.Text = "" Then
MsgBox "Please Enter INCOME TAX", vbCritical
txtINCOMETAX.SetFocus
Else

If Me.txtNHIF.Text = "" Then
MsgBox "Please Enter NHIF", vbCritical
txtNHIF.SetFocus
Else

If Me.txtNSSF.Text = "" Then
MsgBox "Please Enter tNSSF", vbCritical
txtNSSF.SetFocus
Else

If Me.txtPAYE.Text = "" Then
MsgBox "Please Enter PAYE", vbCritical
txtPAYE.SetFocus
Else

If Me.txtTotalDeduction.Text = "" Then
MsgBox "Please Enter TotalDeduction", vbCritical
txtTotalDeduction.SetFocus
Else

If Me.txtWorkerCategory.Text = "" Then
MsgBox "Please EnterWorker Category", vbCritical
txtWorkerCategory.SetFocus
Else

deductionsRS.AddNew
deductionsRS!INCOMETAX = Me.txtINCOMETAX.Text
deductionsRS!NHIF = Me.txtNHIF.Text
deductionsRS!NSSF = Me.txtNSSF.Text
deductionsRS!PAYE = Me.txtPAYE.Text
deductionsRS!TotalDeductions = Me.txtTotalDeduction.Text
deductionsRS!WCategory = Me.txtWorkerCategory.Text
deductionsRS.Update
deductionsRS.Requery
MsgBox "Deductions Updated Successfully"
Call clear
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

Set deductionsRS = Nothing

db.Close
db.Open constr

Set deductionsRS = New Recordset
deductionsRS.Open "select * from deductions", db, adOpenStatic, adLockOptimistic

Me.txtWorkerCategory.Enabled = False
Me.cbocategory.AddItem "Permanent"
Me.cbocategory.AddItem "Contract"
Me.cbocategory.AddItem "Casual"
End Sub

Private Sub txtdeds_Change()
Dim FOUND As Boolean
Dim z As String
z = Me.txtdeds.Text
deductionsRS.MoveFirst
While Not deductionsRS.EOF And FOUND = False
If z = deductionsRS!WCategory Then
'txttake.Text = !TotalDeductions
FOUND = True
Me.Show
FrmPayslip.txtDeductions = deductionsRS!TotalDeductions
Unload FrmDeductions
Exit Sub
End If
deductionsRS.MoveNext
If deductionsRS.EOF And FOUND = False Then
MsgBox ("DEDUCTION NOT VALID")
End If
Wend

End Sub

