VERSION 5.00
Begin VB.Form FrmAllowances 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "EMPLOYEE ALLOWANCES"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF00FF&
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   5775
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
         TabIndex        =   19
         Top             =   360
         Width           =   1095
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
         TabIndex        =   18
         Top             =   360
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
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
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Caption         =   "ALLOWANCES"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      Begin VB.ComboBox cbocategory 
         Height          =   315
         Left            =   3480
         TabIndex        =   21
         Text            =   "Select Worker Category"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtTotalAllowances 
         DataField       =   "TotalAllowances"
         Height          =   285
         Left            =   2085
         TabIndex        =   14
         Top             =   2520
         Width           =   1320
      End
      Begin VB.TextBox txtTransport 
         DataField       =   "Transport"
         Height          =   285
         Left            =   2085
         TabIndex        =   12
         Top             =   2160
         Width           =   1320
      End
      Begin VB.TextBox txtRisk 
         DataField       =   "Risk"
         Height          =   285
         Left            =   2085
         TabIndex        =   10
         Top             =   1755
         Width           =   1320
      End
      Begin VB.TextBox txtMedical 
         DataField       =   "Medical"
         Height          =   285
         Left            =   2085
         TabIndex        =   8
         Top             =   1380
         Width           =   1320
      End
      Begin VB.TextBox txtHouse 
         DataField       =   "House"
         Height          =   285
         Left            =   2085
         TabIndex        =   6
         Top             =   1005
         Width           =   1320
      End
      Begin VB.TextBox txtEntertainment 
         DataField       =   "Entertainment"
         Height          =   285
         Left            =   2085
         TabIndex        =   4
         Top             =   615
         Width           =   1320
      End
      Begin VB.TextBox txtWorkerCategory 
         DataField       =   "WorkerCategory"
         Height          =   285
         Left            =   2085
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtallowances 
         Height          =   285
         Left            =   4320
         TabIndex        =   22
         Text            =   "all"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TotalAllowances:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   13
         Top             =   2565
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transport:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   2190
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Risk:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medical:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1425
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entertainment:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WorkerCategory:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   285
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JUBELEE BEACH HOTEL EMPLOYEE ALLOWANCES"
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
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "FrmAllowances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public constr As String
'Public newflag As Boolean
Dim WithEvents allowancesRS As Recordset
Attribute allowancesRS.VB_VarHelpID = -1
Private Sub cbocategory_Click()
Me.txtWorkerCategory.Enabled = True
Me.txtWorkerCategory.Text = cbocategory.Text
End Sub

Private Sub cmdAdd_Click()
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
Me.cmdDelete.Enabled = False
Me.cbocategory.SetFocus
End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Allowance Operations ,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub
Private Sub clear()
Me.txtEntertainment.Text = ""
Me.txtHouse.Text = ""
Me.txtMedical.Text = ""
Me.txtRisk.Text = ""
Me.txtTotalAllowances.Text = ""
Me.txtTransport.Text = ""
Me.txtWorkerCategory.Text = ""
End Sub

Private Sub cmdDelete_Click()
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
Dim h As String
Dim FOUND As Boolean
Dim Y As String
End Sub

Private Sub cmdUpdate_Click()
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
Me.cmdDelete.Enabled = False
'With openrecordset("select * from Allowances")

If Me.txtEntertainment.Text = "" Then
MsgBox "Please Enter Entertainment", vbCritical
txtEntertainment.SetFocus
Else

If Me.txtHouse.Text = "" Then
MsgBox "Please Enter House ", vbCritical
txtHouse.SetFocus
Else

If Me.txtMedical.Text = "" Then
MsgBox "Please Enter  Medical", vbCritical
txtMedical.SetFocus
Else

If Me.txtRisk.Text = "" Then
MsgBox "Please Enter Risk ", vbCritical
txtRisk.SetFocus
Else

If Me.txtTotalAllowances.Text = "" Then
MsgBox "Please Enter Total Allowances", vbCritical
txtTotalAllowances.SetFocus
Else

If Me.txtTransport.Text = "" Then
MsgBox "Please Enter Transport ", vbCritical
txtTransport.SetFocus
Else

If Me.txtWorkerCategory.Text = "" Then
MsgBox "Please Enter Worker Category ", vbCritical
txtWorkerCategory.SetFocus
Else

allowancesRS.AddNew
allowancesRS!Entertainment = Me.txtEntertainment.Text
allowancesRS!House = Me.txtHouse.Text
allowancesRS!Medical = Me.txtMedical.Text
allowancesRS!Risk = Me.txtRisk.Text
allowancesRS!TotalAllowances = Me.txtTotalAllowances.Text
allowancesRS!Transport = Me.txtTransport.Text
allowancesRS!WCategory = Me.txtWorkerCategory.Text
allowancesRS.Update
allowancesRS.Requery
MsgBox "Allowances Updated Successfully"
Call clear
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = True
End If
End If
End If
End If
End If
End If
End If
'End With
End Sub

Private Sub Form_Load()
Dim db As Connection
Dim cntrl As Control
constr = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open constr

Set allowancesRS = Nothing

db.Close
db.Open constr

Set allowancesRS = New Recordset
allowancesRS.Open "select * from allowances", db, adOpenStatic, adLockOptimistic

Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = True
Me.txtWorkerCategory.Enabled = False
Me.cbocategory.AddItem "Permanent"
Me.cbocategory.AddItem "Contract"
Me.cbocategory.AddItem "Casual"
End Sub


Private Sub txtallowances_Change()
Dim FOUND As Boolean
Dim z As String
'With openrecordset("select * from Allowances")
z = txtAllowances.Text
allowancesRS.MoveFirst
While Not allowancesRS.EOF And FOUND = False
If z = allowancesRS!WCategory Then
FOUND = True
Me.Show
FrmPayslip.txtAllowances.Text = allowancesRS!TotalAllowances & ""
Unload FrmAllowances
Exit Sub
End If
allowancesRS.MoveNext
If allowancesRS.EOF And FOUND = False Then
MsgBox ("ALLOWANCE NOT VALID")
End If
Wend

End Sub

Private Sub txtTransport_Change()
Dim f As Double
Dim h As Double
Dim b As Double
Dim l As Double
Dim O As Double
Dim p As Double
f = Val(Me.txtEntertainment.Text)
h = Val(Me.txtHouse.Text)
b = Val(Me.txtMedical.Text)
l = Val(Me.txtRisk.Text)
O = Val(Me.txtTransport.Text)
p = Val(Me.txtTotalAllowances.Text)
p = f + h + b + l + O
Me.txtTotalAllowances.Text = "" & (p)

End Sub
