VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JUBELEE BEACH HOTEL LOGIN"
   ClientHeight    =   1605
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   948.287
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      Height          =   390
      Left            =   375
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000FF00&
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   1500
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CONSTR As String
Dim WithEvents usersRS As Recordset
Attribute usersRS.VB_VarHelpID = -1
Private Sub CMDCANCEL_Click()
Dim t As String
t = MsgBox("The System will end Users Login Operations ,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub

Private Sub cmdOK_Click()
Dim FOUND As Boolean
'Checking whether the correct login details are provided

usersRS.MoveFirst
Do While usersRS.EOF = False And FOUND = False
usersRS.MoveFirst
If txtUserName.Text = usersRS!Name And txtPassword.Text = usersRS!Pass Then
FOUND = True
MsgBox "Welcome! You have logged in ", vbInformation
frmmain.Show
Unload Me
End If
usersRS.MoveNext
Loop
If usersRS.EOF And FOUND = False Then
If MsgBox("Access Denied.Retry?", vbYesNo) = vbYes Then
Me.txtPassword = ""
Exit Sub
Else
End
End If
End If

End Sub

Private Sub Form_Load()
Dim db As Connection
Dim cntrl As Control
CONSTR = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open CONSTR

Set usersRS = Nothing

db.Close
db.Open CONSTR

Set usersRS = New Recordset
usersRS.Open "select * from users", db, adOpenStatic, adLockOptimistic
End Sub
