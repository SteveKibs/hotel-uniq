VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmsplash 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   Picture         =   "frmsplash.frx":0000
   ScaleHeight     =   4695
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   240
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If ProgressBar1.Value < ProgressBar1.Max Then
ProgressBar1.Value = ProgressBar1 + 1
Else
Timer1.Enabled = False
Unload Me
frmLogin.Show
End If
End Sub
