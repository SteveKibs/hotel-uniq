VERSION 5.00
Begin VB.Form FRMORDER 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   Picture         =   "FRMORDER.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   1800
      Top             =   840
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label7"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label6"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label4"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   2  'Dash
      BorderWidth     =   8
      X1              =   0
      X2              =   5880
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2775
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FoodName:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   3165
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FoodCost:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FoodNo:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   3915
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OrderNo:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   4305
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OrderTime:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReceiptNo:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   1815
   End
End
Attribute VB_Name = "FRMORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Timer1.Interval = 25 Then
Me.PrintForm
Unload Me
Else
End If

End Sub
