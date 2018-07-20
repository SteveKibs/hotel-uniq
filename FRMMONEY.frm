VERSION 5.00
Begin VB.Form FRMMONEY 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   Picture         =   "FRMMONEY.frx":0000
   ScaleHeight     =   5895
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   2760
      Top             =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   11280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   255
      Left            =   9600
      TabIndex        =   18
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      Height          =   255
      Left            =   9120
      TabIndex        =   17
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   255
      Left            =   9120
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   255
      Left            =   9480
      TabIndex        =   15
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   10
      Left            =   8520
      TabIndex        =   10
      Top             =   720
      Width           =   570
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   9
      Left            =   8520
      TabIndex        =   9
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReceiptNo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   8
      Left            =   8400
      TabIndex        =   8
      Top             =   5400
      Width           =   1185
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   870
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   930
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AmountPaid:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckNo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   4
      Left            =   8520
      TabIndex        =   4
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CurrencyType:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   4800
      TabIndex        =   3
      Top             =   3120
      Width           =   1530
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PaymentMode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1560
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paymentof:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1350
   End
End
Attribute VB_Name = "FRMMONEY"
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
