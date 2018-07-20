VERSION 5.00
Begin VB.Form FRMPAYSLIPPS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   Picture         =   "FRMPAYSLIPPS.frx":0000
   ScaleHeight     =   6600
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   2400
      Top             =   1080
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label11"
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label9"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label7"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label6"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label4"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   2400
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Worker Category:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   10
      Left            =   -105
      TabIndex        =   10
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WorkerNo:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   9
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PrintedBy:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   8
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PayslipNo:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   525
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NetSalary:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   870
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allowances:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   945
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BasicSalary:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1005
   End
End
Attribute VB_Name = "FRMPAYSLIPPS"
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
