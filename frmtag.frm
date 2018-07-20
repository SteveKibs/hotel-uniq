VERSION 5.00
Begin VB.Form frmtag 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   5640
      Picture         =   "frmtag.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   4395
      TabIndex        =   11
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TypeofVisit:"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoofDays:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CustomerNo:"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Names:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmtag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
