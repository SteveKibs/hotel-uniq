VERSION 5.00
Begin VB.Form Frmbooking 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "BOOKING OFFICE"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbobookingof 
      Height          =   315
      Left            =   3480
      TabIndex        =   13
      Text            =   "Select Booking item "
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtDate 
      DataField       =   "Date"
      Height          =   285
      Left            =   2295
      TabIndex        =   12
      Top             =   1110
      Width           =   1320
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000FFFF&
      Caption         =   "&New Booking"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Refresh"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   975
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   975
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
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtReceiptNo 
      DataField       =   "ReceiptNo"
      Height          =   285
      Left            =   2295
      TabIndex        =   6
      Top             =   195
      Width           =   3375
   End
   Begin VB.TextBox txtBookingOf 
      DataField       =   "BookingOf"
      Height          =   285
      Left            =   2295
      TabIndex        =   4
      Top             =   1500
      Width           =   1200
   End
   Begin VB.TextBox txtCustomerNo 
      DataField       =   "CustomerNo"
      Height          =   285
      Left            =   2295
      TabIndex        =   1
      Top             =   735
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReceiptNo:"
      Height          =   255
      Index           =   3
      Left            =   450
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BookingOf:"
      Height          =   255
      Index           =   2
      Left            =   450
      TabIndex        =   3
      Top             =   1545
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Index           =   1
      Left            =   450
      TabIndex        =   2
      Top             =   1155
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Frmbooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim found As Boolean
Dim h As String
cmdAdd.Enabled = False
cmdupdaate.Enabled = True
h = MsgBox("Enter Customer Payment Receipt no")
Me.txtReceiptNo.SetFocus
With openrecordset("select * from Accounts")
While found = False And .EOF = False
.MoveFirst
If h = !ReceiptNo Then
found = True
Me.txtBookingOf.Text = !Paymentof & ""
Unload Frmbooking
End If
Wend
End With
End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Booking Operations,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub

Private Sub cmdUpdate_Click()
With openrecordset("select * from BookingOffice")

If Me.txtCustomerNo.Text = "" Then
MsgBox "Enter Customer No", vbCritical
txtCustomerNo.SetFocus
End If

If Me.txtDate.Text = "" Then
MsgBox "Enter Date", vbCritical
txtDate.SetFocus
End If

If Me.txtBookingOf.Text = "" Then
MsgBox "Enter BookingOf", vbCritical
txtBookingOf.SetFocus
End If

If Me.txtReceiptNo.Text = "" Then
MsgBox "Enter ReceiptNo", vbCritical
txtReceiptNo.SetFocus
Else
.AddNew
!CustomerNo = Me.txtCustomerNo.Text
!Date = Me.txtDate.Text
!BookingOf = Me.txtBookingOf.Text
!ReceiptNo = Me.txtReceiptNo.Text
.Update
.Requery
MsgBox "You booked a room"
Call clear
End If
End With
End Sub
Private Sub clear()
 Me.txtCustomerNo.Text = ""
Me.txtDate.Text = ""
Me.txtBookingOf.Text = ""
Me.txtReceiptNo.Text = ""
End Sub

Private Sub Form_Load()
cbobookingof.AddItem "Tour Site"
cbobookingof.AddItem "Accomodation"
cbobookingof.AddItem "Transport"
End Sub
