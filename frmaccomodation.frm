VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmaccomodation 
   BackColor       =   &H0080FF80&
   Caption         =   $"frmaccomodation.frx":0000
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   4440
   End
   Begin VB.TextBox txtCustomerNo 
      Height          =   285
      Left            =   7440
      TabIndex        =   30
      Top             =   1080
      Width           =   2895
   End
   Begin MSDataListLib.DataList HouseName 
      Bindings        =   "frmaccomodation.frx":008B
      Height          =   2595
      Left            =   840
      TabIndex        =   28
      Top             =   1200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4577
      _Version        =   393216
      BackColor       =   16744576
      ListField       =   "HouseName"
      Object.DataMember      =   "Houses"
   End
   Begin VB.TextBox TXTHouseMaster 
      Height          =   285
      Left            =   7440
      TabIndex        =   27
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtDurationb 
      Height          =   285
      Left            =   7440
      TabIndex        =   25
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   855
      Left            =   3960
      TabIndex        =   17
      Top             =   5640
      Width           =   9975
      Begin VB.CommandButton cmdclear 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Clear House"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   2175
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
         Height          =   540
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Height          =   540
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Cancel Operation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Assign House"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox txtReceiptNo 
      DataField       =   "ReceiptNo"
      Height          =   285
      Left            =   7455
      TabIndex        =   16
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox txtExpeiryDate 
      DataField       =   "ExpeiryDate"
      Height          =   285
      Left            =   7455
      TabIndex        =   14
      Top             =   4200
      Width           =   1320
   End
   Begin VB.TextBox txtEntertainment 
      DataField       =   "Entertainment"
      Height          =   285
      Left            =   7455
      TabIndex        =   12
      Top             =   3825
      Width           =   3375
   End
   Begin VB.TextBox txtBookedOn 
      DataField       =   "BookedOn"
      Height          =   285
      Left            =   7455
      TabIndex        =   10
      Top             =   3450
      Width           =   1320
   End
   Begin VB.TextBox txtCost 
      DataField       =   "Cost"
      Height          =   285
      Left            =   7455
      TabIndex        =   8
      Top             =   3060
      Width           =   1320
   End
   Begin VB.TextBox txtOccupants 
      DataField       =   "Occupants"
      Height          =   285
      Left            =   7455
      TabIndex        =   6
      Top             =   2205
      Width           =   1320
   End
   Begin VB.TextBox txtRoomNo 
      DataField       =   "RoomNo"
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      Top             =   1560
      Width           =   1080
   End
   Begin VB.TextBox txtTypeofRooms 
      DataField       =   "TypeofRooms"
      Height          =   285
      Left            =   7455
      TabIndex        =   3
      Top             =   1800
      Width           =   1680
   End
   Begin VB.TextBox txtHouseName 
      DataField       =   "HouseName"
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   825
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "RoomNo"
      Height          =   255
      Left            =   3960
      TabIndex        =   31
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo"
      Height          =   255
      Left            =   6000
      TabIndex        =   29
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "House Master"
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Duration Booked"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Assign Tourists to THEIR RESPECTIVE  laxurious Houses"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   840
      TabIndex        =   22
      Top             =   360
      Width           =   14535
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReceiptNo:"
      Height          =   255
      Index           =   8
      Left            =   5610
      TabIndex        =   15
      Top             =   4635
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ExpeiryDate:"
      Height          =   255
      Index           =   7
      Left            =   5610
      TabIndex        =   13
      Top             =   4245
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entertainment:"
      Height          =   255
      Index           =   6
      Left            =   5610
      TabIndex        =   11
      Top             =   3870
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BookedOn:"
      Height          =   255
      Index           =   5
      Left            =   5610
      TabIndex        =   9
      Top             =   3495
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost:"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   7
      Top             =   3105
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupants:"
      Height          =   255
      Index           =   3
      Left            =   5610
      TabIndex        =   5
      Top             =   2250
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TypeofRooms:"
      Height          =   255
      Index           =   1
      Left            =   5610
      TabIndex        =   2
      Top             =   1845
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HouseName:"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   870
      Width           =   1815
   End
End
Attribute VB_Name = "frmaccomodation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Me.HouseName.Enabled = True
Me.HouseName.Visible = True

End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Accomodation Operations,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdUpdate_Click()
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
With openrecordset("select*from accomodation")
If Me.txtBookedOn.Text = "" Then
MsgBox "Enter Booked On ", vbCritical
txtBookedOn.SetFocus
End If

If Me.txtCost.Text = "" Then
MsgBox "Enter Cost ", vbCritical
txtCost.SetFocus
End If

If Me.txtEntertainment.Text = "" Then
MsgBox "Enter Entertainment ", vbCritical
txtEntertainment.SetFocus
End If

If Me.txtExpeiryDate.Text = "" Then
MsgBox "Enter ExpeiryDate ", vbCritical
txtExpeiryDate.SetFocus
End If

If Me.txtHouseName.Text = "" Then
MsgBox "Enter HouseName ", vbCritical
txtHouseName.SetFocus
End If

If Me.txtOccupants.Text = "" Then
MsgBox "Enter Occupants", vbCritical
txtOccupants.SetFocus
End If

If Me.txtReceiptNo.Text = "" Then
MsgBox "Enter ReceiptNo", vbCritical
txtReceiptNo.SetFocus
End If

If Me.txtTypeofRooms.Text = "" Then
MsgBox "Enter TypeofRooms", vbCritical
txtTypeofRooms.SetFocus
End If

If Me.txtRoomNo.Text = "" Then
MsgBox "Enter RoomNo", vbCritical
txtRoomNo.SetFocus
Else
.AddNew

!BookedOn = Me.txtBookedOn.Text
!Cost = Me.txtCost.Text
!Entertainment = Me.txtEntertainment.Text
!ExpeiryDate = Me.txtExpeiryDate.Text
!HouseName = Me.txtHouseName.Text
!Occupants = Me.txtOccupants.Text
!ReceiptNo = Me.txtReceiptNo.Text
!TypeofRooms = Me.txtTypeofRooms.Text
!RoomNo = Me.txtRoomNo.Text
.Update
.Requery
MsgBox "accomodated"
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Call clear
'Me.txtBookedOn
'Me.txtCost
'Me.txtEntertainment
'Me.txtExpeiryDate
'Me.txtHouseName
'Me.txtOccupants
'Me.txtReceiptNo
'Me.txtTypeofRooms
'Me.txtRoomNo
'
End If
End With
End Sub
Private Sub clear()
Me.txtBookedOn.Text = ""
Me.txtCost.Text = ""
Me.txtEntertainment.Text = ""
Me.txtExpeiryDate.Text = ""
Me.txtHouseName.Text = ""
Me.txtOccupants.Text = ""
Me.txtReceiptNo.Text = ""
Me.txtTypeofRooms.Text = ""
Me.txtRoomNo.Text = ""
Me.txtCustomerNo.Text = ""
Me.txtDurationb.Text = ""
End Sub

Private Sub Form_Load()
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.HouseName.Enabled = False
Me.HouseName.Visible = False
With openrecordset("select * from Houses")
Me.txtBookedOn.Text = Date & ""
''cbohname.AddItem !HouseName
'cbohname.AddItem !HouseName
'Set HouseName.DataSource = openrecordset("select HouseName from Houses")
'Set RoomNo.DataSource = openrecordset("select RoomNo from Houses")
'Set List1.DataSource = openrecordset("select * from accomodation")
End With
End Sub

Private Sub HouseName_Click()
txtHouseName.Text = HouseName.BoundText
End Sub

Private Sub RoomNo_Click()
Me.txtRoomNo.Text = HouseName.Index
End Sub

Private Sub Timer1_Timer()
txtBookedOn.Text = Format(Now(), "dddd, yyyy mmmm dd")
End Sub

Private Sub FINDER()
Dim h As String
Dim found As Boolean
h = Me.txtHouseName
With openrecordset("select * from Houses")
.MoveFirst
Do While .EOF = False And found = False
If h = !HouseName Then
Me.txtEntertainment.Text = !Entertainment & ""
Me.txtCost.Text = !BookingAmount & ""
Me.txtOccupants.Text = !Occupants & ""
Me.txtHouseName.Text = !HouseName & ""
Me.txtTypeofRooms.Text = !TypeofRooms & ""
Me.txtRoomNo.Text = !Houseno & ""
Me.Show
found = True
Exit Sub
End If
.MoveNext
Loop
If .EOF = True And found = False Then
MsgBox "NOT A VALID HOUSE"
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Exit Sub
End If
End With
End Sub

Private Sub txtRoomNo_Change()
Dim found As Boolean
Dim h As String
h = InputBox("enter customer no here!")
With openrecordset("select * from Reception")
Do While .EOF = False And found = False
.MoveFirst
If h = !CustomerNo Then
found = True
Me.txtCustomerNo.Text = !CustomerNo
Me.TXTHouseMaster.Text = !FirstName & "     " & !LastName
Me.txtDurationb.Text = !NoofDays
Me.Show
Call FINDER
Exit Sub
End If
.MoveNext
Loop
If .EOF = True And found = False Then
Exit Sub
Else
End If
End With
End Sub
