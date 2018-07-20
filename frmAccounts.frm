VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAccounts 
   BackColor       =   &H00FF00FF&
   Caption         =   "                                                          JUBELEE BEACH HOTEL FINANCE OFFICE"
   ClientHeight    =   11010
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   13440
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtCustomerNo 
         DataField       =   "CustomerNo"
         Height          =   285
         Left            =   1245
         TabIndex        =   20
         Top             =   180
         Width           =   3375
      End
      Begin VB.TextBox txtPaymentof 
         DataField       =   "Paymentof"
         Height          =   285
         Left            =   1245
         TabIndex        =   19
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtPaymentMode 
         DataField       =   "PaymentMode"
         Height          =   285
         Left            =   1245
         TabIndex        =   18
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox txtCurrencyType 
         DataField       =   "CurrencyType"
         Height          =   285
         Left            =   1245
         TabIndex        =   17
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtCheckNo 
         DataField       =   "CheckNo"
         Height          =   285
         Left            =   6765
         TabIndex        =   16
         Top             =   1455
         Width           =   2400
      End
      Begin VB.TextBox txtAmountPaid 
         DataField       =   "AmountPaid"
         Height          =   285
         Left            =   5805
         TabIndex        =   15
         Top             =   1965
         Width           =   1320
      End
      Begin VB.TextBox txtBalance 
         DataField       =   "Balance"
         Height          =   285
         Left            =   5805
         TabIndex        =   14
         Top             =   2340
         Width           =   1320
      End
      Begin VB.TextBox txtChange 
         DataField       =   "Change"
         Height          =   285
         Left            =   5805
         TabIndex        =   13
         Top             =   2715
         Width           =   1320
      End
      Begin VB.TextBox txtReceiptNo 
         DataField       =   "ReceiptNo"
         Height          =   285
         Left            =   5760
         TabIndex        =   12
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtTime 
         DataField       =   "Time"
         Height          =   285
         Left            =   7800
         TabIndex        =   11
         Top             =   240
         Width           =   1920
      End
      Begin VB.TextBox txtDate 
         DataField       =   "Date"
         Height          =   285
         Left            =   7800
         TabIndex        =   10
         Top             =   600
         Width           =   1920
      End
      Begin VB.OptionButton Optcash 
         BackColor       =   &H00FF00FF&
         Caption         =   "Cash"
         Height          =   255
         Left            =   2805
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Optcheque 
         BackColor       =   &H00FF00FF&
         Caption         =   "Cheque"
         Height          =   255
         Left            =   4245
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF00FF&
         Height          =   1815
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H00FF00FF&
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
            Height          =   420
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton cmdUpdate 
            BackColor       =   &H00FF00FF&
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
            Height          =   420
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FF00FF&
            Caption         =   "&Receive Payment"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.ComboBox CBOPAYMENTOF 
         Height          =   315
         Left            =   4605
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cbocurrency 
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2880
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2640
         Width           =   150
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CustomerNo:"
         Height          =   255
         Index           =   0
         Left            =   -600
         TabIndex        =   31
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paymentof:"
         Height          =   255
         Index           =   1
         Left            =   -600
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PaymentMode:"
         Height          =   255
         Index           =   2
         Left            =   -600
         TabIndex        =   29
         Top             =   990
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CurrencyType:"
         Height          =   255
         Index           =   3
         Left            =   -600
         TabIndex        =   28
         Top             =   1485
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CheckNo:"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   27
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AmountPaid:"
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   26
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   25
         Top             =   2385
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change:"
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReceiptNo:"
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   23
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         Height          =   255
         Index           =   9
         Left            =   6000
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   255
         Index           =   10
         Left            =   6000
         TabIndex        =   21
         Top             =   600
         Width           =   1815
      End
   End
   Begin MSDataGridLib.DataGrid AccountsOffice 
      Bindings        =   "frmAccounts.frx":0000
      Height          =   6855
      Left            =   -120
      TabIndex        =   0
      Top             =   3960
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   12091
      _Version        =   393216
      BackColor       =   16711935
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Accounts"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "CustomerNo"
         Caption         =   "CustomerNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Paymentof"
         Caption         =   "Paymentof"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "PaymentMode"
         Caption         =   "PaymentMode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CurrencyType"
         Caption         =   "CurrencyType"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "CheckNo"
         Caption         =   "CheckNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "AmountPaid"
         Caption         =   "AmountPaid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Balance"
         Caption         =   "Balance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Change"
         Caption         =   "Change"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "ReceiptNo"
         Caption         =   "ReceiptNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Time"
         Caption         =   "Time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Date"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public constr As String
Dim WithEvents accountsRS As Recordset
Attribute accountsRS.VB_VarHelpID = -1
Dim WithEvents receptionRS As Recordset
Attribute receptionRS.VB_VarHelpID = -1
Dim WithEvents menuRS As Recordset
Attribute menuRS.VB_VarHelpID = -1
Dim WithEvents ordersRS As Recordset
Attribute ordersRS.VB_VarHelpID = -1
'Dim WithEvents classesRS As Recordset
'Dim WithEvents invoiceRS As Recordset
'Dim WithEvents feesRS As Recordset
Private Sub cbocurrency_Click()
Me.txtCurrencyType.Text = cbocurrency.Text
'Me.cbocurrency.Visible = False
'Me.cbocurrency.Enabled = False
End Sub

Private Sub CBOPAYMENTOF_Click()
Me.txtPaymentof.Text = CBOPAYMENTOF.Text
'Me.CBOPAYMENTOF.Visible = False
'Me.CBOPAYMENTOF.Enabled = False
End Sub

Private Sub cmdAdd_Click()
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
If accountsRS.BOF = True Then
txtReceiptNo.Text = "21548745"
Call add
Exit Sub
Else
accountsRS.MoveFirst
accountsRS.MoveLast
txtReceiptNo.Text = (accountsRS!ReceiptNo + 1)
Call add
Exit Sub
End If
End Sub
Private Sub add()
Dim db As Connection
Dim cntrl As Control
constr = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open constr

Set receptionRS = Nothing

db.Close
db.Open constr

Set receptionRS = New Recordset
receptionRS.Open "select * from reception", db, adOpenStatic, adLockOptimistic

Dim h As String
Dim FOUND As Boolean
h = InputBox("Enter Customer No")
receptionRS.MoveFirst
Do While receptionRS.EOF = False And FOUND = False
If h = receptionRS!CustomerNo Then
 Me.txtCustomerNo.Text = receptionRS!CustomerNo
FOUND = True
Exit Sub
End If
receptionRS.MoveNext
Loop
If receptionRS.EOF = True And FOUND = False Then
MsgBox "NOT A VALID CUSTOMER NO"
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Exit Sub
End If

End Sub
Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Accounts Operations,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub

Private Sub cmdUpdate_Click()
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False

If Me.txtCustomerNo.Text = "" Then
MsgBox "Enter CustomerNo ", vbCritical
 Me.txtCustomerNo.SetFocus
End If

If Me.txtPaymentof.Text = "" Then
MsgBox "Enter Paymentof ", vbCritical
txtPaymentof.SetFocus
End If

If Me.txtPaymentMode.Text = "" Then
MsgBox "Enter PaymentMode ", vbCritical
txtPaymentMode.SetFocus
End If

If Me.txtCurrencyType.Text = "" Then
MsgBox "Enter CurrencyType ", vbCritical
txtCurrencyType.SetFocus
End If


If Me.txtAmountPaid.Text = "" Then
MsgBox "Enter AmountPaid ", vbCritical
txtAmountPaid.SetFocus
End If

If Me.txtBalance.Text = "" Then
MsgBox "Enter Balance ", vbCritical
txtBalance.SetFocus
End If

If Me.txtChange.Text = "" Then
MsgBox "Enter Change ", vbCritical
txtChange.SetFocus
End If

If Me.txtReceiptNo.Text = "" Then
MsgBox "EnterReceiptNo", vbCritical
txtReceiptNo.SetFocus
End If

If Me.txtTime.Text = "" Then
MsgBox "EnterTime", vbCritical
txtTime.SetFocus
End If

If Me.txtDate.Text = "" Then
MsgBox "EnterDate", vbCritical
txtDate.SetFocus

Else
accountsRS.AddNew
accountsRS!CustomerNo = Me.txtCustomerNo.Text
accountsRS!Paymentof = Me.txtPaymentof.Text
accountsRS!PaymentMode = Me.txtPaymentMode.Text
accountsRS!CurrencyType = Me.txtCurrencyType.Text
accountsRS!CheckNo = Me.txtCheckNo.Text
accountsRS!AmountPaid = Me.txtAmountPaid.Text
accountsRS!Balance = Me.txtBalance.Text
accountsRS!Change = Me.txtChange.Text
accountsRS!ReceiptNo = Me.txtReceiptNo.Text
accountsRS!Time = Me.txtTime.Text
accountsRS!Date = Me.txtDate.Text
accountsRS.Update
accountsRS.Requery
MsgBox "your account has been credited"
FRMMONEY.Show
FRMMONEY.Label1.Caption = Me.txtCustomerNo.Text
FRMMONEY.Label2.Caption = Me.txtPaymentof.Text
FRMMONEY.Label3.Caption = Me.txtPaymentMode.Text
FRMMONEY.Label4.Caption = Me.txtCurrencyType.Text
FRMMONEY.Label5.Caption = Me.txtCheckNo.Text
FRMMONEY.Label9.Caption = Me.txtAmountPaid.Text
FRMMONEY.Label10.Caption = Me.txtBalance.Text
FRMMONEY.Label11.Caption = Me.txtChange.Text
FRMMONEY.Label8.Caption = Me.txtReceiptNo.Text
FRMMONEY.Label6.Caption = Me.txtTime.Text
FRMMONEY.Label7.Caption = Me.txtDate.Text
Call clear
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.CBOPAYMENTOF.Visible = True
Me.CBOPAYMENTOF.Enabled = True
Me.cbocurrency.Visible = True
Me.cbocurrency.Enabled = True
txtCheckNo.Enabled = True
txtCheckNo.Visible = True

End If

End Sub

Private Sub clear()
 Me.txtCustomerNo.Text = ""
Me.txtPaymentof.Text = ""
Me.txtPaymentMode.Text = ""
Me.txtCurrencyType.Text = ""
Me.txtCheckNo.Text = ""
Me.txtAmountPaid.Text = ""
Me.txtBalance.Text = ""
Me.txtChange.Text = ""
Me.txtReceiptNo.Text = ""
Me.txtTime.Text = ""
Me.txtDate.Text = ""
End Sub



Private Sub Form_Load()
Dim db As Connection
Dim cntrl As Control
constr = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open constr

Set accountsRS = Nothing

db.Close
db.Open constr

Set accountsRS = New Recordset
accountsRS.Open "select * from accounts", db, adOpenStatic, adLockOptimistic

Me.CBOPAYMENTOF.Visible = True
Me.CBOPAYMENTOF.Enabled = True
Me.cbocurrency.Visible = True
Me.cbocurrency.Enabled = True
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.txtTime.Text = "" & Time
Me.txtDate.Text = "" & Date
cbocurrency.AddItem "Ksh"
cbocurrency.AddItem "Us Dollars"
cbocurrency.AddItem "Pounds"
cbocurrency.AddItem "Yarns"
CBOPAYMENTOF.AddItem "Food and Bevarages Order"
CBOPAYMENTOF.AddItem "Snacks"
End Sub

Private Sub Optcash_Click()
txtPaymentMode.Text = Optcash.Caption
txtCheckNo.Enabled = False
txtCheckNo.Visible = False
End Sub

Private Sub Optcheque_Click()
txtPaymentMode.Text = Optcheque.Caption
txtCheckNo.Enabled = True
txtCheckNo.Visible = True
End Sub



Private Sub txtChange_Click()
Dim f As String
Dim h As String
Dim p As String
Dim O As String

f = Val(Me.Text1.Text)
h = Val(Me.txtAmountPaid.Text)
p = Val(Me.txtBalance.Text)
O = Val(Me.txtChange.Text)
O = h - f
p = f - h
Me.txtChange.Text = "" & (O)
If h < f Then
Me.txtBalance.Text = "0"
ElseIf h > f Then
MsgBox "MORE MONEY PLEASE"
Me.txtAmountPaid.SetFocus
'Me.txtBalance.Text = "" & (p)
End If
End Sub

Private Sub txtPaymentof_Change()
Dim k As String
Dim g As String
Dim FOUND As Boolean
k = Me.txtPaymentof.Text
Dim db As Connection
Dim cntrl As Control
constr = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open constr

Set ordersRS = Nothing

db.Close
db.Open constr

Set ordersRS = New Recordset
ordersRS.Open "select * from orders", db, adOpenStatic, adLockOptimistic

If k = "Food and Bevarages Order" Then
g = InputBox("ENTER FOOD NO")
ordersRS.MoveFirst
While ordersRS.EOF = False And FOUND = False
If g = ordersRS!FOODNO Then
Me.Text1.Text = ordersRS!FoodCost
FOUND = True

End If
ordersRS.MoveNext
If ordersRS.EOF = True And FOUND = False Then
MsgBox "NOT A VALID FOOD NO"
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.CBOPAYMENTOF.Visible = True
Me.CBOPAYMENTOF.Enabled = True

ElseIf k = "Snacks" Then
g = InputBox("ENTER SNACK REF NO")
ordersRS.MoveFirst
While ordersRS.EOF = False And FOUND = False
ordersRS.MoveFirst
If g = ordersRS!FOODNO Then
Me.Text1.Text = ordersRS!FoodCost
FOUND = True
End If
ordersRS.MoveNext
If ordersRS.EOF = True And FOUND = False Then
MsgBox "QUICK FOOD NO"

Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False
Me.CBOPAYMENTOF.Enabled = True
Me.CBOPAYMENTOF.Visible = True

End If
Wend

End If
Wend

End If
End Sub
