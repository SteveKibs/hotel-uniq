VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmkitchen 
   BackColor       =   &H00FF00FF&
   Caption         =   "                           KITCHEN DEPARTMENT"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid FoodOrders 
      Bindings        =   "frmkitchen.frx":0000
      Height          =   6975
      Left            =   360
      TabIndex        =   20
      Top             =   4200
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12303
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
      DataMember      =   "orders"
      ColumnCount     =   8
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
      BeginProperty Column02 
         DataField       =   "FoodName"
         Caption         =   "FoodName"
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
         DataField       =   "FoodCost"
         Caption         =   "FoodCost"
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
         DataField       =   "FoodNo"
         Caption         =   "FoodNo"
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
         DataField       =   "OrderNo"
         Caption         =   "OrderNo"
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
         DataField       =   "OrderTime"
         Caption         =   "OrderTime"
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
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Height          =   855
      Left            =   1320
      TabIndex        =   16
      Top             =   3360
      Width           =   5895
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Place Food Order"
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
         TabIndex        =   19
         Top             =   240
         Width           =   1935
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
         Height          =   420
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1815
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
         Height          =   420
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtReceiptNo 
      DataField       =   "ReceiptNo"
      Height          =   285
      Left            =   2910
      TabIndex        =   15
      Top             =   3015
      Width           =   3375
   End
   Begin VB.TextBox txtOrderTime 
      DataField       =   "OrderTime"
      Height          =   285
      Left            =   2910
      TabIndex        =   13
      Top             =   2640
      Width           =   1320
   End
   Begin VB.TextBox txtOrderNo 
      DataField       =   "OrderNo"
      Height          =   285
      Left            =   2910
      TabIndex        =   11
      Top             =   2265
      Width           =   1320
   End
   Begin VB.TextBox txtFoodNo 
      DataField       =   "FoodNo"
      Height          =   285
      Left            =   2910
      TabIndex        =   9
      Top             =   1875
      Width           =   1320
   End
   Begin VB.TextBox txtFoodCost 
      DataField       =   "FoodCost"
      Height          =   285
      Left            =   2910
      TabIndex        =   7
      Top             =   1500
      Width           =   1320
   End
   Begin VB.TextBox txtFoodName 
      DataField       =   "FoodName"
      Height          =   285
      Left            =   2910
      TabIndex        =   5
      Top             =   1125
      Width           =   3375
   End
   Begin VB.TextBox txtDate 
      DataField       =   "Date"
      Height          =   285
      Left            =   2910
      TabIndex        =   3
      Top             =   735
      Width           =   1320
   End
   Begin VB.TextBox txtCustomerNo 
      DataField       =   "CustomerNo"
      Height          =   285
      Left            =   2910
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReceiptNo:"
      Height          =   255
      Index           =   7
      Left            =   1065
      TabIndex        =   14
      Top             =   3060
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OrderTime:"
      Height          =   255
      Index           =   6
      Left            =   1065
      TabIndex        =   12
      Top             =   2685
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OrderNo:"
      Height          =   255
      Index           =   5
      Left            =   1065
      TabIndex        =   10
      Top             =   2310
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FoodNo:"
      Height          =   255
      Index           =   4
      Left            =   1065
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FoodCost:"
      Height          =   255
      Index           =   3
      Left            =   1065
      TabIndex        =   6
      Top             =   1545
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FoodName:"
      Height          =   255
      Index           =   2
      Left            =   1065
      TabIndex        =   4
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Index           =   1
      Left            =   1065
      TabIndex        =   2
      Top             =   780
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      Height          =   255
      Index           =   0
      Left            =   1065
      TabIndex        =   0
      Top             =   405
      Width           =   1815
   End
End
Attribute VB_Name = "frmkitchen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CONSTR As String
Dim WithEvents ordersRS As Recordset
Attribute ordersRS.VB_VarHelpID = -1
Dim WithEvents menuRS As Recordset
Attribute menuRS.VB_VarHelpID = -1
Dim WithEvents receptionRS As Recordset
Attribute receptionRS.VB_VarHelpID = -1


Private Sub cmdAdd_Click()
Me.cmdAdd.Enabled = False
Me.cmdUpdate.Enabled = True
Dim h As String
                    Dim FOUND As Boolean
                    Dim db As Connection
                    Dim cntrl As Control
                    CONSTR = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"
                    
                    Set db = New Connection
                    db.CursorLocation = adUseClient
                    db.Open CONSTR
                    
                    Set receptionRS = Nothing
                    
                    db.Close
                    db.Open CONSTR
                    
                    Set receptionRS = New Recordset
                    receptionRS.Open "select * from reception", db, adOpenStatic, adLockOptimistic
                    If receptionRS.BOF = True Then
                    MsgBox "THIS IS YOUR FIRST CUSTOMER"
                    frmReception.Show
                Else
                receptionRS.MoveFirst
                    
                    h = InputBox("Enter Customer No")
                     Do While receptionRS.EOF = False And FOUND = False
                    
                    If h = receptionRS!CustomerNo Then
                    Me.txtCustomerNo.Text = receptionRS!CustomerNo
                    Me.txtOrderTime.Text = "" & Time
                
Call GENNO
Call GAIN
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
End If
End Sub

Private Sub cmdClose_Click()
Dim t As String
                t = MsgBox("The System will end Kitchen Operations ,are you sure to end", vbYesNo)
                If t = vbYes Then
                Unload Me
                Else
                MsgBox "Sorry Operation Stopped by User", vbExclamation
End If
End Sub
Private Sub cmdUpdate_Click()
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False

                    If Me.txtCustomerNo.Text = "" Then
                    MsgBox "Enter Customer no", vbExclamation
                    Me.txtCustomerNo.SetFocus
                    End If
                    If Me.txtDate.Text = "" Then
                    MsgBox "Enter Date", vbExclamation
                    Me.txtDate.SetFocus
                    End If
                    If Me.txtFoodName.Text = "" Then
                    MsgBox "Enter FoodName", vbExclamation
                    Me.txtFoodName.SetFocus
                    End If
                    If Me.txtFoodCost.Text = "" Then
                    MsgBox "Enter FoodCost", vbExclamation
                    Me.txtFoodCost.SetFocus
                    End If
                    If Me.txtFoodNo.Text = "" Then
                    MsgBox "Enter FoodNo", vbExclamation
                    Me.txtFoodNo.SetFocus
                    End If
                    If Me.txtOrderNo.Text = "" Then
                    MsgBox "Enter OrderNo", vbExclamation
                    Me.txtOrderNo.SetFocus
                    End If
                    If Me.txtOrderTime.Text = "" Then
                    MsgBox "Enter OrderTime", vbExclamation
                    Me.txtOrderTime.SetFocus
                    End If
                    If Me.txtReceiptNo.Text = "" Then
                    MsgBox "Enter ReceiptNo", vbExclamation
                    Me.txtReceiptNo.SetFocus
                    Else
                    ordersRS.AddNew
                    ordersRS!CustomerNo = Me.txtCustomerNo.Text
                    ordersRS!Date = Me.txtDate.Text
                    ordersRS!FoodName = Me.txtFoodName.Text
                    ordersRS!FoodCost = Me.txtFoodCost.Text
                    ordersRS!FOODNO = Me.txtFoodNo.Text
                    ordersRS!OrderNo = Me.txtOrderNo.Text
                    ordersRS!OrderTime = Me.txtOrderTime.Text
                    ordersRS!ReceiptNo = Me.txtReceiptNo.Text
                    ordersRS.Update
                    MsgBox "Order placed,Goodbye!", vbExclamation
                    FRMORDER.Show
                    FRMORDER.Label1.Caption = Me.txtCustomerNo.Text
                    FRMORDER.Label2.Caption = Me.txtDate.Text
                    FRMORDER.Label3.Caption = Me.txtFoodName.Text
                    FRMORDER.Label4.Caption = Me.txtFoodCost.Text
                    FRMORDER.Label5.Caption = Me.txtFoodNo.Text
                    FRMORDER.Label6.Caption = Me.txtOrderNo.Text
                    FRMORDER.Label7.Caption = Me.txtOrderTime.Text
                    FRMORDER.Label8.Caption = Me.txtReceiptNo.Text
                    Call clear
                    Me.cmdAdd.Enabled = True
                    Me.cmdUpdate.Enabled = False
                    
                    End If

End Sub
Private Sub clear()
Me.txtCustomerNo.Text = ""
Me.txtDate.Text = ""
Me.txtFoodName.Text = ""
Me.txtFoodCost.Text = ""
Me.txtFoodNo.Text = ""
Me.txtOrderNo.Text = ""
Me.txtOrderTime.Text = ""
Me.txtReceiptNo.Text = ""
End Sub

Private Sub Form_Load()
Dim db As Connection
Dim cntrl As Control
CONSTR = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open CONSTR

Set ordersRS = Nothing

db.Close
db.Open CONSTR

Set ordersRS = New Recordset
ordersRS.Open "select * from orders", db, adOpenStatic, adLockOptimistic
Me.cmdAdd.Enabled = True
Me.cmdUpdate.Enabled = False


End Sub

Private Sub GENNO()
Dim O As String
                    Dim db As Connection
                    Dim cntrl As Control
                    CONSTR = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"
                    
                    Set db = New Connection
                    db.CursorLocation = adUseClient
                    db.Open CONSTR
                    
                    Set menuRS = Nothing
                    
                    db.Close
                    db.Open CONSTR
                    
                    Set menuRS = New Recordset
                    menuRS.Open "select * from menu", db, adOpenStatic, adLockOptimistic
                    Dim FOUND As Boolean
                    O = InputBox("Enter Food No")
                    
                    menuRS.MoveFirst
                    Do While menuRS.EOF = False And FOUND = False
                    If O = menuRS!FOODNO Then
                    Me.txtFoodNo.Text = menuRS!FOODNO
                    Me.txtDate.Text = Date & ""
                    Me.txtFoodCost.Text = menuRS!FoodCost
                    Me.txtFoodName.Text = menuRS!FoodName
                    Me.txtOrderTime.Text = Time & ""
                    FOUND = True
                    Exit Sub
                    End If
                    menuRS.MoveNext
                    Loop
                    If menuRS.EOF = True And FOUND = False Then
                    MsgBox "NOT A VALID FOOD NO"
                    Me.cmdAdd.Enabled = True
                    Me.cmdUpdate.Enabled = False
                    Exit Sub
                    End If


End Sub

Private Sub GAIN()
Dim db As Connection
Dim cntrl As Control
CONSTR = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open CONSTR

Set ordersRS = Nothing

db.Close
db.Open CONSTR

Set ordersRS = New Recordset
ordersRS.Open "select * from orders", db, adOpenStatic, adLockOptimistic
If ordersRS.BOF = True Then
Me.txtReceiptNo.Text = "JBHON201301"
Else
ordersRS.MoveFirst
ordersRS.MoveLast
Me.txtReceiptNo.Text = ordersRS.RecordCount + 1
Me.txtOrderNo.Text = ordersRS!OrderNo + 1
End If
End Sub
