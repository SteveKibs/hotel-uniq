VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                      KITCHEN MENU LIST"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11145
   ScaleWidth      =   13665
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   2895
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   8535
      Begin VB.TextBox txtFoodNo 
         DataField       =   "FoodNo"
         Height          =   285
         Left            =   3075
         TabIndex        =   14
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox txtFoodName 
         DataField       =   "FoodName"
         Height          =   285
         Left            =   3075
         TabIndex        =   13
         Top             =   495
         Width           =   3375
      End
      Begin VB.TextBox txtFoodCost 
         DataField       =   "FoodCost"
         Height          =   285
         Left            =   3075
         TabIndex        =   12
         Top             =   885
         Width           =   1320
      End
      Begin VB.TextBox txtDateAdded 
         DataField       =   "DateAdded"
         Height          =   285
         Left            =   3075
         TabIndex        =   11
         Top             =   1260
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000FF00&
         Height          =   1215
         Left            =   0
         TabIndex        =   6
         Top             =   1560
         Width           =   4455
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H0000FFFF&
            Caption         =   "&Remove Menu"
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
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   2055
         End
         Begin VB.CommandButton cmdadd 
            BackColor       =   &H0000FFFF&
            Caption         =   "&Add Food to Menu"
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
            TabIndex        =   9
            Top             =   120
            Width           =   2055
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
            Height          =   405
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   600
            Width           =   2055
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
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   600
            Width           =   2055
         End
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   1935
         Left            =   4440
         TabIndex        =   5
         Top             =   840
         Width           =   3975
         _Version        =   524288
         _ExtentX        =   7011
         _ExtentY        =   3413
         _StockProps     =   1
         BackColor       =   33023
         Year            =   2013
         Month           =   2
         Day             =   22
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FoodNo:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2175
         TabIndex        =   1
         Top             =   165
         Width           =   870
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FoodName:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1875
         TabIndex        =   17
         Top             =   540
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FoodCost:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2010
         TabIndex        =   16
         Top             =   930
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DateAdded:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1830
         TabIndex        =   15
         Top             =   1305
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Print New Menu"
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid MENULIST 
      Bindings        =   "frmMenus.frx":0000
      Height          =   7815
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   13785
      _Version        =   393216
      BackColor       =   65280
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
      DataMember      =   "menu"
      ColumnCount     =   3
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   840.189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JUBELEE BEACH HOTELS FOOD AND DRINKS MENU"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   13095
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public constr As String
Dim WithEvents menuRS As Recordset
Attribute menuRS.VB_VarHelpID = -1

Private Sub Calendar1_Click()
Me.txtDateAdded.Text = Calendar1.Value
End Sub

Private Sub cmdClose_Click()
Dim t As String
t = MsgBox("The System will end Adding Food to Menu Operations,are you sure to end", vbYesNo)
If t = vbYes Then
Unload Me
Else
MsgBox "Sorry Operation Stopped by User", vbCritical
End If
End Sub

Private Sub cmdDelete_Click()
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
Dim h As String
Dim FOUND As Boolean
Dim Y As String
h = InputBox("Enter Food No to Remove from Menu")

While menuRS.EOF = False And FOUND = False
menuRS.MoveFirst
If h = menuRS!FOODNO Then
Y = MsgBox("Are Sure to Delete this Food from Menu", vbYesNo)
If Y = vbYes Then
menuRS.Delete
MsgBox "Deleted"
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
Exit Sub
Else
MsgBox "Operation Stopped by User"
Exit Sub
End If
menuRS.MoveNext
If menuRS.EOF = True And FOUND = False Then
MsgBox "This Menu item is not Available"
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
End If
End If
Wend
End With
End Sub

Private Sub cmdAdd_Click()
Me.cmdUpdate.Enabled = True
Me.cmdDelete.Enabled = False
If menuRS.BOF = True Then
Me.txtFoodNo.Text = "1001"
Else
menuRS.MoveFirst
menuRS.MoveLast
Me.txtFoodNo.Text = menuRS!FOODNO + 2
End If
End Sub

Private Sub cmdUpdate_Click()

Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = True

If Me.txtFoodNo.Text = "" Then
MsgBox "Enter Food No", vbCritical
Me.txtFoodNo.SetFocus
End If

If Me.txtFoodName.Text = "" Then
MsgBox "Enter Food Name", vbCritical
Me.txtFoodName.SetFocus
End If

If Me.txtFoodCost.Text = "" Then
MsgBox "Enter Food Cost", vbCritical
Me.txtFoodCost.SetFocus
End If

If Me.txtDateAdded.Text = "" Then
MsgBox "Enter Date Added", vbCritical
Me.txtDateAdded.SetFocus
Else
menuRS.AddNew
menuRS!FOODNO = Me.txtFoodNo.Text
menuRS!FoodName = Me.txtFoodName.Text
menuRS!FoodCost = Me.txtFoodCost.Text
menuRS!DateAdded = Me.txtDateAdded.Text
menuRS.Update
menuRS.Requery
MsgBox "Food Added to Menu successfully", vbCritical
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = True
Call clear
End If

End Sub
Private Sub clear()
Me.txtFoodNo.Text = ""
Me.txtFoodName.Text = ""
Me.txtFoodCost.Text = ""
Me.txtDateAdded.Text = ""
End Sub

Private Sub Command1_Click()
Unload Me
frmlist.Show

End Sub

Private Sub Form_Load()
Dim db As Connection
Dim cntrl As Control
constr = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver=MySQL ODBC 5.1 Driver;SERVER=localhost;UID=root;DATABASE=hotels;PORT=3306"

Set db = New Connection
db.CursorLocation = adUseClient
db.Open constr

Set menuRS = Nothing

db.Close
db.Open constr

Set menuRS = New Recordset
menuRS.Open "select * from menu", db, adOpenStatic, adLockOptimistic

Me.cmdUpdate.Enabled = False
'Me.cmdRefresh.Enabled = True
Me.cmdDelete.Enabled = True


End Sub
