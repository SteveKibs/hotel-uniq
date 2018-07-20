VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmlist 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   5400
      Top             =   2880
   End
   Begin MSDataGridLib.DataGrid MENULIST 
      Bindings        =   "frmlist.frx":0000
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   14631
      _Version        =   393216
      BackColor       =   16777215
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
End
Attribute VB_Name = "frmlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Me.PrintForm
Unload Me
End Sub
