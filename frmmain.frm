VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   "                                                                  JUBELEE BEACH HOTEL MANAGEMENT SYSTEM"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   15240
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmmain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   800
      ImageHeight     =   600
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":29D6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1164
      ButtonWidth     =   2566
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "RECEPTION"
            Key             =   "RECEPTION"
            Object.ToolTipText     =   "Allows admission of new customers to the hotel"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RECEPTION"
                  Object.Tag             =   "RECEPTION"
                  Text            =   "RECEPTION"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "FINANCE OFFICE"
            Key             =   "FINANCE OFFICE"
            Object.ToolTipText     =   "All Payments are Done Here"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "FINANCE OFFICE"
                  Object.Tag             =   "FINANCE OFFICE"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "FOOD ORDERS"
            Key             =   "FOOD ORDERS"
            Object.ToolTipText     =   "All Food and Drinks are Ordered from here!"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "FOOD ORDERS"
                  Object.Tag             =   "FOOD ORDERS"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "HOTEL MENU"
            Key             =   "HOTEL MENU"
            Object.ToolTipText     =   "all foods and drinks and their prices are indicated here"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HOTEL MENU"
                  Object.Tag             =   "HOTEL MENU"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu Workers 
         Caption         =   "WORKERS REGISTER"
      End
      Begin VB.Menu Allowances 
         Caption         =   "ADD ALLOWANCES"
      End
      Begin VB.Menu Deductions 
         Caption         =   "ADD DEDUCTIONS"
      End
      Begin VB.Menu Payslip 
         Caption         =   "WORKER'S PAYSLIP"
      End
   End
   Begin VB.Menu VIEW 
      Caption         =   "View"
      Begin VB.Menu reports 
         Caption         =   "Reports"
         Begin VB.Menu RECEP 
            Caption         =   "RECEPTION OFFICE"
         End
         Begin VB.Menu ACC 
            Caption         =   "ACCOUNTS OFFICE"
         End
         Begin VB.Menu WORK 
            Caption         =   "WORKERS REGISTER"
         End
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ACC_Click()
DPTACCOUNTS.Show
End Sub

Private Sub ACCO_Click()
DPTACCOMODATION.Show
End Sub

Private Sub ACCOUNTSQ_Click()
frmAccounts.Show
End Sub

Private Sub Allowances_Click()
FrmAllowances.Show
End Sub

Private Sub Command10_Click()
FrmPayslip.Show
End Sub

Private Sub Command11_Click()
Frmbooking.Show
End Sub

Private Sub Command7_Click()
frmTourSites.Show
End Sub

Private Sub Command8_Click()
frmaccomodation.Show
End Sub

Private Sub date_Click()
'Shell ("C:\WINDOWS\system32\control.exe date/time")
'Shell ("C:\WINDOWS\system32\calc")
'Shell ("C:\WINDOWS\system32\freecell")
'Shell ("C:\WINDOWS\system32\mplay32")
'Shell ("C:\WINDOWS\system32\mshearts")
'Shell ("C:\WINDOWS\system32\mspaint")
'Shell ("C:\WINDOWS\system32\narrator")
Shell ("C:\WINDOWS\system32\odbcad32")
'Shell ("C:\WINDOWS\system32\sndvol32")
'Shell ("C:\WINDOWS\system32\sndrec32")
'Shell ("C:\WINDOWS\system32\spider")
'Shell ("C:\WINDOWS\system32\telnet")
'Shell ("C:\WINDOWS\system32\wiaacmgr")
Shell ("C:\WINDOWS\system32\wpabaln")

End Sub

Private Sub Deductions_Click()
FrmDeductions.Show
End Sub

Private Sub Houses_Click()
frmaccomodation.Show
End Sub

Private Sub Kitchmenu_Click()

End Sub

Private Sub KITCHEN_Click()
frmkitchen.Show
End Sub

Private Sub mnuKitchmenu_Click()
frmMenu.Show
End Sub

Private Sub NewHouses_Click()
Frmhouses.Show
End Sub

Private Sub Payslip_Click()
FrmPayslip.Show
End Sub

Private Sub RECEP_Click()
DPTRECEPTION.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "FINANCE OFFICE"
frmAccounts.Show
Case "HOTEL MENU"
frmMenu.Show
Case "FOOD ORDERS"
frmkitchen.Show
Case "RECEPTION"
frmReception.Show
Case "ASSIGN ROOM"
FRMASSIGN.Show
End Select

End Sub

Private Sub tours_Click()
frmTourSites.Show
End Sub

Private Sub TRANS_Click()
dpttransport.Show
End Sub

Private Sub Trips_Click()
Frmtransport.Show
End Sub

Private Sub WORK_Click()
DPTWORKERS.Show
End Sub

Private Sub Workers_Click()
FrmWorkers.Show
End Sub
