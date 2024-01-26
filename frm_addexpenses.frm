VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_addexpenses 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Expenses"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5340
   Icon            =   "frm_addexpenses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_addexpenses.frx":10CA
   ScaleHeight     =   3750
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   240
   End
   Begin MSAdodcLib.Adodc ado_expenses 
      Height          =   495
      Left            =   2640
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ado_expenses"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txt_expensesamount 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txt_itemtitle 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
   End
   Begin MSComctlLib.Toolbar tb_additem 
      Height          =   870
      Left            =   1080
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1535
      ButtonWidth     =   2831
      ButtonHeight    =   1429
      TextAlignment   =   1
      ImageList       =   "img_main"
      DisabledImageList=   "img_main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save       "
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel    "
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "frm_addexpenses.frx":5C60
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   4680
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addexpenses.frx":150BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addexpenses.frx":17694
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addexpenses.frx":19C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addexpenses.frx":1AD48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addexpenses.frx":1BE22
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addexpenses.frx":1E3FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADD EXPENSES"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   540
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   2610
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses Amount"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Title"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1020
   End
End
Attribute VB_Name = "frm_addexpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call SetWindowPos(frm_addexpenses.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
Call Conn_ado_expenses

End Sub


Private Sub Form_Terminate()
mdi_wrs.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdi_wrs.Enabled = True
End Sub

Private Sub tb_additem_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case Is = 1
'save
If txt_expensesamount.text = "" Then
Call SetWindowPos(frm_addexpenses.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
MsgBox " Input Amount", vbExclamation + vbOKOnly, "Water Refilling System"
Call SetWindowPos(frm_addexpenses.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
txt_expensesamount.SetFocus
Exit Sub
End If
Conn_ado_expenses
ado_expenses.Recordset.AddNew
ado_expenses.Recordset(0) = txt_itemtitle.text
ado_expenses.Recordset(1) = txt_expensesamount.text
ado_expenses.Recordset(2) = FormatDateTime(Now, vbShortDate)
ado_expenses.Recordset.Update
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = "Administrator"
        frm_history.ado_history.Recordset(1) = "9999"
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Expensed amount of  " & txt_expensesamount.text & " for " & txt_itemtitle.text
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
Call SetWindowPos(frm_addexpenses.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
MsgBox " Saved", vbInformation + vbOKOnly, "Water Refilling System"
Call SetWindowPos(frm_addexpenses.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
ado_expenses.Refresh
mdi_wrs.Enabled = True
        
Unload Me
Case Is = 2
'close
mdi_wrs.Enabled = True
Unload Me

End Select

End Sub

Private Sub Timer1_Timer()
For i = 1 To 4185
Me.Height = i
Me.Top = (Screen.Height \ 2) - (i \ 2)
Next
Timer1.Enabled = False

End Sub

Private Sub txt_expensesamount_Change()
On Error Resume Next
        If IsNumeric(txt_expensesamount.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_expensesamount.text = Trim(Left$(txt_expensesamount.text, Len(txt_expensesamount.text) - 1))
        End If
End Sub
