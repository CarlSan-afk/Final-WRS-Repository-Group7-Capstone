VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_addstocks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Stocks"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5625
   Icon            =   "frm_addstocks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_addstocks.frx":10CA
   ScaleHeight     =   3240
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   240
   End
   Begin VB.CheckBox chk_remove 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Remove / Damage / Lost"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txt_unitcost 
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
      Left            =   600
      TabIndex        =   4
      Text            =   "0"
      Top             =   2520
      Width           =   1455
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
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox txt_addstocks 
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
      Left            =   600
      TabIndex        =   0
      Text            =   "0"
      Top             =   1680
      Width           =   1455
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   2040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addstocks.frx":5C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addstocks.frx":823A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addstocks.frx":A814
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addstocks.frx":B8EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addstocks.frx":C9C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addstocks.frx":EFA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_addstocks.frx":1007C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_additem 
      Height          =   870
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
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
            Caption         =   "Save      "
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel    "
            ImageIndex      =   7
         EndProperty
      EndProperty
      MouseIcon       =   "frm_addstocks.frx":12656
   End
   Begin MSAdodcLib.Adodc ado_addstocks 
      Height          =   330
      Left            =   2760
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
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
      Caption         =   "ado_addstocks"
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
   Begin MSAdodcLib.Adodc ado_expenses 
      Height          =   375
      Left            =   2760
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
   Begin VB.Label lbl_unitcost 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Unit Cost"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   990
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
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label lbl_addremove 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Stocks"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1230
   End
End
Attribute VB_Name = "frm_addstocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chk_remove_Click()
If chk_remove.Value = 1 Then
lbl_addremove.Caption = "Remove Stocks"
lbl_unitcost.Visible = False
txt_unitcost.Visible = False
Else
lbl_addremove.Caption = "Add Stocks"
lbl_unitcost.Visible = True
txt_unitcost.Visible = True
End If

End Sub

Private Sub Form_Activate()
Call SetWindowPos(frm_addstocks.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
Call Conn_ado_addstocks
ado_addstocks.RecordSource = "select * from tbl_item_list where Item_title like  '%" & txt_itemtitle.text & "%'"
ado_addstocks.Refresh
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
If txt_addstocks.text = "" Or txt_unitcost.text = "" Then
    Call SetWindowPos(frm_addstocks.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    MsgBox "Input stocks and Unit Cost", vbExclamation + vbOKOnly, "Water Refilling System"
    Call SetWindowPos(frm_addstocks.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Exit Sub
End If
If chk_remove.Value = 1 Then
    If Val(txt_addstocks.text) < (Val(ado_addstocks.Recordset(13)) + 1) Then
    ado_addstocks.Recordset(13) = Val(ado_addstocks.Recordset(13)) - Val(txt_addstocks.text)
    ado_addstocks.Recordset(15) = Val(txt_addstocks.text) + Val(ado_addstocks.Recordset(15))
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = "Administrator"
        frm_history.ado_history.Recordset(1) = "9999"
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Removed/Damage/Lost  (" & txt_addstocks.text & ")pcs " & txt_itemtitle.text & " to Itemlist"
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
    Else
    Call SetWindowPos(frm_addstocks.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    MsgBox "Inputed number is Greater than Stocks", vbExclamation + vbOKOnly, "Water Refilling System"
    Call SetWindowPos(frm_addstocks.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Exit Sub
    End If
Else
ado_addstocks.Recordset(13) = Val(txt_addstocks.text) + Val(ado_addstocks.Recordset(13))
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = "Administrator"
        frm_history.ado_history.Recordset(1) = "9999"
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Added (" & txt_addstocks & ") Stocks of " & txt_itemtitle.text & " to Itemlist"
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================

End If
ado_addstocks.Recordset.Update
ado_addstocks.Refresh

ado_expenses.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
ado_expenses.RecordSource = "select * from tbl_expenses"
ado_expenses.Refresh
ado_expenses.Recordset.AddNew
ado_expenses.Recordset(0) = txt_itemtitle.text

If chk_remove.Value = 1 Then
ado_expenses.Recordset(1) = "0"
Else
ado_expenses.Recordset(1) = Val(txt_addstocks.text) * Val(txt_unitcost.text)
End If

ado_expenses.Recordset(2) = FormatDateTime(Now, vbShortDate)
ado_expenses.Recordset.Update
ado_expenses.Refresh
Call SetWindowPos(frm_addstocks.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
MsgBox "Saved", vbInformation + vbOKOnly, "Water Refilling System"
Call SetWindowPos(frm_addstocks.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
mdi_wrs.Enabled = True
Call frm_itemlist.lstview_itemlist_Click
Unload Me

Case Is = 2
'close
mdi_wrs.Enabled = True
Unload Me
End Select

End Sub

Private Sub Timer1_Timer()
For i = 1 To 3750
Me.Height = i
Me.Top = (Screen.Height \ 2) - (i \ 2)
Next
Timer1.Enabled = False

End Sub

Private Sub txt_addstocks_Change()
On Error Resume Next
        If IsNumeric(txt_addstocks.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_addstocks.text = Trim(Left$(txt_addstocks.text, Len(txt_addstocks.text) - 1))
        End If
End Sub

Private Sub txt_unitcost_Change()
On Error Resume Next
        If IsNumeric(txt_unitcost.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_unitcost.text = Trim(Left$(txt_unitcost.text, Len(txt_unitcost.text) - 1))
        End If
End Sub
