VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_borrow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Borrow Gallon"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5340
   Icon            =   "frm_borrow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_borrow.frx":10CA
   ScaleHeight     =   3750
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbo_borrow 
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
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox cbo_itemlist 
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
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin MSComctlLib.Toolbar tb_additem 
      Height          =   870
      Left            =   1680
      TabIndex        =   4
      Top             =   2760
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1535
      ButtonWidth     =   3069
      ButtonHeight    =   1429
      TextAlignment   =   1
      ImageList       =   "img_main"
      DisabledImageList=   "img_main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Borrow     "
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel       "
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "frm_borrow.frx":5C60
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   3840
      Top             =   240
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
            Picture         =   "frm_borrow.frx":150BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_borrow.frx":17694
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_borrow.frx":19C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_borrow.frx":1AD48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_borrow.frx":1BE22
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_borrow.frx":1E3FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado_customeritem 
      Height          =   330
      Left            =   2760
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      ConnectStringType=   2
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
      Caption         =   "ado_customeritem"
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
   Begin MSAdodcLib.Adodc ado_itemlist 
      Height          =   330
      Left            =   2760
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      ConnectStringType=   2
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
      Caption         =   "ado_itemlist"
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
   Begin VB.Label lbl_customer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guest"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lbl_idnumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Top             =   1680
      Width           =   945
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
      TabIndex        =   0
      Top             =   600
      Width           =   1020
   End
End
Attribute VB_Name = "frm_borrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_itemlist_Click()
On Error Resume Next
Dim i As Integer
Call Conn_ado_itemlistBR
ado_itemlist.RecordSource = "select * from tbl_item_list where Item_title like  '%" & cbo_itemlist.text & "%'"
ado_itemlist.Refresh
cbo_borrow.Clear

For i = 1 To ado_itemlist.Recordset(13)
cbo_borrow.AddItem i
Next
cbo_borrow.ListIndex = 0



End Sub

Private Sub Form_Load()
On Error Resume Next
Call Conn_ado_itemlistBR
ado_itemlist.RecordSource = "select * from tbl_item_list where pos_item ='Yes'"
ado_itemlist.Refresh
cbo_itemlist.Clear

Do Until ado_itemlist.Recordset.EOF
cbo_itemlist.AddItem ado_itemlist.Recordset(0)
ado_itemlist.Recordset.MoveNext
Loop
cbo_itemlist.ListIndex = 0
Call SetWindowPos(frm_borrow.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub

Private Sub Form_Terminate()
mdi_wrs.Enabled = True
frm_CI.ado_customeritem.Refresh
Call CustomerItemCI_ListLoad
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdi_wrs.Enabled = True
frm_CI.ado_customeritem.Refresh
Call CustomerItemCI_ListLoad
Unload Me
End Sub

Private Sub tb_additem_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case Is = 1
If Val(cbo_borrow.text) <= 0 Then
Call SetWindowPos(frm_borrow.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
 MsgBox "Nothing to Borrow", vbExclamation, "WRS Management System"
 Call SetWindowPos(frm_borrow.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
Exit Sub
End If

Call Conn_ado_customeritemBR
ado_customeritem.RecordSource = "select * from tbl_customer_item  where ID_number like  '%" & lbl_idnumber.Caption & "%' and  Item_title like  '%" & cbo_itemlist.text & "%'"
ado_customeritem.Refresh
    If ado_customeritem.Recordset.RecordCount = 0 Then
    ado_customeritem.Recordset.AddNew
    ado_customeritem.Recordset(0) = lbl_idnumber.Caption
    ado_customeritem.Recordset(1) = lbl_customer.Caption
    ado_customeritem.Recordset(2) = cbo_itemlist.text
    ado_customeritem.Recordset(3) = cbo_borrow.text
    ado_customeritem.Recordset(4) = FormatDateTime(Now, vbShortDate)
    ado_customeritem.Recordset.Update
    ado_customeritem.Refresh
    
    ado_itemlist.Recordset(13) = Val(ado_itemlist.Recordset(13)) - Val(cbo_borrow.text)
    ado_itemlist.Recordset(14) = Val(ado_itemlist.Recordset(14)) + Val(cbo_borrow.text)
    ado_itemlist.Recordset.Update
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = lbl_customer.Caption
        frm_history.ado_history.Recordset(1) = lbl_idnumber.Caption
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Borrowed  (" & cbo_borrow.text & ") " & cbo_itemlist.text
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
    frm_CI.ado_customeritem.Refresh
    Call CustomerItemCI_ListLoad
    Call SetWindowPos(frm_borrow.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    MsgBox "Success", vbInformation, "WRS Management System"
    Call SetWindowPos(frm_borrow.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
        
    cbo_itemlist.ListIndex = 0
    mdi_wrs.Enabled = True
    frm_CI.ado_customeritem.Refresh
    Call CustomerItemCI_ListLoad
    Call CustomerItem_ListLoad
    Unload Me
    Else
    ado_customeritem.Recordset(0) = lbl_idnumber.Caption
    ado_customeritem.Recordset(1) = lbl_customer.Caption
    ado_customeritem.Recordset(2) = cbo_itemlist.text
    ado_customeritem.Recordset(3) = Val(cbo_borrow.text) + Val(ado_customeritem.Recordset(3))
    ado_customeritem.Recordset(4) = FormatDateTime(Now, vbShortDate)
    ado_customeritem.Recordset.Update
    ado_customeritem.Refresh
    
    ado_itemlist.Recordset(13) = Val(ado_itemlist.Recordset(13)) - Val(cbo_borrow.text)
    ado_itemlist.Recordset(14) = Val(ado_itemlist.Recordset(14)) + Val(cbo_borrow.text)
    ado_itemlist.Recordset.Update
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = lbl_customer.Caption
        frm_history.ado_history.Recordset(1) = lbl_idnumber.Caption
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Borrowed  (" & cbo_borrow.text & ") " & cbo_itemlist.text
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
    frm_CI.ado_customeritem.Refresh
    Call CustomerItemCI_ListLoad
    Call CustomerItem_ListLoad
    Call SetWindowPos(frm_borrow.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    MsgBox "Success", vbInformation, "WRS Management System"
    Call SetWindowPos(frm_borrow.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
        
    cbo_itemlist.ListIndex = 0
    mdi_wrs.Enabled = True
    frm_CI.ado_customeritem.Refresh
    Call CustomerItemCI_ListLoad
    Call CustomerItem_ListLoad
    Unload Me
    End If

Case Is = 2
mdi_wrs.Enabled = True
frm_CI.ado_customeritem.Refresh
Call CustomerItemCI_ListLoad
Unload Me

End Select
End Sub
