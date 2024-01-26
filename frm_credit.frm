VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_credit 
   BorderStyle     =   0  'None
   Caption         =   "Credit"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_credit.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_search 
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
      Left            =   5640
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox cbo_search 
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
      ItemData        =   "frm_credit.frx":DBEB
      Left            =   3240
      List            =   "frm_credit.frx":DBF5
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8040
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc ado_credit 
      Height          =   375
      Left            =   360
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "ado_credit"
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
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   14535
      Begin MSComctlLib.ListView lstview_credit 
         Height          =   5775
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12632319
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Balance"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Name"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Classification"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID Number"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Credited"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Delivered By"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Item Title"
            Object.Width           =   6410
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar_main 
      Height          =   870
      Left            =   11280
      TabIndex        =   1
      Top             =   8040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1535
      ButtonWidth     =   2831
      ButtonHeight    =   1429
      TextAlignment   =   1
      ImageList       =   "img_main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pay       "
            ImageIndex      =   4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print       "
            ImageIndex      =   5
         EndProperty
      EndProperty
      MouseIcon       =   "frm_credit.frx":DC18
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   7920
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483646
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_credit.frx":1D072
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_credit.frx":1F64C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_credit.frx":21C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_credit.frx":22D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_credit.frx":252DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Credit: "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   7680
      Width           =   1590
   End
   Begin VB.Label lbl_total_credit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   7680
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CREDIT LOG"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   2475
   End
End
Attribute VB_Name = "frm_credit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_search_Change()
    Select Case cbo_search.text
        Case Is = "Customer Name"
            frm_credit.ado_credit.RecordSource = "select * from tbl_credit where customer_name like '%" & txt_search.text & "%' "
            frm_credit.ado_credit.Refresh
            gvCmdCredit = frm_credit.ado_credit.RecordSource
            Call Credit_ListLoadS
        Case Is = "Classification"
            frm_credit.ado_credit.RecordSource = "select * from tbl_credit where classification like '%" & txt_search.text & "%' "
            frm_credit.ado_credit.Refresh
            gvCmdCredit = frm_credit.ado_credit.RecordSource
            Call Credit_ListLoadS
        Case Else
            Call Credit_ListLoad
    End Select
End Sub

Private Sub Form_Activate()

Call Credit_ListLoad
Call Creditmath

End Sub

Private Sub Form_Load()
Call Conn_ado_credit
Call Credit_ListLoad
'lbl_idnumber.Caption = lstview_credit.SelectedItem.SubItems(3)

Select Case mdi_wrs.WindowState
Case Is = vbNormal
mdi_wrs.Width = frm_credit.Width + 315
mdi_wrs.Height = frm_credit.Height + 955
DoCenter frm_credit, mdi_wrs
Case Is = vbMaximized
DoCenter frm_credit, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select

End Sub



Private Sub Toolbar_main_ButtonClick(ByVal Button As MSComctlLib.Button)
' Pay Credit
Select Case Button.Index
Case Is = 1
    If lstview_credit.ListItems.Count = 0 Then Exit Sub
    frm_payment.Show
    frm_payment.txt_amounttopay.text = Val(lstview_credit.SelectedItem) * -1
    mdi_wrs.Enabled = False
    frm_payment.Caption = "Pay Credit"
    Call Conn_ado_payment
    frm_payment.ado_payment.RecordSource = "select * from tbl_credit where id_number like  '%" & lstview_credit.SelectedItem.SubItems(3) & "%'"
    frm_payment.ado_payment.Refresh
    With frm_payment
    .lbl_amount.Caption = .ado_payment.Recordset(1)
    .lbl_customer_name.Caption = .ado_payment.Recordset(2)
    .lbl_classification.Caption = .ado_payment.Recordset(3)
    .lbl_id_number.Caption = .ado_payment.Recordset(4)
    .lbl_deliveryman.Caption = .ado_payment.Recordset(6)
    .lbl_sold_item.Caption = .ado_payment.Recordset(7)
    End With
    Call Creditmath
    
'    Call Conn_ado_history
'    frm_history.ado_history.Recordset.AddNew
'    frm_history.ado_history.Recordset(0) = ado_delivery.Recordset(2)
'    frm_history.ado_history.Recordset(1) = ado_delivery.Recordset(4)
'    frm_history.ado_history.Recordset(2) = ado_delivery.Recordset(3)
'    frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
'    frm_history.ado_history.Recordset(4) = ado_delivery.Recordset(7) & "/Pay " & ado_delivery.Recordset(1)
'    frm_history.ado_history.Recordset.Update
'    frm_history.ado_history.Refresh

Case Is = 2
''print
'Call Conn_ado_credit
'
' With CrystalReport1
'.ReportFileName = App.path & "\credit.rpt"
'.RetrieveDataFiles
'.WindowState = crptMaximized
'.Destination = crptToWindow
'.Action = 1
'

'End With

    gvReportValue = 6
    If gvReportValue = 6 Then
        Set frmPreview6 = New frmPreview
        frmPreview6.Show
    End If
End Select


End Sub

Private Sub txt_search_Change()
    Select Case cbo_search.text
        Case Is = "Customer Name"
            frm_credit.ado_credit.RecordSource = "select * from tbl_credit where customer_name like '%" & txt_search.text & "%' "
            frm_credit.ado_credit.Refresh
            gvCmdCredit = frm_credit.ado_credit.RecordSource
            Call Credit_ListLoadS
        Case Is = "Classification"
            frm_credit.ado_credit.RecordSource = "select * from tbl_credit where classification like '%" & txt_search.text & "%' "
            frm_credit.ado_credit.Refresh
            gvCmdCredit = frm_credit.ado_credit.RecordSource
            Call Credit_ListLoadS
        Case Else
            Call Credit_ListLoad
    End Select
End Sub
