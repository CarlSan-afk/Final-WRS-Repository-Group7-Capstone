VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_history 
   BorderStyle     =   0  'None
   Caption         =   "History"
   ClientHeight    =   9810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_history.frx":0000
   ScaleHeight     =   9810
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_print_sales 
      BackColor       =   &H00FFFF80&
      Caption         =   "Print "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9050
      Width           =   2535
   End
   Begin VB.OptionButton opt_all 
      BackColor       =   &H8000000D&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   5520
      TabIndex        =   8
      Top             =   480
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton opt_thismonth 
      BackColor       =   &H8000000D&
      Caption         =   "Passed 30 days"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10200
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton opt_thisweek 
      BackColor       =   &H8000000D&
      Caption         =   "Passed 7 days"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   8280
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.OptionButton opt_today 
      BackColor       =   &H8000000D&
      Caption         =   "Passed 24 hours"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc ado_history 
      Height          =   375
      Left            =   360
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "ado_history"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   5055
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
         ItemData        =   "frm_history.frx":DBEB
         Left            =   120
         List            =   "frm_history.frx":DBFB
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
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
         Left            =   2400
         TabIndex        =   2
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSComctlLib.ListView lstview_history 
      Height          =   7575
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13361
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer Name"
         Object.Width           =   3076
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "ID Number"
         Object.Width           =   3076
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Classification"
         Object.Width           =   3076
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Transaction Date"
         Object.Width           =   3076
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Transaction Details"
         Object.Width           =   13334
      EndProperty
   End
   Begin VB.Label lbl_history 
      Caption         =   "all"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbl_addedit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HISTORY LOG"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   12000
      TabIndex        =   4
      Top             =   480
      Width           =   2625
   End
End
Attribute VB_Name = "frm_history"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_search_Click()
Dim fromdate As Date, todate As Date
'fromdate = DateAdd("d" - 1, "")
Call Conn_ado_history
Select Case lbl_history.Caption
Case Is = "all"
    Select Case cbo_search.text
    Case Is = "Customer"
    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.text & "%' order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Classification"
    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.text & "%'order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Transaction Details"
    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.text & "%'order by transaction_date desc "
    ado_history.Refresh
    Case Is = "All"
    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
    ado_history.Refresh
    End Select
Case Is = "today"
    Select Case cbo_search.text
    Case Is = "Customer"
    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.text & "%' And transaction_date between  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Classification"
    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Transaction Details"
    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "All"
    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
    ado_history.Refresh
    End Select
Case Is = "thisweek"
    Select Case cbo_search.text
    Case Is = "Customer"
    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.text & "%' And transaction_date between  '" & Format(DateAdd("d", -7, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Classification"
    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -7, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Transaction Details"
    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -7, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "All"
    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
    ado_history.Refresh
    End Select
Case Is = "thismonth"
    Select Case cbo_search.text
    Case Is = "Customer"
    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.text & "%' And transaction_date between  '" & Format(DateAdd("d", -30, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Classification"
    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -30, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "' order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Transaction Details"
    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -30, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "All"
    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
    ado_history.Refresh
    End Select
End Select

'==================
    lstview_history.ListItems.Clear
    
    Do Until ado_history.Recordset.EOF
    Set list = lstview_history.ListItems.Add(, , ado_history.Recordset(0) & "")
    
    list.SubItems(1) = ado_history.Recordset(1) & ""
    list.SubItems(2) = ado_history.Recordset(2) & ""
    list.SubItems(3) = ado_history.Recordset(3) & ""
    list.SubItems(4) = ado_history.Recordset(4) & ""
    
    
        
    ado_history.Recordset.MoveNext
        Loop
    
    With lstview_history
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
    gvCmd1 = ado_history.RecordSource
End Sub

Private Sub cmd_print_sales_Click()
gvReportValue = 3
If gvReportValue = 3 Then
    Set frmPreview3 = New frmPreview
    frmPreview3.Show
End If
End Sub

Private Sub Form_Activate()
Call cbo_search_Click
End Sub

Private Sub Form_Load()
cbo_search.ListIndex = 1
Select Case mdi_wrs.WindowState
Case Is = vbNormal
mdi_wrs.Width = frm_history.Width + 315
mdi_wrs.Height = frm_history.Height + 955
DoCenter frm_history, mdi_wrs
Case Is = vbMaximized
DoCenter frm_history, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select

End Sub



Private Sub opt_all_Click()
lbl_history.Caption = "all"
Call cbo_search_Click
End Sub

Private Sub opt_thismonth_Click()
lbl_history.Caption = "thismonth"
Call cbo_search_Click
End Sub

Private Sub opt_thisweek_Click()
lbl_history.Caption = "thisweek"
Call cbo_search_Click
End Sub

Private Sub opt_today_Click()
lbl_history.Caption = "today"
Call cbo_search_Click
End Sub

Private Sub txt_search_Change()
Call Conn_ado_history
Select Case lbl_history.Caption
'Case Is = "all"
'    Select Case cbo_search.Text
'    Case Is = "Customer"
'    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.Text & "%'order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "Classification"
'    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.Text & "%'order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "Transaction Details"
'    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.Text & "%'order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "All"
'    ado_history.RecordSource = "select * from tbl_history "
'    ado_history.Refresh
'    End Select
'Case Is = "today"
'    Select Case cbo_search.Text
'    Case Is = "Customer"
'    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.Text & "%' And transaction_date between  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "Classification"
'    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.Text & "%'And transaction_date between # " & Now - 1 & " # and # " & Now & " # order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "Transaction Details"
'    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.Text & "%'And transaction_date between # " & Now - 1 & " # and # " & Now & " # order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "All"
'    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
'    ado_history.Refresh
'    End Select
'Case Is = "thisweek"
'    Select Case cbo_search.Text
'    Case Is = "Customer"
'    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.Text & "%' And transaction_date between # " & Now - 7 & " # and # " & Now & " # order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "Classification"
'    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.Text & "%'And transaction_date between # " & Now - 7 & " # and # " & Now & " # order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "Transaction Details"
'    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.Text & "%'And transaction_date between # " & Now - 7 & " # and # " & Now & " # order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "All"
'    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
'    ado_history.Refresh
'    End Select
'Case Is = "thismonth"
'    Select Case cbo_search.Text
'    Case Is = "Customer"
'    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.Text & "%' And transaction_date between # " & Now - 30 & " # and # " & Now & " # order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "Classification"
'    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.Text & "%'And transaction_date between # " & Now - 30 & " # and # " & Now & " # order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "Transaction Details"
'    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.Text & "%'And transaction_date between # " & Now - 30 & " # and # " & Now & " # order by transaction_date desc "
'    ado_history.Refresh
'    Case Is = "All"
'    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
'    ado_history.Refresh
'    End Select
'End Select


Case Is = "all"
    Select Case cbo_search.text
    Case Is = "Customer"
    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.text & "%' order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Classification"
    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.text & "%'order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Transaction Details"
    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.text & "%'order by transaction_date desc "
    ado_history.Refresh
    Case Is = "All"
    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
    ado_history.Refresh
    End Select
Case Is = "today"
    Select Case cbo_search.text
    Case Is = "Customer"
    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.text & "%' And transaction_date between  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Classification"
    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Transaction Details"
    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "All"
    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
    ado_history.Refresh
    End Select
Case Is = "thisweek"
    Select Case cbo_search.text
    Case Is = "Customer"
    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.text & "%' And transaction_date between  '" & Format(DateAdd("d", -7, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Classification"
    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -7, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Transaction Details"
    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -7, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "All"
    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
    ado_history.Refresh
    End Select
Case Is = "thismonth"
    Select Case cbo_search.text
    Case Is = "Customer"
    ado_history.RecordSource = "select * from tbl_history where customer_name like  '%" & txt_search.text & "%' And transaction_date between  '" & Format(DateAdd("d", -30, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Classification"
    ado_history.RecordSource = "select * from tbl_history where classification like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -30, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "' order by transaction_date desc "
    ado_history.Refresh
    Case Is = "Transaction Details"
    ado_history.RecordSource = "select * from tbl_history where transaction_details like  '%" & txt_search.text & "%'And transaction_date between  '" & Format(DateAdd("d", -30, Now), "yyyy-MM-dd") & "'  and  '" & Format(DateAdd("d", -1, Now), "yyyy-MM-dd") & "'  order by transaction_date desc "
    ado_history.Refresh
    Case Is = "All"
    ado_history.RecordSource = "select * from tbl_history order by transaction_date desc "
    ado_history.Refresh
    End Select
End Select
'==================
    lstview_history.ListItems.Clear
    
    Do Until ado_history.Recordset.EOF
    Set list = lstview_history.ListItems.Add(, , ado_history.Recordset(0) & "")
    
    list.SubItems(1) = ado_history.Recordset(1) & ""
    list.SubItems(2) = ado_history.Recordset(2) & ""
    list.SubItems(3) = ado_history.Recordset(3) & ""
    list.SubItems(4) = ado_history.Recordset(4) & ""
    
    
        
    ado_history.Recordset.MoveNext
        Loop
    
    With lstview_history
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
    gvCmd1 = ado_history.RecordSource
End Sub
