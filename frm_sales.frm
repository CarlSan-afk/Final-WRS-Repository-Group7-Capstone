VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_sales 
   BorderStyle     =   0  'None
   Caption         =   "Sales"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_sales.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmd_delete 
      BackColor       =   &H00FFFF80&
      Caption         =   "Delete"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8400
      Width           =   2535
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   14640
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc ado_sales 
      Height          =   375
      Left            =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "ado_sales"
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
   Begin VB.CommandButton cmd_print_sales 
      BackColor       =   &H00FFFF80&
      Caption         =   "Print Sales Report"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8400
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   9720
      TabIndex        =   8
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   130023425
      CurrentDate     =   44499
   End
   Begin VB.CommandButton cmd_thismonth 
      BackColor       =   &H00FFFF00&
      Caption         =   "This Month"
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmd_thisweek 
      BackColor       =   &H00FFFF00&
      Caption         =   "This Week"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmd_yesterday 
      BackColor       =   &H00FFFF80&
      Caption         =   "Yesterday"
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmd_today 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   8295
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
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Customer"
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
         TabIndex        =   21
         Top             =   120
         Width           =   1905
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   12360
      TabIndex        =   9
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   130023425
      CurrentDate     =   44499
   End
   Begin MSAdodcLib.Adodc ado_expenses 
      Height          =   375
      Left            =   2280
      Top             =   9120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.Frame fm_sales 
      Caption         =   "SALES LOG"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   14775
      Begin MSComctlLib.ListView lstview_sales 
         Height          =   5415
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Amount"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Name"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Classification"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "ID Number"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Date Sold"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Delivered By"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Item Title"
            Object.Width           =   6410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "TN"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fm_expenses 
      Caption         =   "EXPENSES LOG"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   14775
      Begin MSComctlLib.ListView lstview_expenses 
         Height          =   5415
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title of Expenses"
            Object.Width           =   8546
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Cost of Expenses"
            Object.Width           =   8546
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Incurred Date"
            Object.Width           =   8546
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "TN"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALES AND EXPENSES"
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
      Left            =   2280
      TabIndex        =   22
      Top             =   480
      Width           =   4515
   End
   Begin VB.Label lbl_profit 
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
      Left            =   7440
      TabIndex        =   20
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profit : "
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
      Left            =   6480
      TabIndex        =   19
      Top             =   8520
      Width           =   885
   End
   Begin VB.Label lbl_expenses 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   4680
      TabIndex        =   18
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses : "
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
      Left            =   3360
      TabIndex        =   17
      Top             =   8520
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Revenue : "
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
      TabIndex        =   16
      Top             =   8520
      Width           =   1785
   End
   Begin VB.Label lbl_sales 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2160
      TabIndex        =   13
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   12360
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Left            =   9720
      TabIndex        =   10
      Top             =   360
      Width           =   1170
   End
End
Attribute VB_Name = "frm_sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_delete_Click()

'On Error GoTo err

    

    If fm_sales.Visible = True Then
        If lstview_sales.ListItems.Count = 0 Then Exit Sub
        answer = MsgBox("You are about to DELETE record, Do you want to continue?", vbExclamation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        
        Call Conn_ado_sales
        ado_sales.RecordSource = "select * from tbl_sales where TN = " & lstview_sales.SelectedItem.SubItems(7)
        ado_sales.Refresh
'        If Left(ado_sales.Recordset(6), 3) = "Buy" Then
'
'        End If
        ado_sales.Recordset.Delete
        ado_sales.Recordset.Update
        ado_sales.Refresh
        Call Sales_ListLoad
        End If
    Else
        If lstview_expenses.ListItems.Count = 0 Then Exit Sub
        answer = MsgBox("You are about to DELETE record, Do you want to continue?", vbExclamation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        
        ado_expenses.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
        ado_expenses.RecordSource = "select * from tbl_expenses where TN = " & lstview_expenses.SelectedItem.SubItems(3)
        ado_expenses.Refresh
        ado_expenses.Recordset.Delete
        ado_expenses.Recordset.Update
        ado_expenses.Refresh
        Call cmd_today_Click
        End If
    End If

'err: Exit Sub



End Sub

Private Sub cmd_print_sales_Click()
' Dim strQryString As String
' Dim rs As New ADODB.Recordset
' 'set rs =
'strQryString = "{tbl_sales.date_of_sale} >= '" & Format(DTPicker1.Value, "yyyy/MM/dd") & "' And {tbl_sales.date_of_sale} <= '" & Format(DTPicker2.Value, "yyyy/MM/dd") & "' and {tbl_sales.Customer_name}  like ""*" & txt_search.Text & "*"""
' 'strQryString = "{tbl_sales.Customer_name}  like ""*" & txt_search.Text & "*"""
'    With CrystalReport1
'        .ReportFileName = App.Path & "\sales.rpt"
'        .Connect = gvConnectString2
'        .DiscardSavedData = True
'        .RetrieveDataFiles
'        .ReportSource = 0
'        .SQLQuery = "Select*  from tbl_sales "
'        .ReportTitle = "Water Willy Sales Report"
'        .Destination = crptToWindow
'        .PrintFileType = crptCrystal
'        .WindowState = crptMaximized
'        .WindowMaxButton = False
'        .WindowMinButton = False
'        .SelectionFormula = strQryString
'        .Action = 1
'    End With
    If fm_sales.Visible = True Then
        gvReportValue = 2
        If gvReportValue = 2 Then
            Set frmPreview2 = New frmPreview
            frmPreview2.Show
        End If
    Else: gvReportValue = 5
        If gvReportValue = 5 Then
            Set frmPreview5 = New frmPreview
            frmPreview5.Show
        End If

    End If
End Sub

Private Sub cmd_today_Click()
DTPicker1.Value = Now
DTPicker2.Value = Now
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub

Private Sub cmd_yesterday_Click()
DTPicker1.Value = Now - 1
DTPicker2.Value = Now - 1
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses

End Sub

Private Sub cmd_thisweek_Click()
DTPicker1.Value = DateAdd("d", -Weekday(Date) + 1, Date)
DTPicker2.Value = DateAdd("d", -Weekday(Date) + 7, Date)
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub

Private Sub cmd_thismonth_Click()
DTPicker1.Value = DateSerial(Year(Date), Month(Date), 1)
DTPicker2.Value = DateSerial(Year(Date), Month(Date) + 1, 0)
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub





Private Sub Command5_Click()

End Sub

Private Sub DTPicker1_Change()
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub



Private Sub DTPicker2_Change()
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub

Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
Call ConnDataBaseSales
Call List2LOADSales

Call ConnDataBaseExpenses
Call List2LOADExpenses
End Sub

Private Sub fm_expenses_Click()
fm_sales.Visible = True
fm_expenses.Visible = False
cmd_print_sales.Caption = "Print Sales Report"
End Sub

Private Sub fm_sales_Click()
fm_expenses.Visible = True
fm_sales.Visible = False
cmd_print_sales.Caption = "Print Expenses Report"
End Sub

Private Sub Form_Activate()
Call cmd_today_Click
Call Conn_ado_sales
Call Sales_ListLoad
End Sub

Private Sub Form_Load()

Select Case mdi_wrs.WindowState
Case Is = vbNormal
mdi_wrs.Width = frm_sales.Width + 315
mdi_wrs.Height = frm_sales.Height + 955
DoCenter frm_sales, mdi_wrs
Case Is = vbMaximized
DoCenter frm_sales, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select
End Sub

Private Sub Label3_Click()

End Sub

Private Sub lbl_expenses_Change()
lbl_profit.Caption = Val(lbl_sales.Caption) - Val(lbl_expenses.Caption)
If Val(lbl_profit.Caption) <= 0 Then
lbl_profit.ForeColor = vbRed
Else
lbl_profit.ForeColor = vbBlack
End If
End Sub


Private Sub lbl_sales_Change()
lbl_profit.Caption = Val(lbl_sales.Caption) - Val(lbl_expenses.Caption)
If Val(lbl_profit.Caption) <= 0 Then
lbl_profit.ForeColor = vbRed
Else
lbl_profit.ForeColor = vbBlack
End If

End Sub


Private Sub lstview_sales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstview_sales.SortKey = ColumnHeader.Index - 1

    If lstview_sales.SortKey = 1 Then lstview_sales.SortKey = 3 ' **** This column changes the key
    lstview_sales.SortOrder = (lstview_sales.SortOrder - 1) * -1
    lstview_sales.Sorted = True
End Sub

Private Sub txt_search_Change()

Call Conn_ado_sales
    ado_sales.RecordSource = "select * from tbl_sales where Customer_name like  '%" & txt_search.text & "%'  and date_of_sale between  '" & Format(frm_sales.DTPicker1.Value, "yyyy/mm/dd") & "'  and  '" & Format(frm_sales.DTPicker2.Value, "yyyy/mm/dd") & " ' "
    
    ado_sales.Refresh
    lstview_sales.ListItems.Clear
    Do Until ado_sales.Recordset.EOF
    Set list = lstview_sales.ListItems.Add(, , ado_sales.Recordset(0) & "")
    
    list.SubItems(1) = ado_sales.Recordset(1) & ""
    list.SubItems(2) = ado_sales.Recordset(2) & ""
    list.SubItems(3) = ado_sales.Recordset(3) & ""
    list.SubItems(4) = ado_sales.Recordset(4) & ""
    list.SubItems(5) = ado_sales.Recordset(5) & ""
    list.SubItems(6) = ado_sales.Recordset(6) & ""
    list.SubItems(7) = ado_sales.Recordset(7) & ""
    
        
    ado_sales.Recordset.MoveNext
        Loop
    
    With lstview_sales
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    gvCmd = ""
    gvCmd = ado_sales.RecordSource
    
    Set list = Nothing

End Sub
