VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_task 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task Calendar"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12435
   Icon            =   "frm_task.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   12435
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc ado_task 
      Height          =   375
      Left            =   240
      Top             =   7080
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
      Caption         =   "ado_task"
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
   Begin MSComctlLib.ListView lstview_task 
      Height          =   5295
      Left            =   8040
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
         Text            =   "Task"
         Object.Width           =   4881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   4882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.MonthView cal_task 
      Height          =   6210
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   10954
      _Version        =   393216
      ForeColor       =   -2147483635
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777152
      MultiSelect     =   -1  'True
      StartOfWeek     =   131072001
      TitleBackColor  =   -2147483635
      TitleForeColor  =   -2147483637
      CurrentDate     =   44515
   End
   Begin MSComctlLib.Toolbar tb_task 
      Height          =   870
      Left            =   8040
      TabIndex        =   3
      Top             =   5520
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1535
      ButtonWidth     =   3810
      ButtonHeight    =   1429
      TextAlignment   =   1
      ImageList       =   "img_main"
      DisabledImageList=   "img_main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Task          "
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Change Status   "
            ImageIndex      =   4
         EndProperty
      EndProperty
      MouseIcon       =   "frm_task.frx":10CA
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   7680
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_task.frx":10524
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_task.frx":12AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_task.frx":150D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_task.frx":176B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl_id 
      Caption         =   "Label1"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label lbl_task 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label lbl_date 
      Caption         =   "Label1"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
   End
End
Attribute VB_Name = "frm_task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cal_task_Click()
On Error Resume Next
If cal_task.Value < Now - 1 Then
tb_task.Buttons(1).Enabled = False
Else
tb_task.Buttons(1).Enabled = True
End If

Call Conn_ado_task
Call Task_ListLoad
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)
End Sub

Private Sub cal_task_DateClick(ByVal DateClicked As Date)
On Error Resume Next
Call Conn_ado_task
Call Task_ListLoad
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)

End Sub

Private Sub cal_task_DateDblClick(ByVal DateDblClicked As Date)
    frm_assignment.Show
    frm_assignment.cmd_add.Caption = "Add to Calendar"
    frm_assignment.cmd_delete.Visible = False
    lbl_date.Caption = cal_task.Value
    frm_task.Enabled = False
    frm_assignment.lbl_task.Caption = "Add task for " & cal_task.Value
End Sub



Private Sub cal_task_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Call Conn_ado_task
Call Task_ListLoad
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)
End Sub

Private Sub cal_task_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Call Conn_ado_task
Call Task_ListLoad
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)
End Sub

Private Sub cal_task_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
On Error Resume Next
Call Conn_ado_task
Call Task_ListLoad
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)
End Sub

Private Sub cal_task_Validate(Cancel As Boolean)
On Error Resume Next
If cal_task.Value < Now - 1 Then
tb_task.Buttons(1).Enabled = False
Else
tb_task.Buttons(1).Enabled = True
End If

Call Conn_ado_task
Call Task_ListLoad
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)
End Sub

Private Sub Form_Activate()
On Error Resume Next

ado_task.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
ado_task.RecordSource = "select * from tbl_task  where Status like  'Assigned' and Task_date >= '" & Format(Now, "yyyy-MM-dd") & "'"
ado_task.Refresh
Do Until ado_task.Recordset.EOF
    'ado_task.Recordset(3) = FormatDateTime(Now, vbShortDate)
    ado_task.Recordset.MoveNext
Loop
    ado_task.Recordset.Update
    ado_task.Refresh
Call Conn_ado_task
Call Task_ListLoad
cal_task.Value = FormatDateTime(Now, vbShortDate)
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)


End Sub

Private Sub Form_Load()
On Error Resume Next
Call Conn_ado_task
Call Task_ListLoad
cal_task.Value = FormatDateTime(Now, vbShortDate)
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)

End Sub


Private Sub lstview_task_Click()
On Error Resume Next
lbl_date.Caption = cal_task.Value
lbl_task.Caption = lstview_task.SelectedItem
lbl_id.Caption = lstview_task.SelectedItem.SubItems(2)
End Sub




Private Sub lstview_task_DblClick()

Call tb_task_ButtonClick(tb_task.Buttons(2))
End Sub

Private Sub tb_task_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case Is = 1
    frm_assignment.Show
    frm_assignment.cmd_add.Caption = "Add to Calendar"
    frm_assignment.cmd_delete.Visible = False
    frm_assignment.cmd_resched.Visible = False
    lbl_date.Caption = cal_task.Value
    frm_task.Enabled = False
    frm_assignment.lbl_task.Caption = "Add task for " & cal_task.Value
    
    Case Is = 2
    If lstview_task.ListItems.Count = 0 Then Exit Sub
    frm_assignment.Show
    frm_assignment.cmd_add.Caption = "Change Status"
    frm_assignment.cmd_delete.Visible = True
    frm_assignment.cmd_resched.Visible = True
    lbl_date.Caption = cal_task.Value
    frm_assignment.txt_task.text = lbl_task.Caption
    frm_task.Enabled = False
    frm_assignment.lbl_task.Caption = "Change Status Dated " & cal_task.Value
    
    End Select
End Sub
