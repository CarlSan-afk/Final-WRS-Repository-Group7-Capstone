VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frm_customer 
   BorderStyle     =   0  'None
   Caption         =   "Customer"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_customer.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_addtocalendar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Add to Calendar"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   7215
      Begin VB.CommandButton cmd_select 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1695
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
         Left            =   2160
         TabIndex        =   0
         Top             =   120
         Width           =   3135
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
         ItemData        =   "frm_customer.frx":DBEB
         Left            =   120
         List            =   "frm_customer.frx":DBFB
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   14415
      Begin MSComctlLib.ListView lstview_customer 
         Height          =   6975
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   12303
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID Number"
            Object.Width           =   2503
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   2503
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Classification"
            Object.Width           =   2503
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Last Transaction Date"
            Object.Width           =   2503
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address"
            Object.Width           =   2503
         EndProperty
      End
      Begin VB.Image img_location 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   6375
         Left            =   7440
         Picture         =   "frm_customer.frx":DC27
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6855
      End
   End
   Begin MSAdodcLib.Adodc ado_customer 
      Height          =   330
      Left            =   11400
      Top             =   8760
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
      Caption         =   "ado_customer"
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
   Begin MSAdodcLib.Adodc ado_task 
      Height          =   375
      Left            =   9000
      Top             =   8760
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
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
      Left            =   600
      TabIndex        =   8
      Top             =   8640
      Width           =   810
   End
   Begin VB.Label lbl_count 
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
      Left            =   1560
      TabIndex        =   7
      Top             =   8640
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER LIST"
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
      Left            =   9240
      TabIndex        =   6
      Top             =   480
      Width           =   3390
   End
End
Attribute VB_Name = "frm_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_search_Click()
Call Customer_ListLoad
lbl_count.Caption = lstview_customer.ListItems.Count
End Sub

Private Sub cmd_addtocalendar_Click()
If lstview_customer.ListItems.Count = 0 Then Exit Sub
Call Conn_ado_taskCUS
answer = MsgBox("Add " & lstview_customer.SelectedItem.SubItems(1) & " to Task Calendar Today!", vbonformation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        ado_task.Recordset.AddNew
        ado_task.Recordset(1) = lstview_customer.SelectedItem.SubItems(1)
        ado_task.Recordset(2) = "Assigned"
        ado_task.Recordset(3) = Now
        ado_task.Recordset.Update
        ado_task.Refresh
        Else
         MsgBox " Action Cancelled", vbInformation + vbOKOnly, "WRS Management System  "
        End If

End Sub

Private Sub cmd_select_Click()
On Error GoTo err

frm_main.txt_customername.text = ado_customer.Recordset(1)
frm_main.txt_idnumber.text = ado_customer.Recordset(0)
frm_main.txt_classification.text = ado_customer.Recordset(2)
frm_main.txt_address.text = ado_customer.Recordset(3)
frm_main.txt_contact.text = ado_customer.Recordset(7)
frm_main.txtIsp.text = ado_customer.Recordset(8)
Call Conn_ado_main
frm_main.ado_main.RecordSource = "select * from tbl_customer_info where ID_number = '" & ado_customer.Recordset(0) & "' order by date_of_last_buy desc"
frm_main.img_location.Picture = LoadPicture(ado_customer.Recordset(6))
frm_main.Show
Me.Hide
Call CustomerItem_ListLoad
Call Delivery_ListLoad
Exit Sub
err:
frm_main.img_location.Picture = LoadPicture(App.path & "\Default.jpg")
frm_main.Show
Me.Hide
Call CustomerItem_ListLoad
Call Delivery_ListLoad

Exit Sub

End Sub

Private Sub Form_Activate()
Call Customer_ListLoad
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 83 And (Shift And vbCtrlMask) > 0 Then 'Ctl + S
       Call cmd_select_Click  'Click button 1
    End If

End Sub

Private Sub Form_Load()
Select Case mdi_wrs.WindowState
Case Is = vbNormal
mdi_wrs.Width = frm_customer.Width + 315
mdi_wrs.Height = frm_customer.Height + 955
DoCenter frm_customer, mdi_wrs
Case Is = vbMaximized
DoCenter frm_customer, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select

cbo_search.ListIndex = 1
Call Conn_ado_customer
Call Customer_ListLoad
lbl_count.Caption = lstview_customer.ListItems.Count
End Sub



Private Sub lstview_customer_Click()
On Error GoTo err
Call Conn_ado_customer
ado_customer.RecordSource = "select * from tbl_customer_info where ID_number like  '%" & lstview_customer.SelectedItem.text & "%' order by date_of_last_buy desc"
ado_customer.Refresh
img_location.Picture = LoadPicture(ado_customer.Recordset(6))

Exit Sub
err:
img_location.Picture = LoadPicture(App.path & "\Default.jpg")



End Sub

Private Sub lstview_customer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstview_customer.SortKey = ColumnHeader.Index - 1

    If lstview_customer.SortKey = 1 Then lstview_customer.SortKey = 3 ' **** This column changes the key
    lstview_customer.SortOrder = (lstview_customer.SortOrder - 1) * -1
    lstview_customer.Sorted = True
End Sub





Private Sub lstview_customer_KeyDown(KeyCode As Integer, Shift As Integer)
Call lstview_customer_Click
End Sub

Private Sub lstview_customer_KeyUp(KeyCode As Integer, Shift As Integer)
Call lstview_customer_Click
End Sub

Private Sub txt_search_Change()
Call cbo_search_Click
lbl_count.Caption = lstview_customer.ListItems.Count

   

End Sub


