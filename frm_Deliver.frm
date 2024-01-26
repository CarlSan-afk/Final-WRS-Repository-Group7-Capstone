VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Deliver 
   BorderStyle     =   0  'None
   Caption         =   "Delivery"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_Deliver.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   11280
      Top             =   8880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc ado_delivery 
      Height          =   330
      Left            =   600
      Top             =   9000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "ado_delivery"
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
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   14415
      Begin MSComctlLib.ListView lstview_order 
         Height          =   6495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   11456
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Transaction #"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer Name"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Classification"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID Number"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date of Sale"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Delivered By"
            Object.Width           =   3205
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Item Title"
            Object.Width           =   6410
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   14280
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483646
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":DBEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":101C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":1279F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":14D79
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":17353
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":1842D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":19507
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":1A5E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":1CBBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":1F195
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Deliver.frx":2176F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   7335
      Begin MSComctlLib.Toolbar Toolbar_main 
         Height          =   810
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   1429
         ButtonWidth     =   3201
         ButtonHeight    =   1429
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "img_main"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Pay      "
               ImageIndex      =   8
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Delivered "
               ImageIndex      =   9
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancel    "
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Print     "
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frm_Deliver.frx":23D49
      End
   End
   Begin MSAdodcLib.Adodc ado_sales 
      Height          =   330
      Left            =   3720
      Top             =   9000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSAdodcLib.Adodc ado_credit 
      Height          =   330
      Left            =   6840
      Top             =   9000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FOR DELIVERY"
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
      Left            =   9000
      TabIndex        =   4
      Top             =   720
      Width           =   2925
   End
End
Attribute VB_Name = "frm_Deliver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call FORDelivery_ListLoad
End Sub

Private Sub Form_Load()

Select Case mdi_wrs.WindowState
Case Is = vbNormal
mdi_wrs.Width = frm_Deliver.Width + 315
mdi_wrs.Height = frm_Deliver.Height + 955
DoCenter frm_Deliver, mdi_wrs
Case Is = vbMaximized
DoCenter frm_Deliver, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select

End Sub


Private Sub Toolbar_main_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case Is = 1
' Pay
If lstview_order.ListItems.Count = 0 Then Exit Sub
If lstview_order.SelectedItem.text = 0 Then Exit Sub
frm_payment.Show
'frm_payment.txt_amounttopay.Text = Val(lstview_order.SelectedItem) * -1

frm_payment.txt_amounttopay.text = Val(lstview_order.SelectedItem.SubItems(1)) * -1
mdi_wrs.Enabled = False
frm_payment.Caption = "Pay Balance."
Case Is = 2
' Deliver
If lstview_order.ListItems.Count = 0 Then Exit Sub
Conn_ado_FORdelivery
ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & lstview_order.SelectedItem.SubItems(4) & "%'"
ado_delivery.Refresh
    Select Case Val(ado_delivery.Recordset(0))
    Case Is > 0
    'MsgBox " greater than"
    Case Is = 0
        If Val(ado_delivery.Recordset(1)) = 0 Then
        ado_delivery.Recordset.Delete
        ado_delivery.Recordset.Update
        ado_delivery.Refresh
        Else
        Call Conn_ado_salesDELIVER
        ado_sales.Recordset.AddNew
        ado_sales.Recordset(0) = ado_delivery.Recordset(1)
        ado_sales.Recordset(1) = ado_delivery.Recordset(2)
        ado_sales.Recordset(2) = ado_delivery.Recordset(3)
        ado_sales.Recordset(3) = ado_delivery.Recordset(4)
        ado_sales.Recordset(4) = FormatDateTime(Now, vbShortDate)
        ado_sales.Recordset(5) = ado_delivery.Recordset(6)
        ado_sales.Recordset(6) = ado_delivery.Recordset(7)
        ado_sales.Recordset.Update
        ado_sales.Refresh
        
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = ado_delivery.Recordset(2)
        frm_history.ado_history.Recordset(1) = ado_delivery.Recordset(4)
        frm_history.ado_history.Recordset(2) = ado_delivery.Recordset(3)
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = ado_delivery.Recordset(7) & "/Pay " & ado_delivery.Recordset(1)
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
        '*********************
        frm_task.ado_task.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
        frm_task.ado_task.RecordSource = "select * from tbl_task  where Status like  'Assigned' and Task like  '%" & lstview_order.SelectedItem.SubItems(1) & "%'"
        frm_task.ado_task.Refresh
        If frm_task.ado_task.Recordset.RecordCount = 0 Then
        Else
        frm_task.ado_task.Recordset(2) = "Completed"
        frm_task.ado_task.Recordset.Update
        frm_task.ado_task.Refresh
        
        frm_task.ado_task.Recordset.AddNew
        frm_task.ado_task.Recordset(1) = lstview_order.SelectedItem.SubItems(1)
        frm_task.ado_task.Recordset(2) = "Assigned"
        frm_task.ado_task.Recordset(3) = DateAdd("d", 3, Now)
        frm_task.ado_task.Recordset.Update
        frm_task.ado_task.Refresh
        End If
        '*********************
        
                
        Call Conn_ado_customer
        frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & ado_delivery.Recordset(4) & "%' order by date_of_last_buy desc"
        frm_customer.ado_customer.Refresh
        frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
        frm_customer.ado_customer.Recordset.Update
        frm_customer.ado_customer.Refresh
        frm_customer.ado_customer.Visible = False
        
        ado_delivery.Recordset.Delete
        ado_delivery.Recordset.Update
        ado_delivery.Refresh
        
        End If
       
    Case Is < 0
        Dim sales As Integer
        sales = Val(ado_delivery.Recordset(1)) - (Val(ado_delivery.Recordset(0)) * -1)
            If sales = 0 Then
            Else
            Call Conn_ado_salesDELIVER
            ado_sales.Recordset.AddNew
            ado_sales.Recordset(0) = sales
            ado_sales.Recordset(1) = ado_delivery.Recordset(2)
            ado_sales.Recordset(2) = ado_delivery.Recordset(3)
            ado_sales.Recordset(3) = ado_delivery.Recordset(4)
            ado_sales.Recordset(4) = FormatDateTime(Now, vbShortDate)
            ado_sales.Recordset(5) = ado_delivery.Recordset(6)
            ado_sales.Recordset(6) = ado_delivery.Recordset(7)
            ado_sales.Recordset.Update
            ado_sales.Refresh
            '=====================
            Call Conn_ado_history
            frm_history.ado_history.Recordset.AddNew
            frm_history.ado_history.Recordset(0) = ado_delivery.Recordset(2)
            frm_history.ado_history.Recordset(1) = ado_delivery.Recordset(4)
            frm_history.ado_history.Recordset(2) = ado_delivery.Recordset(3)
            frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
            'frm_history.ado_history.Recordset(4) = ado_delivery.Recordset(7) & "/Pay " & ado_delivery.Recordset(1)
            frm_history.ado_history.Recordset(4) = "/Remaining Balance: " & ado_delivery.Recordset(0) & "/" & ado_delivery.Recordset(7)
            frm_history.ado_history.Recordset.Update
            frm_history.ado_history.Refresh
            '=====================
            Call Conn_ado_customer
            frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & ado_delivery.Recordset(4) & "%' order by date_of_last_buy desc"
            frm_customer.ado_customer.Refresh
            frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
            frm_customer.ado_customer.Recordset.Update
            frm_customer.ado_customer.Refresh
            frm_customer.ado_customer.Visible = False
            End If
        
        Call Conn_ado_creditDELIVER
        ado_credit.RecordSource = "select * from tbl_credit where ID_number like  '%" & lstview_order.SelectedItem.SubItems(3) & "%'"
        ado_credit.Refresh
            If ado_credit.Recordset.RecordCount = 0 Then
            ado_credit.Recordset.AddNew
            ado_credit.Recordset(0) = ado_delivery.Recordset(0)
            ado_credit.Recordset(1) = ado_delivery.Recordset(1)
            ado_credit.Recordset(2) = ado_delivery.Recordset(2)
            ado_credit.Recordset(3) = ado_delivery.Recordset(3)
            ado_credit.Recordset(4) = ado_delivery.Recordset(4)
            ado_credit.Recordset(5) = ado_delivery.Recordset(5)
            ado_credit.Recordset(6) = ado_delivery.Recordset(6)
            ado_credit.Recordset(7) = ado_delivery.Recordset(7)
            ado_credit.Recordset.Update
            ado_credit.Refresh
            
            
            Call Conn_ado_customer
            frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & ado_delivery.Recordset(4) & "%' order by date_of_last_buy desc"
            frm_customer.ado_customer.Refresh
            frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
            frm_customer.ado_customer.Recordset.Update
            frm_customer.ado_customer.Refresh
            frm_customer.Visible = False
            ado_delivery.Recordset.Delete
            ado_delivery.Recordset.Update
            ado_delivery.Refresh
            
            Else
                answer = MsgBox("Customer has exiting Credit. Do you want to proceed?", vbExclamation + vbOKCancel, "Confirm")
                If answer = vbOK Then
                ado_credit.Recordset(0) = Val(ado_credit.Recordset(0)) + Val(ado_delivery.Recordset(0))
                ado_credit.Recordset(1) = Val(ado_credit.Recordset(1)) + Val(ado_delivery.Recordset(1))
                ado_credit.Recordset(2) = ado_delivery.Recordset(2)
                ado_credit.Recordset(3) = ado_delivery.Recordset(3)
                ado_credit.Recordset(4) = ado_delivery.Recordset(4)
                ado_credit.Recordset(5) = ado_delivery.Recordset(5)
                ado_credit.Recordset(6) = ado_delivery.Recordset(6)
                ado_credit.Recordset(7) = ado_delivery.Recordset(7)
                ado_credit.Recordset.Update
                ado_credit.Refresh
                
                 Call Conn_ado_customer
                frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & ado_delivery.Recordset(4) & "%' order by date_of_last_buy desc"
                frm_customer.ado_customer.Refresh
                frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
                frm_customer.ado_customer.Recordset.Update
                frm_customer.ado_customer.Refresh
                frm_customer.ado_customer.Visible = False
                ado_delivery.Recordset.Delete
                ado_delivery.Recordset.Update
                ado_delivery.Refresh
                Else
                MsgBox "Action canceled", vbInformation, "Confirm"
                Exit Sub
                
                End If
            End If
    
    End Select

Call Conn_ado_creditDELIVER
Call FORDelivery_ListLoad
Call Conn_ado_delivery
Call Delivery_ListLoad

'Call Credit_ListLoad



Case Is = 3
' Cancel
If lstview_order.ListItems.Count = 0 Then Exit Sub
answer = MsgBox("You are about to cancel the transcation!, Do you want to PROCEED?", vbExclamation + vbOKCancel, "Confirm")
If answer = vbOK Then
        
    If lstview_order.ListItems.Count = 0 Then Exit Sub
    Call Conn_ado_FORdelivery
    ado_delivery.RecordSource = "select * from tbl_delivery where Id_number like  '%" & lstview_order.SelectedItem.SubItems(4) & "%'"
    ado_delivery.Refresh
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = ado_delivery.Recordset(2)
        frm_history.ado_history.Recordset(1) = ado_delivery.Recordset(4)
        frm_history.ado_history.Recordset(2) = ado_delivery.Recordset(3)
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Cancelled " & ado_delivery.Recordset(7)
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
        Call Conn_ado_customer
        frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & ado_delivery.Recordset(4) & "%'order by date_of_last_buy desc"
        frm_customer.ado_customer.Refresh
        frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
        frm_customer.ado_customer.Recordset.Update
        frm_customer.ado_customer.Refresh
        frm_customer.ado_customer.Visible = False
    If ado_delivery.Recordset.EOF Then Exit Sub
    ado_delivery.Recordset.Delete
    ado_delivery.Recordset.Update
    ado_delivery.Refresh
       
    
    Call Conn_ado_FORdelivery
    Call FORDelivery_ListLoad
        
 Else
 Exit Sub
 End If
 

Case Is = 4
' Print
'Call Conn_ado_FORdelivery
'
' With CrystalReport1
'.ReportFileName = App.Path & "\delivery.rpt"
'.RetrieveDataFiles
'.WindowState = crptMaximized
'.Destination = crptToWindow
'.Action = 1
'
'
'End With
gvReportValue = 1
frmPreview.Show
End Select
        frm_sales.ado_sales.Visible = False
        frm_main.ado_delivery.Visible = False
        frm_sales.ado_sales.Visible = False
        frm_history.ado_history.Visible = False
        frm_credit.Visible = False
        frm_history.Visible = False
        frm_sales.Visible = False
End Sub
