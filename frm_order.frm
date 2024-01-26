VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_order 
   Caption         =   "Order"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   Icon            =   "frm_order.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_order.frx":10CA
   ScaleHeight     =   6810
   ScaleWidth      =   8235
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Height          =   2655
      Left            =   3720
      TabIndex        =   18
      Top             =   2040
      Width           =   4335
      Begin MSComctlLib.ListView lstview_order 
         Height          =   1455
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Title"
            Object.Width           =   2408
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Quantity"
            Object.Width           =   2408
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Total Amount"
            Object.Width           =   2408
         EndProperty
      End
      Begin VB.Label lbl_gallon 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   28
         Top             =   2160
         Width           =   90
      End
      Begin VB.Label lbl_item 
         BackColor       =   &H8000000D&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl_total 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   21
         Top             =   2160
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      Begin VB.TextBox txt_contact 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Text            =   "N/A"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txt_customername 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "Guest"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txt_idnumber 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Text            =   "1000"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txt_classification 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "Owned House"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txt_address 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   1
         Text            =   "N/A"
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact : "
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   29
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name :"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address : "
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   7
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number :"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4245
         TabIndex        =   6
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classification : "
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   600
         Width           =   1125
      End
   End
   Begin MSComctlLib.Toolbar Toolbar_main 
      Height          =   1140
      Left            =   4920
      TabIndex        =   19
      Top             =   5280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2011
      ButtonWidth     =   1799
      ButtonHeight    =   1852
      ImageList       =   "img_main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add to Cart"
            ImageIndex      =   4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Order"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "frm_order.frx":7C25
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   4200
      Top             =   5640
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
            Picture         =   "frm_order.frx":1707F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_order.frx":19659
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_order.frx":1BC33
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_order.frx":1E20D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_order.frx":207E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_order.frx":22DC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_order.frx":2539B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado_order 
      Height          =   330
      Left            =   6000
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "ado_order"
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
   Begin MSAdodcLib.Adodc ado_deliveryman 
      Height          =   330
      Left            =   3600
      Top             =   4680
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
      Caption         =   "ado_deliveryman"
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
      Left            =   600
      Top             =   6480
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   3255
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Text            =   "1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtQrCode 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   2775
      End
      Begin VB.Frame fm_promo 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   1560
         TabIndex        =   24
         Top             =   1560
         Width           =   1335
         Begin VB.OptionButton opt_10 
            BackColor       =   &H00FFFF80&
            Caption         =   "10 Plus 1"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton opt_5 
            BackColor       =   &H00FFFF80&
            Caption         =   "5  Plus 1"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton opt_reg 
            BackColor       =   &H00FFFF80&
            Caption         =   "Regular"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CheckBox chk_own 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Own Gallon"
         Height          =   195
         Left            =   3120
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cbo_deliveryman 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3720
         Width           =   2775
      End
      Begin VB.ComboBox cbo_quantity 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cbo_serviceoption 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frm_order.frx":27975
         Left            =   240
         List            =   "frm_order.frx":27982
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cbo_itemtitle 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frm_order.frx":2799D
         Left            =   240
         List            =   "frm_order.frx":2799F
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QR Code"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deliveryman"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Quantity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Option:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Title :"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   780
      End
   End
End
Attribute VB_Name = "frm_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_itemtitle_Click()
On Error GoTo err
ado_order.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
ado_order.RecordSource = "select * from tbl_item_list where Item_title like  '%" & cbo_itemtitle.text & "%'"

ado_order.Refresh
cbo_quantity.Clear
Select Case chk_own.Value
Case Is = 0
    For i = 0 To Nz(ado_order.Recordset(13).Value, 0)
    cbo_quantity.AddItem i
    Next
    cbo_quantity.ListIndex = 1
Case Is = 1
    For i = 1 To 50
    cbo_quantity.AddItem i
    Next
    cbo_quantity.ListIndex = 0
End Select

If Not Nz(ado_order.Recordset(16), "") = "" And cbo_itemtitle.text <> " " Then
   txtQrCode.text = ado_order.Recordset(16)
Else
    txtQrCode.text = ""
End If
If ado_order.Recordset(1) = "Bottle" Then
chk_own.Visible = False
Else
chk_own.Visible = True
End If

If cbo_serviceoption.text = "Buy" Then
chk_own.Visible = False
Else
chk_own.Visible = True
End If

Exit Sub
err: cbo_quantity.ListIndex = 0

End Sub




Private Sub cbo_serviceoption_Click()
If cbo_serviceoption.text = "Buy" Then
chk_own.Visible = False
fm_promo.Visible = False
chk_own.Value = 0
Call Toolbar_main_ButtonClick(Toolbar_main.Buttons(3))
Else
chk_own.Visible = True
fm_promo.Visible = True
End If
If Not cbo_serviceoption.text = "Deliver" Then
    cbo_deliveryman.Visible = False
    Label9.Visible = False
Else
    cbo_deliveryman.Visible = True
    Label9.Visible = True
End If
End Sub

Private Sub chk_own_Click()
Call cbo_itemtitle_Click

End Sub

Private Sub Form_Load()
cbo_itemtitle.Clear
ado_order.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
ado_order.RecordSource = "select * from tbl_item_list where pos_item ='Yes'"
ado_order.Refresh
cbo_itemtitle.AddItem " "
Do Until ado_order.Recordset.EOF
cbo_itemtitle.AddItem ado_order.Recordset(0)
ado_order.Recordset.MoveNext
Loop


cbo_deliveryman.Clear
ado_deliveryman.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
ado_deliveryman.RecordSource = "select * from tbl_deliveryman "
ado_deliveryman.Refresh
Do Until ado_deliveryman.Recordset.EOF
cbo_deliveryman.AddItem ado_deliveryman.Recordset(0)
ado_deliveryman.Recordset.MoveNext
Loop

cbo_itemtitle.ListIndex = 0
cbo_serviceoption.ListIndex = 0
cbo_deliveryman.ListIndex = 0

'Me.txtQrCode.Text = ""
End Sub

Private Sub Form_Terminate()
mdi_wrs.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdi_wrs.Enabled = True
End Sub



Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Public Sub Toolbar_main_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim list As ListItem
Dim i As Integer
Dim Index As Integer
Dim TotalValue As Double
Dim rs As New ADODB.Recordset

If Not IsNumeric(txtQuantity.text) Then
    MsgBox "Invalid Qunatity", vbOKOnly
    Exit Sub
End If

cbo_quantity.AddItem txtQuantity.text
cbo_quantity.text = txtQuantity.text

Select Case Button.Index
Case Is = 1
'Add to Cart
    If cbo_itemtitle.text = " " Then
        MsgBox "Invalid item", vbExclamation
        Exit Sub
    End If
    Set rs = gvConnection.Execute("select * from tbl_item_list where Item_title =  '" & cbo_itemtitle.text & "'")
    
    If rs("Stocks") < CInt(txtQuantity.text) Then
        MsgBox "" & rs("Item_title") & " " & "The item is less than the actual quantity, Availble Quantity" & " " & rs("Stocks"), vbExclamation
        Exit Sub
    End If
    If cbo_itemtitle.text = "" Or cbo_itemtitle.text = " " Then
        MsgBox "Please select item", vbOKOnly
        Exit Sub
    End If
    If cbo_quantity.text = "0" Then Exit Sub
    For i = 0 To lstview_order.ListItems.Count
    Set list = lstview_order.FindItem(cbo_itemtitle.text)
    If Not list Is Nothing Then
        MsgBox "Item already exists in the Order", vbOKOnly
        Exit Sub
    End If
    Next
        
    
    lbl_item.Caption = lbl_item.Caption & cbo_serviceoption.text & " (" & cbo_quantity.text & ") " & cbo_itemtitle.text & " / "

    
    Set list = lstview_order.ListItems.Add(, , cbo_itemtitle.text & "")
    list.SubItems(1) = cbo_quantity.text & ""
    
    Select Case txt_classification.text
    
    Case Is = "Household"
        If cbo_serviceoption.text = "Deliver" Then
            list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(3)) & ""
        ElseIf cbo_serviceoption.text = "Pick-up" Then
            list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(4)) & ""
        Else
            list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(5)) & ""
            ado_itemlist.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
            ado_itemlist.RecordSource = "select * from tbl_item_list where Item_title =  '" & cbo_itemtitle.text & "'"
            ado_itemlist.Refresh
            ado_itemlist.Recordset(13) = Val(ado_itemlist.Recordset(13)) - Val(cbo_quantity.text)
            'ado_itemlist.Recordset.Update
    
        End If
    Case Is = "Reseller"
       If cbo_serviceoption.text = "Deliver" Then
        list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(6)) & ""
        ElseIf cbo_serviceoption.text = "Pick-up" Then
        list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(7)) & ""
        Else
        list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(8)) & ""
        ado_itemlist.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
        ado_itemlist.RecordSource = "select * from tbl_item_list where Item_title =  '" & cbo_itemtitle.text & "'"
        ado_itemlist.Refresh
        ado_itemlist.Recordset(13) = Val(ado_itemlist.Recordset(13)) - Val(cbo_quantity.text)
        End If
    Case Is = "Dealer"
        If cbo_serviceoption.text = "Deliver" Then
        list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(9)) & ""
        ElseIf cbo_serviceoption.text = "Pick-up" Then
        list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(10)) & ""
        Else
        list.SubItems(2) = Val(cbo_quantity.text) * Val(ado_order.Recordset(11)) & ""
        ado_itemlist.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
        ado_itemlist.RecordSource = "select * from tbl_item_list where Item_title =  '" & cbo_itemtitle.text & "'"
        ado_itemlist.Refresh
        ado_itemlist.Recordset(13) = Val(ado_itemlist.Recordset(13)) - Val(cbo_quantity.text)
        End If
    End Select
    
    If opt_5.Value = True Then
    list.SubItems(1) = Val(cbo_quantity.text) + Int(Val(cbo_quantity) / 5) & ""
    End If
    
    If opt_10.Value = True Then
    list.SubItems(1) = Val(cbo_quantity.text) + Int(Val(cbo_quantity) / 10) & ""
    End If
    
        For Index = 1 To lstview_order.ListItems.Count
        TotalValue = TotalValue + lstview_order.ListItems(Index).SubItems(1)
        Next
        lbl_gallon.Caption = TotalValue
    
    
    Call matinik
    cbo_itemtitle = " "
    txtQrCode.text = ""
    txtQuantity.text = "1"
    txtQrCode.SetFocus


Case Is = 2
'Order
        If gvIsReturn = True Then
        Call Conn_ado_sales
        'If txt_payment.Text = "" Or txt_payment.Text = 0 Then Exit Sub
        frm_sales.ado_sales.Recordset.AddNew
        frm_sales.ado_sales.Recordset(0) = CInt("-" + lbl_total)
        frm_sales.ado_sales.Recordset(1) = frm_main.txt_customername
        frm_sales.ado_sales.Recordset(2) = frm_main.txt_classification
        frm_sales.ado_sales.Recordset(3) = frm_main.txt_idnumber
        frm_sales.ado_sales.Recordset(4) = FormatDateTime(Now, vbShortDate)
        'frm_main.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_main.txt_idnumber.Text & "%'"
        'frm_main.ado_delivery.Refresh
        'frm_sales.ado_sales.Recordset(5) = frm_main.ado_delivery.Recordset(6)
        frm_sales.ado_sales.Recordset(6) = lbl_item.Caption
        frm_sales.ado_sales.Recordset.Update
        frm_sales.ado_sales.Refresh
            
            '=====================
            Call Conn_ado_history
            frm_history.ado_history.Recordset.AddNew
            frm_history.ado_history.Recordset(0) = frm_main.txt_customername
            frm_history.ado_history.Recordset(1) = frm_main.txt_idnumber
            frm_history.ado_history.Recordset(2) = frm_main.txt_classification
            frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
            frm_history.ado_history.Recordset(4) = lbl_item.Caption
            frm_history.ado_history.Recordset.Update
            frm_history.ado_history.Refresh
            '=====================
            '*********************
            
            Unload Me
        End If

If lstview_order.ListItems.Count = 0 Then
Exit Sub
End If


If frm_payment.Visible = False Then
If cbo_serviceoption.text = "Deliver" And gvIsReturn = False Then
    If cbo_deliveryman.text = "N/A" Then
        MsgBox "Please select delivery man", vbExclamation
        Exit Sub
    End If
End If
frm_payment.Show
frm_payment.txt_amounttopay.text = lbl_total.Caption
frm_order.Enabled = False
frm_payment.Caption = "Pay Order"
If cbo_serviceoption.text = "Buy" Then
'    Dim lngIndex As Long
'    Dim lngTot As Long, cmd As String
'    'Dim rs As New ADODB.Recordset,
'    For lngIndex = 1 To frm_order.lstview_order.ListItems.Count
'        cmd = ""
'        cmd = "update tbl_item_list set stocks = stocks - " & frm_order.lstview_order.ListItems(lngIndex).SubItems(lngIndex) & " where Item_title = '" & frm_order.lstview_order.ListItems(lngIndex).Text & "'"
'        gvConnection.Execute (cmd)
'    Next
'    ado_itemlist.Recordset.Update
    ado_itemlist.Refresh
Else
End If

Else

'----------------


Call Conn_ado_delivery
frm_main.ado_delivery.Recordset.AddNew
frm_main.ado_delivery.Recordset(0) = frm_payment.txt_change.text
frm_main.ado_delivery.Recordset(1) = frm_payment.txt_amounttopay.text
frm_main.ado_delivery.Recordset(2) = txt_customername.text
frm_main.ado_delivery.Recordset(3) = txt_classification.text
frm_main.ado_delivery.Recordset(4) = txt_idnumber.text
frm_main.ado_delivery.Recordset(5) = FormatDateTime(Now, vbShortDate)
frm_main.ado_delivery.Recordset(6) = cbo_deliveryman.text
frm_main.ado_delivery.Recordset(7) = lbl_item.Caption
'frm_main.ado_delivery.Recordset (8)=
frm_main.ado_delivery.Recordset.Update
frm_main.ado_delivery.Refresh
Unload frm_payment
Unload frm_order
Call Delivery_ListLoad

'----------------


End If



Case Is = 3
'clear
On Error Resume Next
lstview_order.ListItems.Clear
lbl_item.Caption = ""
lbl_total.Caption = "0"
lbl_gallon.Caption = "0"

End Select

End Sub

Private Sub txtQrCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rs As New ADODB.Recordset
    Set rs = gvConnection.Execute("select * from tbl_item_list where qrCode =  '" & txtQrCode.text & "'")
    If Not rs.EOF Then
        cbo_itemtitle.text = rs("Item_title")
        txtQuantity.SetFocus
    End If
    
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    NumberOnly KeyAscii
End Sub
