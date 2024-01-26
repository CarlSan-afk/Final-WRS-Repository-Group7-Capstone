VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_itemlist 
   BorderStyle     =   0  'None
   Caption         =   "Item List"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_itemlist.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fbg 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      TabIndex        =   46
      Top             =   8160
      Width           =   3855
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Item List"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   240
         TabIndex        =   47
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.Frame fm_listview 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   4200
      Left            =   4440
      TabIndex        =   9
      Top             =   480
      Width           =   10575
      Begin MSComctlLib.ListView lstview_itemlist 
         Height          =   3135
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Title"
            Object.Width           =   5794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   5794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "POS Item"
            Object.Width           =   5794
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   2250
         Left            =   8160
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label lbl_addedit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stocks and Expenses"
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
         Left            =   360
         TabIndex        =   41
         Top             =   240
         Width           =   4440
      End
   End
   Begin MSAdodcLib.Adodc ado_itemlist 
      Height          =   330
      Left            =   960
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
   Begin VB.Frame fm_discription 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5295
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      Begin VB.TextBox txtQrCode 
         Appearance      =   0  'Flat
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txt_discription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3720
         Width           =   3375
      End
      Begin VB.TextBox txt_positem 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox txt_type 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txt_itemtitle 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QR Code"
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
         Left            =   255
         TabIndex        =   42
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discription:"
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
         TabIndex        =   8
         Top             =   3360
         Width           =   1260
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POS Item"
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
         TabIndex        =   6
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Top             =   1680
         Width           =   555
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
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   1020
      End
   End
   Begin MSComctlLib.Toolbar Toolbar_main 
      Height          =   840
      Left            =   4320
      TabIndex        =   15
      Top             =   8040
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1482
      ButtonWidth     =   3678
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "img_main"
      DisabledImageList=   "img_main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "D e l e t e    "
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "S t o c k s    "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E d i t   I t e m   "
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "A d d   I t e m  "
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print QRCode"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "frm_itemlist.frx":DBEB
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   240
      Top             =   8760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":1D045
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":1F61F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":21BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":241D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":267AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":28D87
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":2B361
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":2D93B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":2FF15
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_itemlist.frx":589FD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fm_HC 
      BackColor       =   &H00FFFF80&
      Caption         =   "Household Charges"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4440
      TabIndex        =   17
      Top             =   4800
      Width           =   3375
      Begin VB.Frame fm_charges 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
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
         Height          =   2535
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3135
         Begin VB.TextBox txt_purchase_HC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   40
            Top             =   2040
            Width           =   2655
         End
         Begin VB.TextBox txt_pickup_HC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   22
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txt_delivery_HC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Value"
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
            TabIndex        =   25
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pick up Charges"
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
            TabIndex        =   24
            Top             =   1080
            Width           =   1800
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Charge"
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
            TabIndex        =   23
            Top             =   360
            Width           =   1845
         End
      End
   End
   Begin VB.Frame fm_DC 
      BackColor       =   &H00FFFF80&
      Caption         =   "Dealer Charges"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   11640
      TabIndex        =   18
      Top             =   4800
      Width           =   3375
      Begin VB.Frame Frame2 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
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
         Height          =   2535
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   3135
         Begin VB.TextBox txt_purchase_DC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   29
            Top             =   2040
            Width           =   2655
         End
         Begin VB.TextBox txt_delivery_DC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txt_pickup_DC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   27
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Charge"
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
            TabIndex        =   32
            Top             =   360
            Width           =   1845
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pick up Charges"
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
            TabIndex        =   31
            Top             =   1080
            Width           =   1800
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Value"
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
            TabIndex        =   30
            Top             =   1800
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fm_RC 
      BackColor       =   &H00FFFF80&
      Caption         =   "Reseller Charges"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   8040
      TabIndex        =   19
      Top             =   4800
      Width           =   3375
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Height          =   2535
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   3135
         Begin VB.TextBox txt_purchase_RC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   39
            Top             =   2040
            Width           =   2655
         End
         Begin VB.TextBox txt_delivery_RC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   35
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txt_pickup_RC 
            Appearance      =   0  'Flat
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
            Left            =   240
            TabIndex        =   34
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Charge"
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
            TabIndex        =   38
            Top             =   360
            Width           =   1845
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pick up Charges"
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
            TabIndex        =   37
            Top             =   1080
            Width           =   1800
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Value"
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
            TabIndex        =   36
            Top             =   1800
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fm_stocks 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2295
      Left            =   360
      TabIndex        =   10
      Top             =   6000
      Width           =   3855
      Begin VB.TextBox txt_damaged 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   43
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txt_borrowed 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txt_available 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Damaged / Lost / Removed"
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
         TabIndex        =   44
         Top             =   1440
         Width           =   3030
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borrowed"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available in Stocks"
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
         TabIndex        =   13
         Top             =   0
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_itemlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QRGen As clsQRGen

Private Sub Form_Load()
Set QRGen = New clsQRGen
Select Case mdi_wrs.WindowState
Case Is = vbNormal
    mdi_wrs.Width = frm_itemlist.Width + 315
    mdi_wrs.Height = frm_itemlist.Height + 955
    DoCenter frm_itemlist, mdi_wrs
Case Is = vbMaximized
    DoCenter frm_itemlist, mdi_wrs
Case Is = vbMinimized
    Exit Sub
End Select


Call Itemlist_ListLoad
Call lstview_itemlist_Click

End Sub


Private Sub Label18_Click()
    gvReportValue = 7
    If gvReportValue = 7 Then
        Set frmPreview7 = New frmPreview
        frmPreview7.Show
    End If
End Sub


Public Sub lstview_itemlist_Click()
On Error GoTo err
Call Conn_ado_itemlist
ado_itemlist.RecordSource = "select * from tbl_item_list where Item_title like  '%" & lstview_itemlist.SelectedItem.text & "%'"
ado_itemlist.Refresh
txt_itemtitle.text = ado_itemlist.Recordset(0)
txt_type.text = ado_itemlist.Recordset(1)
txt_positem.text = ado_itemlist.Recordset(2)

txt_delivery_HC.text = ado_itemlist.Recordset(3)
txt_pickup_HC.text = ado_itemlist.Recordset(4)
txt_purchase_HC.text = ado_itemlist.Recordset(5)

txt_delivery_RC.text = ado_itemlist.Recordset(6)
txt_pickup_RC.text = ado_itemlist.Recordset(7)
txt_purchase_RC.text = ado_itemlist.Recordset(8)

txt_delivery_DC.text = ado_itemlist.Recordset(9)
txt_pickup_DC.text = ado_itemlist.Recordset(10)
txt_purchase_DC.text = ado_itemlist.Recordset(11)

txt_discription.text = ado_itemlist.Recordset(12)
txt_available.text = ado_itemlist.Recordset(13)
txt_borrowed.text = ado_itemlist.Recordset(14)

txt_damaged.text = ado_itemlist.Recordset(15)
txtQrCode.text = Nz(ado_itemlist.Recordset(16), "")

Setup
Set frm_itemlist.Image1.Picture = QRGen.QRCodegenBarcode(txtQrCode.text)
    
If lstview_itemlist.SelectedItem.SubItems(2) = "No" Then
Toolbar_main.Buttons(2).Caption = "Add Expenses"
fm_HC.Visible = False
fm_RC.Visible = False
fm_DC.Visible = False
fm_stocks.Visible = False
fm_listview.Height = 7320
lstview_itemlist.Height = 6150
fm_discription.Height = 7335
txt_discription.Height = 4335


Toolbar_main.Buttons(1).Enabled = True
Else
Toolbar_main.Buttons(2).Caption = "S t o c k s   "
Toolbar_main.Buttons(1).Enabled = False
fm_HC.Visible = True
fm_RC.Visible = True
fm_DC.Visible = True
fm_stocks.Visible = True
fm_listview.Height = 4200
lstview_itemlist.Height = 3135
fm_discription.Height = 5295
txt_discription.Height = 2295


End If


err:
Exit Sub
End Sub

Private Sub lstview_itemlist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
lstview_itemlist.SortKey = ColumnHeader.Index - 1

    If lstview_itemlist.SortKey = 1 Then lstview_itemlist.SortKey = 3 ' **** This column changes the key
    lstview_itemlist.SortOrder = (lstview_itemlist.SortOrder - 1) * -1
    lstview_itemlist.Sorted = True
End Sub

Private Sub lstview_itemlist_KeyDown(KeyCode As Integer, Shift As Integer)
Call lstview_itemlist_Click
End Sub

Private Sub lstview_itemlist_KeyUp(KeyCode As Integer, Shift As Integer)
Call lstview_itemlist_Click
End Sub

Private Sub Toolbar_main_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case Is = 1
'delete
On Error Resume Next
answer = MsgBox("You are about to DELETE this Record!, Do you want to PROCEED?", vbExclamation + vbOKCancel, "Confirm")
    If answer = vbOK Then
        ado_itemlist.Recordset.Delete
        ado_itemlist.Recordset.Update
        ado_itemlist.Refresh
        Call lstview_itemlist_Click
        Call Itemlist_ListLoad
    Else
    End If

Case Is = 2
'Add Stocks
    If Toolbar_main.Buttons(2).Caption = "S t o c k s   " Then
        frm_addstocks.Show
        frm_addstocks.txt_itemtitle.text = txt_itemtitle.text
        mdi_wrs.Enabled = False
    Else 'add expenses
        frm_addexpenses.Show
        frm_addexpenses.txt_itemtitle.text = txt_itemtitle.text
        mdi_wrs.Enabled = False
    End If

Case Is = 3
    'Edit Item
    mdi_wrs.Enabled = False
    frm_additem.Show
    frm_additem.lbl_addedit.Caption = "EDIT ITEM"
    frm_additem.Caption = "EDIT ITEM"
    frm_additem.txt_itemtitle.Enabled = False
    frm_additem.txt_itemtitle.text = txt_itemtitle.text
    frm_additem.cbo_type.text = txt_type.text
    frm_additem.cbo_pos.text = txt_positem.text
    
    frm_additem.txt_delivery_HC.text = txt_delivery_HC.text
    frm_additem.txt_pickup_HC.text = txt_pickup_HC.text
    frm_additem.txt_purchase_HC.text = txt_purchase_HC.text
    
    frm_additem.txt_delivery_RC.text = txt_delivery_RC.text
    frm_additem.txt_pickup_RC.text = txt_pickup_RC.text
    frm_additem.txt_purchase_RC.text = txt_purchase_RC.text
    
    frm_additem.txt_delivery_DC.text = txt_delivery_DC.text
    frm_additem.txt_pickup_DC.text = txt_pickup_DC.text
    frm_additem.txt_purchase_DC.text = txt_purchase_DC.text
    
    frm_additem.txt_discription.text = txt_discription.text
    frm_additem.txtQrCode.text = txtQrCode.text
    'Set frm_additem.Image1.Picture = QRGen.QRCodegenBarcode(txtQrCode.Text)
Case Is = 4
    'add item
    mdi_wrs.Enabled = False
    frm_additem.Show
    frm_additem.lbl_addedit.Caption = "ADD ITEM"
    frm_additem.Caption = "ADD ITEM"

Case Is = 5
        'If txt_positem.text = "No" Then
        If Me.lstview_itemlist.SelectedItem.SubItems(2) = "No" Then
            MsgBox "Non POS-item unable to print QR Code.!", vbInformation
            Exit Sub
        End If
        gvReportValue = 4
        If gvReportValue = 4 Then
            Set frmPreview4QR = New frmPreview
            frmPreview4QR.Show
        End If
End Select
End Sub
Private Sub Setup()
    QRGen.xBoostEcl = 1
    QRGen.xColor = "Black"
    QRGen.xEcl = CLng("1")
    QRGen.xMask = CLng("-1")
    QRGen.xMaxVersion = 40
    QRGen.xMinVersion = 1
    QRGen.xModuleSize = 120
    QRGen.xSquare = 1
End Sub
