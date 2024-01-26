VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   0  'None
   Caption         =   "Point of Sale"
   ClientHeight    =   9660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frm_main.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc ado_sales 
      Height          =   450
      Left            =   8760
      Top             =   8880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   794
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
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   7920
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
      Begin VB.Image img_fb 
         Height          =   960
         Left            =   120
         Picture         =   "frm_main.frx":DBEB
         Stretch         =   -1  'True
         Top             =   240
         Width           =   960
      End
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":101B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1278F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":14D69
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":17343
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1991D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1BEF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1E4D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":20AAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":23085
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   8895
      Begin VB.TextBox txtIsp 
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
         Left            =   5880
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txt_contact 
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
         Left            =   2880
         TabIndex        =   18
         Text            =   "N/A"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txt_address 
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
         Left            =   2880
         TabIndex        =   11
         Text            =   "N/A"
         Top             =   2520
         Width           =   5775
      End
      Begin VB.TextBox txt_classification 
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
         Left            =   2880
         TabIndex        =   10
         Text            =   "Owned House"
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox txt_idnumber 
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
         Left            =   2880
         TabIndex        =   9
         Text            =   "1000"
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txt_customername 
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
         Left            =   2880
         TabIndex        =   8
         Text            =   "Guest"
         Top             =   360
         Width           =   4335
      End
      Begin VB.Frame fm_customeritem 
         Caption         =   "Customer's Item"
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   8535
         Begin MSComctlLib.ListView lstview_CI 
            Height          =   1095
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   1931
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item Title"
               Object.Width           =   7104
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Borrowed Gallon"
               Object.Width           =   7104
            EndProperty
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number :"
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
         TabIndex        =   19
         Top             =   1560
         Width           =   1965
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classification : "
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
         TabIndex        =   7
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number :"
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
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address : "
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
         Top             =   2640
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name :"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc ado_main 
      Height          =   450
      Left            =   480
      Top             =   8880
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   794
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
      Caption         =   "ado_main"
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
   Begin MSAdodcLib.Adodc ado_customeritem 
      Height          =   450
      Left            =   3240
      Top             =   8880
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   794
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
   Begin MSAdodcLib.Adodc ado_delivery 
      Height          =   450
      Left            =   6000
      Top             =   8880
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   794
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
   Begin MSComctlLib.ImageList img_pay 
      Left            =   14280
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483646
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":2565F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":27C39
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":2A213
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":2C7ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":2D8C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":2FEA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":3247B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   975
      Left            =   9600
      TabIndex        =   16
      Top             =   5400
      Width           =   5295
      Begin MSComctlLib.Toolbar tb_pay 
         Height          =   810
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1429
         ButtonWidth     =   3096
         ButtonHeight    =   1429
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "img_pay"
         DisabledImageList=   "img_pay"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Pay     "
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Deliver    "
               ImageIndex      =   6
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancel       "
               ImageIndex      =   7
            EndProperty
         EndProperty
         MouseIcon       =   "frm_main.frx":34A55
      End
   End
   Begin MSAdodcLib.Adodc ado_credit 
      Height          =   450
      Left            =   11160
      Top             =   8880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   794
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
      Height          =   2655
      Left            =   360
      TabIndex        =   12
      Top             =   6480
      Width           =   14415
      Begin MSComctlLib.ListView lstview_order 
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   4260
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   14415
      Begin MSComctlLib.Toolbar Toolbar_main 
         Height          =   840
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   1482
         ButtonWidth     =   4260
         ButtonHeight    =   1429
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "img_main"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&CUSTOMER       "
               Object.ToolTipText     =   "F1"
               ImageIndex      =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&NEW CUSTOMER"
               Object.ToolTipText     =   "F2"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&RETURN GALLON"
               Object.ToolTipText     =   "F3"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&ORDER                 "
               Object.ToolTipText     =   "F4"
               ImageIndex      =   8
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BorderStyle     =   1
         MouseIcon       =   "frm_main.frx":43EAF
      End
      Begin VB.Label POINTOFSALE 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "POINT OF SALE"
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
         Height          =   735
         Left            =   9840
         TabIndex        =   20
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Image img_location 
      Height          =   4095
      Left            =   9600
      Picture         =   "frm_main.frx":46489
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   5295
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 67 And (Shift And vbCtrlMask) > 0 Then 'Ctl + C
        Toolbar_main_ButtonClick Toolbar_main.Buttons.Item(1)  'Click button 1
       
    ElseIf KeyCode = 78 And (Shift And vbCtrlMask) > 0 Then 'Ctl + N
        Toolbar_main_ButtonClick Toolbar_main.Buttons.Item(2) 'Click button 2
    ElseIf KeyCode = 82 And (Shift And vbCtrlMask) > 0 Then 'Ctl + R
        Toolbar_main_ButtonClick Toolbar_main.Buttons.Item(3) 'Click button 3
    ElseIf KeyCode = 79 And (Shift And vbCtrlMask) > 0 Then 'Ctl + O
        Toolbar_main_ButtonClick Toolbar_main.Buttons.Item(4) 'Click button 4
    ElseIf KeyCode = vbKeyF1 Then
         Toolbar_main_ButtonClick Toolbar_main.Buttons.Item(1)
    ElseIf KeyCode = vbKeyF2 Then
         Toolbar_main_ButtonClick Toolbar_main.Buttons.Item(2)
    ElseIf KeyCode = vbKeyF3 Then
         Toolbar_main_ButtonClick Toolbar_main.Buttons.Item(3)
    ElseIf KeyCode = vbKeyF4 Then
         Toolbar_main_ButtonClick Toolbar_main.Buttons.Item(4)
    ElseIf KeyCode = vbKeyF5 Then
         tb_pay_ButtonClick tb_pay.Buttons.Item(1)
    ElseIf KeyCode = vbKeyF6 Then
         tb_pay_ButtonClick tb_pay.Buttons.Item(2)
    ElseIf KeyCode = vbKeyF7 Then
         tb_pay_ButtonClick tb_pay.Buttons.Item(3)
    ElseIf KeyCode = 80 And (Shift And vbCtrlMask) > 0 Then 'Ctl + P
         tb_pay_ButtonClick tb_pay.Buttons.Item(1)
    ElseIf KeyCode = 68 And (Shift And vbCtrlMask) > 0 Then 'Ctl + D
         tb_pay_ButtonClick tb_pay.Buttons.Item(2)
    ElseIf KeyCode = 88 And (Shift And vbCtrlMask) > 0 Then 'Ctl + X
         tb_pay_ButtonClick tb_pay.Buttons.Item(3)
    End If

End Sub

Private Sub Form_Load()
Select Case mdi_wrs.WindowState
Case Is = vbNormal
'/GROUP7 - set mdi always Maximized
    mdi_wrs.Width = frm_main.Width + 315
    mdi_wrs.Height = frm_main.Height + 955
    DoCenter frm_main, mdi_wrs
'    mdi_wrs.WindowState = vbMaximized
'    DoCenter frm_main, mdi_wrs

Case Is = vbMaximized
    DoCenter frm_main, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select

txt_customername.text = "Guest"
txt_idnumber.text = "1000"
txt_classification.text = "Household"
txt_address.text = "N/A"
img_location.Picture = LoadPicture(App.path & "\Default.jpg")
txt_contact = "N/A"
Call Conn_ado_delivery
Call CustomerItem_ListLoad
Call Delivery_ListLoad
End Sub




Private Sub img_fb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim sURL As String
Dim quote As String
ado_main.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
ado_main.RecordSource = "select * from tbl_customer_info where ID_number like  '%" & txt_idnumber.text & "%'order by date_of_last_buy desc"
ado_main.Refresh

sURL = quote & " " & ado_main.Recordset(4) & " " & quote
Shell "C:\Program Files\Google\Chrome\Application\chrome.exe" & sURL, vbMaximizedFocus
Call SetWindowPos(mdi_wrs.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub img_fb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetWindowPos(mdi_wrs.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub img_fb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetWindowPos(mdi_wrs.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub tb_pay_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case Is = 1
' Pay Balance
If lstview_order.ListItems.Count = 0 Then Exit Sub
If lstview_order.SelectedItem.text = 0 Then Exit Sub
frm_payment.Show
'SelectedItem.SubItems (1)
'frm_payment.txt_amounttopay.Text = Val(lstview_order.SelectedItem) * -1
frm_payment.txt_amounttopay.text = Val(lstview_order.SelectedItem.SubItems(1)) * -1
mdi_wrs.Enabled = False
frm_payment.Caption = "Pay Balance"
Case Is = 2
' Deliver
'On Error Resume Next
If lstview_order.ListItems.Count = 0 Then Exit Sub
Conn_ado_delivery
ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
ado_delivery.Refresh

    Select Case Val(ado_delivery.Recordset(0))
      
    Case Is = 0
   
        If Val(ado_delivery.Recordset(1)) = 0 Then
        ado_delivery.Recordset.Delete
        ado_delivery.Recordset.Update
        ado_delivery.Refresh
        Call Delivery_ListLoad
        Else
        Call Conn_ado_salesMAIN
                
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
        frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & txt_idnumber.text & "%'order by date_of_last_buy desc"
        frm_customer.ado_customer.Refresh
        frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
        frm_customer.ado_customer.Recordset.Update
        frm_customer.ado_customer.Refresh
        
        ado_delivery.Recordset.Delete
        ado_delivery.Recordset.Update
        ado_delivery.Refresh
        Call Delivery_ListLoad
        Exit Sub
        End If
       
    Case Is < 0
        Dim sales As Integer
        sales = Val(ado_delivery.Recordset(1)) - (Val(ado_delivery.Recordset(0)) * -1)
            If sales = 0 Then
            Else
            Call Conn_ado_salesMAIN
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
            frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & txt_idnumber.text & "%'order by date_of_last_buy desc"
            frm_customer.ado_customer.Refresh
            frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
            frm_customer.ado_customer.Recordset.Update
            frm_customer.ado_customer.Refresh
        
            Call Delivery_ListLoad
            End If
        
        Call Conn_ado_creditMAIN
        ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
        ado_delivery.Refresh
        
        ado_credit.RecordSource = "select * from tbl_credit where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
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
            '=====================
            Call Conn_ado_history
            frm_history.ado_history.Recordset.AddNew
            frm_history.ado_history.Recordset(0) = ado_delivery.Recordset(2)
            frm_history.ado_history.Recordset(1) = ado_delivery.Recordset(4)
            frm_history.ado_history.Recordset(2) = ado_delivery.Recordset(3)
            frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
            frm_history.ado_history.Recordset(4) = "Credited " & ado_delivery.Recordset(0) & "Amounting to " & ado_delivery.Recordset(0)
            frm_history.ado_history.Recordset.Update
            frm_history.ado_history.Refresh
            '=====================
            Call Conn_ado_customer
            frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & txt_idnumber.text & "%' order by date_of_last_buy desc"
            frm_customer.ado_customer.Refresh
            frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
            frm_customer.ado_customer.Recordset.Update
            frm_customer.ado_customer.Refresh
            
            ado_delivery.Recordset.Delete
            ado_delivery.Recordset.Update
            ado_delivery.Refresh
            'Call Credit_ListLoad
            Call Delivery_ListLoad
        Else
            answer = MsgBox("Customer has exiting Credit. Do you want to proceed?", vbExclamation + vbOKCancel, "Confirm")
            Dim mvcredit As Currency, mvamount As Currency
            mvcredit = Val(ado_credit.Recordset(0)) + Val(ado_delivery.Recordset(0))
            mvamount = Val(ado_credit.Recordset(1)) + Val(ado_delivery.Recordset(1))
            If answer = vbOK Then
            ado_credit.Recordset(0) = mvcredit
            ado_credit.Recordset(1) = mvamount
            ado_credit.Recordset(2) = ado_delivery.Recordset(2)
            ado_credit.Recordset(3) = ado_delivery.Recordset(3)
            ado_credit.Recordset(4) = ado_delivery.Recordset(4)
            ado_credit.Recordset(5) = ado_delivery.Recordset(5)
            ado_credit.Recordset(6) = ado_delivery.Recordset(6)
            ado_credit.Recordset(7) = ado_delivery.Recordset(7)
            ado_credit.Recordset.Update
            ado_credit.Refresh
            '=====================
            Call Conn_ado_history
            frm_history.ado_history.Recordset.AddNew
            frm_history.ado_history.Recordset(0) = ado_delivery.Recordset(2)
            frm_history.ado_history.Recordset(1) = ado_delivery.Recordset(4)
            frm_history.ado_history.Recordset(2) = ado_delivery.Recordset(3)
            frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
            frm_history.ado_history.Recordset(4) = "Credited " & mvcredit
            frm_history.ado_history.Recordset.Update
            frm_history.ado_history.Refresh
            '=====================
            Call Conn_ado_customer
            frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where ID_number  like  '%" & txt_idnumber.text & "%' order by date_of_last_buy desc"
            frm_customer.ado_customer.Refresh
            frm_customer.ado_customer.Recordset(5) = FormatDateTime(Now, vbGeneralDate)
            frm_customer.ado_customer.Recordset.Update
            frm_customer.ado_customer.Refresh
        
        
            ado_delivery.Recordset.Delete
            ado_delivery.Recordset.Update
            ado_delivery.Refresh
            'Call Credit_ListLoad
            Call Delivery_ListLoad
            Else
            MsgBox "Action canceled", vbInformation, "Confirm"
            Exit Sub
            
            End If
        End If
           
    End Select

'Call Credit_ListLoad
Call Delivery_ListLoad


Case Is = 3
' Cancel
If lstview_order.ListItems.Count = 0 Then Exit Sub
answer = MsgBox("You are about to cancel the transcation!, Do you want to PROCEED?", vbExclamation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        
Call Conn_ado_delivery
ado_delivery.RecordSource = "select * from tbl_delivery where Id_number like  '%" & txt_idnumber.text & "%'"
ado_delivery.Refresh
            '=====================
            Call Conn_ado_history
            frm_history.ado_history.Recordset.AddNew
            frm_history.ado_history.Recordset(0) = ado_delivery.Recordset(2)
            frm_history.ado_history.Recordset(1) = ado_delivery.Recordset(4)
            frm_history.ado_history.Recordset(2) = ado_delivery.Recordset(3)
            frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
            frm_history.ado_history.Recordset(4) = "Canceled Delivery " & "with " & ado_delivery.Recordset(0) & " Balance"
            frm_history.ado_history.Recordset.Update
            frm_history.ado_history.Refresh
            '=====================
If ado_delivery.Recordset.EOF Then Exit Sub
ado_delivery.Recordset.Delete
ado_delivery.Recordset.Update
ado_delivery.Refresh

Call Conn_ado_delivery
Call Delivery_ListLoad
Else
Exit Sub
End If


End Select


End Sub

Private Sub Toolbar_main_ButtonClick(ByVal Button As MSComctlLib.Button)
gvIsReturn = False
Select Case Button.Index
Case Is = 1
'Customer
frm_customer.Show
Me.Hide
Case Is = 2
'New Customer
frm_CI.Show
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(2))
frm_CI.txt_search.text = ""
frm_CI.txt_customername.text = ""
frm_CI.txt_facebooklink.text = ""
frm_CI.txt_address.text = ""
frm_CI.txt_contact.text = ""


Me.Hide
Case Is = 3
'Return Gallon

If txt_idnumber.text = "1000" Then Exit Sub
If lstview_ci.ListItems.Count = 0 Then Exit Sub

frm_return.Show
frm_return.lbl_customer.Caption = txt_customername.text
frm_return.lbl_idnumber.Caption = txt_idnumber.text
mdi_wrs.Enabled = False


Case Is = 4
'order
If lstview_order.ListItems.Count > 0 Then
MsgBox "Deliver the Previous Order!", vbInformation, "Confirm"
Exit Sub
End If
frm_order.Show
frm_order.txt_customername.text = txt_customername.text
frm_order.txt_classification.text = txt_classification.text
frm_order.txt_address.text = txt_address.text
frm_order.txt_idnumber.text = txt_idnumber.text
frm_order.txt_contact.text = txt_contact.text
frm_order.txtQrCode.SetFocus
mdi_wrs.Enabled = False
Case Is = 5
    frm_order.Caption = "Return Order"
    gvIsReturn = True
    
frm_order.Show
frm_order.txt_customername.text = txt_customername.text
frm_order.txt_classification.text = txt_classification.text
frm_order.txt_address.text = txt_address.text
frm_order.txt_idnumber.text = txt_idnumber.text
frm_order.txt_contact.text = txt_contact.text
mdi_wrs.Enabled = False

End Select

End Sub


