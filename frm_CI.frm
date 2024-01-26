VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_CI 
   BorderStyle     =   0  'None
   Caption         =   "Customers Information"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_CI.frx":0000
   ScaleHeight     =   9705
   ScaleMode       =   0  'User
   ScaleWidth      =   31223.63
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   480
      TabIndex        =   24
      Top             =   720
      Width           =   10215
      Begin MSComctlLib.Toolbar tb_mainbutton 
         Height          =   810
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   1429
         ButtonWidth     =   3598
         ButtonHeight    =   1429
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "img_main"
         DisabledImageList=   "img_main"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "SAVE           "
               ImageIndex      =   9
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ADD NEW       "
               ImageIndex      =   10
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "EDIT           "
               ImageIndex      =   11
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "RETURN         "
               ImageIndex      =   13
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "BORROW         "
               ImageIndex      =   14
            EndProperty
         EndProperty
         MouseIcon       =   "frm_CI.frx":DBEB
      End
   End
   Begin MSComDlg.CommonDialog cd_upload 
      Left            =   11040
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_clientsitem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "Client's Item"
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmd_image 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "Image"
      Height          =   375
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc ado_ci 
      Height          =   330
      Left            =   120
      Top             =   9120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "ado_ci"
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
      Caption         =   "Search"
      Height          =   735
      Left            =   360
      TabIndex        =   16
      Top             =   5520
      Width           =   6615
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Print"
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cbo_search 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frm_CI.frx":1D045
         Left            =   120
         List            =   "frm_CI.frx":1D058
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txt_search 
         Appearance      =   0  'Flat
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
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Customer List"
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   360
      TabIndex        =   14
      Top             =   6360
      Width           =   14775
      Begin MSComctlLib.ListView lstview_ci 
         Height          =   2295
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID Number"
            Object.Width           =   4238
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Name"
            Object.Width           =   4238
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Classification"
            Object.Width           =   4238
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Address"
            Object.Width           =   4238
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Facebook"
            Object.Width           =   4238
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Transaction Date"
            Object.Width           =   4238
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Image Path"
            Object.Width           =   4238
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Contact"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ISP"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Customer's Information"
      Height          =   3735
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   10455
      Begin VB.ComboBox cboIsp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frm_CI.frx":1D094
         Left            =   2400
         List            =   "frm_CI.frx":1D0AA
         TabIndex        =   29
         Text            =   "Smart"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txt_contact 
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   26
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txt_facebooklink 
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txt_transactiondate 
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   7080
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ComboBox cbo_classification 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         Height          =   405
         ItemData        =   "frm_CI.frx":1D0E2
         Left            =   7080
         List            =   "frm_CI.frx":1D0EF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txt_customername 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txt_idnumber 
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   7080
         TabIndex        =   7
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txt_address 
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   3000
         Width           =   7695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SimCard: "
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
         Left            =   975
         TabIndex        =   28
         Top             =   1680
         Width           =   1110
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number: "
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
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   1965
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Facebook Link:"
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
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Transaction Date"
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
         Left            =   7080
         TabIndex        =   15
         Top             =   1080
         Width           =   2355
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   960
         TabIndex        =   12
         Top             =   3120
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   7080
         TabIndex        =   11
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Left            =   7080
         TabIndex        =   10
         Top             =   2040
         Width           =   1635
      End
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   480
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":1D110
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":1F6EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":21CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":2429E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":26878
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":28E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":29F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":2B006
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":2C0E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":2E6BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":30C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":3326E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":35848
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CI.frx":37E22
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado_customeritem 
      Height          =   330
      Left            =   2400
      Top             =   9120
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin VB.Frame fm_clientsitem 
      Height          =   3015
      Left            =   10920
      TabIndex        =   18
      Top             =   2280
      Width           =   3975
      Begin MSComctlLib.ListView lstview_Clientitem 
         Height          =   2535
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Title"
            Object.Width           =   3294
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Borrowed Gallon"
            Object.Width           =   3294
         EndProperty
      End
   End
   Begin VB.Frame fm_image 
      Height          =   3015
      Left            =   10920
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Image img_image 
         Height          =   2655
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3735
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_CI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selection As Boolean


Private Sub cbo_search_Click()
Call CI_ListLoad
End Sub



Private Sub cmd_addtocalendar_Click()

End Sub

Private Sub cmd_clientsitem_Click()
fm_clientsitem.Visible = True
fm_image.Visible = False
End Sub

Private Sub cmd_image_Click()
fm_clientsitem.Visible = False
fm_image.Visible = True
End Sub



Private Sub cmd_print_Click()
'Dim strQryString As String
'Select Case cbo_search.text
'Case Is = "All"
'Case Is = "Customer Name"
'strQryString = "{tbl_customer_info.Customer_Name}  like ""*" & txt_search.text & "*"""
'Case Is = "Classification"
'strQryString = "{tbl_customer_info.Classification}  like ""*" & txt_search.text & "*"""
'Case Is = "Address"
'strQryString = "{tbl_customer_info.Address}  like ""*" & txt_search.text & "*"""
'End Select
'
'
'
'' strQryString = "{tbl_customer_info.Customer_Name}  like ""*" & txt_search.Text & "*"""
'    With CrystalReport1
'        .ReportFileName = App.path & "\customer.rpt"
'        .Connect = App.path & "\WRSv3.mdb"
'        .DiscardSavedData = True
'        .RetrieveDataFiles
'        .ReportSource = 0
'        .SQLQuery = "Select * from tbl_customer_info order "
'        .ReportTitle = "Customer Information"
'        .Destination = crptToWindow
'        .PrintFileType = crptCrystal
'        .WindowState = crptMaximized
'        .WindowMaxButton = False
'        .WindowMinButton = False
'        .SelectionFormula = strQryString
'        .Action = 1
'    End With
gvReportValue = 8
If gvReportValue = 8 Then
    Set frmPreview8 = New frmPreview
    frmPreview8.Show
End If
End Sub

Private Sub Form_Activate()
Call CI_ListLoad
End Sub

Private Sub Form_Load()

Select Case mdi_wrs.WindowState
Case Is = vbNormal
mdi_wrs.Width = frm_CI.Width + 315
mdi_wrs.Height = frm_CI.Height + 955
DoCenter frm_CI, mdi_wrs
Case Is = vbMaximized
DoCenter frm_CI, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select




On Error GoTo err
cbo_search.text = "Customer Name"
Call Conn_ado_ci
If ado_ci.Recordset.RecordCount = 0 Then
ado_ci.Recordset.AddNew
ado_ci.Recordset(0) = "1000"
ado_ci.Recordset(1) = "Guest"
ado_ci.Recordset(2) = "Household"
ado_ci.Recordset(3) = "N/A"
ado_ci.Recordset(4) = "https://www.facebook.com/waterwillywaterrefilling"
ado_ci.Recordset(5) = FormatDateTime(Now, vbShortDate)
ado_ci.Recordset(6) = LoadPicture(App.path & "\Default.jpg")
ado_ci.Recordset(7) = "N/A"
ado_ci.Recordset.Update
ado_ci.Refresh
End If

tb_mainbutton.Buttons(1).Enabled = False
tb_mainbutton.Buttons(3).Enabled = True
tb_mainbutton.Buttons(4).Enabled = True
cbo_classification.ListIndex = 0
txt_idnumber.text = ado_ci.Recordset(0)
txt_customername.text = ado_ci.Recordset(1)
cbo_classification.text = ado_ci.Recordset(2)
txt_address.text = ado_ci.Recordset(3)
txt_facebooklink.text = ado_ci.Recordset(4)
txt_transactiondate.text = ado_ci.Recordset(5)
img_image.Picture = LoadPicture(ado_ci.Recordset(6))
txt_contact.text = ado_ci.Recordset(7)
Call CI_ListLoad
err:
Call Conn_ado_ci
Exit Sub


End Sub





Private Sub img_image_DblClick()
If txt_customername.Enabled = False Then
Exit Sub
End If

cd_upload.Filter = "Picture File| *.JPG"
cd_upload.ShowOpen
If cd_upload.FileName <> "" Then
img_image.Picture = LoadPicture(cd_upload.FileName)
End If

End Sub

Private Sub lstview_ci_Click()
On Error GoTo err
Call Conn_ado_ci
txt_idnumber.text = lstview_ci.SelectedItem.text
txt_customername.text = lstview_ci.SelectedItem.SubItems(1)
cbo_classification.text = lstview_ci.SelectedItem.SubItems(2)
txt_address.text = lstview_ci.SelectedItem.SubItems(3)
txt_facebooklink.text = lstview_ci.SelectedItem.SubItems(4)
txt_transactiondate.text = lstview_ci.SelectedItem.SubItems(5)
img_image.Picture = LoadPicture(lstview_ci.SelectedItem.SubItems(6))
txt_contact.text = lstview_ci.SelectedItem.SubItems(7)
Me.cboIsp.text = lstview_ci.SelectedItem.SubItems(8)
ado_ci.RecordSource = "select * from tbl_customer_info where ID_number like  '%" & lstview_ci.SelectedItem.text & "%' order by date_of_last_buy desc"
ado_ci.Refresh
Exit Sub
err:
img_image.Picture = LoadPicture(App.path & "\Default.jpg")
'ado_ci.RecordSource = "select * from tbl_customer_info where ID_number like  '%" & lstview_ci.SelectedItem.Text & "%'"
'ado_ci.Refresh

End Sub

Private Sub lstview_ci_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstview_ci.SortKey = ColumnHeader.Index - 1

    If lstview_ci.SortKey = 1 Then lstview_ci.SortKey = 3 ' **** This column changes the key
    lstview_ci.SortOrder = (lstview_ci.SortOrder - 1) * -1
    lstview_ci.Sorted = True
End Sub

Private Sub lstview_ci_KeyDown(KeyCode As Integer, Shift As Integer)
Call lstview_ci_Click
End Sub

Private Sub lstview_ci_KeyUp(KeyCode As Integer, Shift As Integer)
Call lstview_ci_Click
End Sub

Public Sub tb_mainbutton_ButtonClick(ByVal Button As Button)

Select Case Button.Index
Case Is = 1

' save
On Error Resume Next
    If txt_idnumber.text = "" Or txt_customername.text = "" Or txt_address.text = "" Or txt_transactiondate.text = "" Then
    answer = MsgBox("Please Complete the Details?", vbExclamation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        
        Else
        MsgBox "Action canceled", vbInformation, "Confirm"
        ado_ci.Refresh
        txt_idnumber.text = ado_ci.Recordset(0)
        txt_customername.text = ado_ci.Recordset(1)
        cbo_classification = ado_ci.Recordset(2)
        txt_address.text = ado_ci.Recordset(3)
        txt_facebooklink.text = ado_ci.Recordset(4)
        txt_transactiondate.text = ado_ci.Recordset(5)
        img_image.Picture = LoadPicture(ado_ci.Recordset(6))
        txt_contact.text = ado_ci.Recordset(7)
        
        txt_customername.Enabled = False
        cbo_classification.Enabled = False
        txt_address.Enabled = False
        txt_facebooklink.Enabled = False
        txt_contact.Enabled = False
        
        tb_mainbutton.Buttons(1).Enabled = False
        tb_mainbutton.Buttons(2).Enabled = True
        tb_mainbutton.Buttons(3).Enabled = True
        tb_mainbutton.Buttons(4).Enabled = True
        tb_mainbutton.Buttons(5).Enabled = True
        
        lstview_ci.Enabled = True
        Exit Sub
        End If
    
    End If
        
    
    If selection = True Then
        If txt_customername.text = "" Then
            MsgBox "Customer name is required!.", vbInformation
            Exit Sub
        ElseIf Me.txt_address.text = "" Then
            MsgBox "Customer address is required!.", vbInformation
            Exit Sub
        End If
        ado_ci.RecordSource = "select * from tbl_customer_info where Customer_Name = '" & txt_customername.text & "' order by date_of_last_buy desc"
        ado_ci.Refresh
        answer = MsgBox("Inputed Data will be added,Do you want to proceed?", vbExclamation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        
            If ado_ci.Recordset.RecordCount = 0 Then
            ado_ci.Recordset.AddNew
            'ado_ci.Recordset(0) = txt_idnumber.Text
            ado_ci.Recordset(1) = txt_customername.text
            ado_ci.Recordset(2) = cbo_classification.text
            ado_ci.Recordset(3) = txt_address.text
            ado_ci.Recordset(4) = txt_facebooklink.text
            ado_ci.Recordset(5) = txt_transactiondate.text
            ado_ci.Recordset(6) = cd_upload.FileName
            ado_ci.Recordset(7) = txt_contact.text
            ado_ci.Recordset(8) = cboIsp.text
            ado_ci.Recordset.Update
            ado_ci.Refresh
            MsgBox "Saved!", vbInformation + vbOKOnly, "Water Refilling System"
            '=====================
            Call Conn_ado_history
            frm_history.ado_history.Recordset.AddNew
            frm_history.ado_history.Recordset(0) = txt_customername.text
            frm_history.ado_history.Recordset(1) = txt_idnumber.text
            frm_history.ado_history.Recordset(2) = cbo_classification.text
            frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
            frm_history.ado_history.Recordset(4) = " Added New Customer "
            frm_history.ado_history.Recordset.Update
            frm_history.ado_history.Refresh
            '=====================
            lstview_ci.Enabled = True
            Call Conn_ado_ci
            Call Conn_MDI
            Call Customer_ListLoad
            Else
            MsgBox "Customer Name is Already Exist!", vbExclamation + vbOKOnly, "Water Refilling System"
            lstview_ci.Enabled = True
            Exit Sub
            End If
        
        Else
            MsgBox "Action canceled", vbInformation, "Confirm"
            ado_ci.Refresh
            txt_search.text = "Guest"
            txt_idnumber.text = ado_ci.Recordset(0)
            txt_customername.text = ado_ci.Recordset(1)
            cbo_classification = ado_ci.Recordset(2)
            txt_address.text = ado_ci.Recordset(3)
            txt_facebooklink.text = ado_ci.Recordset(4)
            txt_transactiondate.text = ado_ci.Recordset(5)
            img_image.Picture = LoadPicture(ado_ci.Recordset(6))
            txt_contact.text = ado_ci.Recordset(7)
            
            txt_customername.Enabled = False
            cbo_classification.Enabled = False
            txt_address.Enabled = False
            txt_facebooklink.Enabled = False
            txt_contact.Enabled = False
            
            tb_mainbutton.Buttons(1).Enabled = False
            tb_mainbutton.Buttons(2).Enabled = True
            tb_mainbutton.Buttons(3).Enabled = True
            tb_mainbutton.Buttons(4).Enabled = True
            tb_mainbutton.Buttons(5).Enabled = True
            
            lstview_ci.Enabled = True
            Exit Sub
        End If
    Else
        ado_ci.Recordset(0) = txt_idnumber.text
        ado_ci.Recordset(1) = txt_customername.text
        ado_ci.Recordset(2) = cbo_classification.text
        ado_ci.Recordset(3) = txt_address.text
        ado_ci.Recordset(4) = txt_facebooklink.text
        ado_ci.Recordset(5) = txt_transactiondate.text
        ado_ci.Recordset(6) = cd_upload.FileName
        ado_ci.Recordset(7) = txt_contact.text
        ado_ci.Recordset(8) = cboIsp.text
        ado_ci.Recordset.Update
        ado_ci.Refresh
        MsgBox "Saved!", vbInformation + vbOKOnly, "Water Refilling System"
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = txt_customername.text
        frm_history.ado_history.Recordset(1) = txt_idnumber.text
        frm_history.ado_history.Recordset(2) = cbo_classification.text
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = " Edited Customer Information "
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
        Call Conn_ado_ci
        Call Conn_MDI
        Call Customer_ListLoad
        lstview_ci.Enabled = True
    End If
    tb_mainbutton.Buttons(1).Enabled = False
    tb_mainbutton.Buttons(2).Enabled = True
    tb_mainbutton.Buttons(3).Enabled = True
    tb_mainbutton.Buttons(4).Enabled = True
    tb_mainbutton.Buttons(5).Enabled = True
    txt_customername.Enabled = False
    cbo_classification.Enabled = False
    txt_address.Enabled = False
    txt_facebooklink.Enabled = False
    txt_contact.Enabled = False
    lstview_ci.Enabled = True
    Call CI_ListLoad
    Call Customer_ListLoad
    
Case Is = 2
' add new
Call Conn_ado_ci

txt_idnumber.text = ado_ci.Recordset.RecordCount + 1000
txt_transactiondate.text = FormatDateTime(Now, vbShortDate)
txt_customername.text = ""
txt_address.text = ""
txt_facebooklink.text = ""
img_image.Picture = LoadPicture(App.path & "\Default.jpg")
txt_contact.text = ""


tb_mainbutton.Buttons(1).Enabled = True
tb_mainbutton.Buttons(2).Enabled = False
tb_mainbutton.Buttons(3).Enabled = False
tb_mainbutton.Buttons(4).Enabled = False
tb_mainbutton.Buttons(5).Enabled = False
txt_customername.Enabled = True
cbo_classification.Enabled = True
txt_address.Enabled = True
txt_facebooklink.Enabled = True
lstview_ci.Enabled = False
txt_contact.Enabled = True
txt_customername.SetFocus
selection = True

Case Is = 3
'edit
If txt_idnumber.text = "1000" Then Exit Sub
If lstview_ci.ListItems.Count = 0 Then
Exit Sub
End If

selection = False
tb_mainbutton.Buttons(1).Enabled = True
tb_mainbutton.Buttons(2).Enabled = False
tb_mainbutton.Buttons(3).Enabled = False
tb_mainbutton.Buttons(4).Enabled = False
tb_mainbutton.Buttons(5).Enabled = False

txt_customername.Enabled = True
cbo_classification.Enabled = True
txt_address.Enabled = True
txt_facebooklink.Enabled = True
txt_contact.Enabled = True
lstview_ci.Enabled = False
Me.cboIsp.Enabled = True
Call lstview_ci_Click

Case Is = 4
'return gallon
If txt_idnumber.text = "1000" Then Exit Sub
If lstview_Clientitem.ListItems.Count = 0 Then Exit Sub

frm_return.Show
frm_return.lbl_customer.Caption = txt_customername.text
frm_return.lbl_idnumber.Caption = txt_idnumber.text
mdi_wrs.Enabled = False

Case Is = 5
'borrow gallon
If txt_idnumber.text = "1000" Then Exit Sub
frm_borrow.Show
frm_borrow.lbl_customer.Caption = txt_customername.text
frm_borrow.lbl_idnumber.Caption = txt_idnumber.text
mdi_wrs.Enabled = False



End Select
End Sub





Private Sub txt_address_Change()
If txt_address.text = "" Then
txt_address.BackColor = &H80FF80
Else
txt_address.BackColor = &HFFFFFF
End If
End Sub

Private Sub txt_customername_Change()
If txt_customername.text = "" Then
txt_customername.BackColor = &H80FF80
Else
txt_customername.BackColor = &HFFFFFF
End If

End Sub

Private Sub txt_contact_Change()
If txt_contact.text = "" Then
txt_contact.BackColor = &H80FF80
Else
txt_contact.BackColor = &HFFFFFF
End If

End Sub

Private Sub txt_facebooklink_Change()
If txt_facebooklink.text = "" Then
'txt_facebooklink.BackColor = &H80FF80
'Else
'txt_facebooklink.BackColor = &HFFFFFF
End If
End Sub

Private Sub txt_idnumber_Change()
Call Conn_ado_customeritemCI
ado_customeritem.RecordSource = "select * from tbl_customer_item where ID_number like  '%" & frm_CI.txt_idnumber.text & "%'"
ado_customeritem.Refresh
Call CustomerItemCI_ListLoad

End Sub

Private Sub txt_search_Change()
Call cbo_search_Click
Call lstview_ci_Click
End Sub

