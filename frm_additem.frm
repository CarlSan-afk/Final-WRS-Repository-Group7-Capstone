VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_additem 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADD ITEM"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10995
   Icon            =   "frm_additem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_additem.frx":10CA
   ScaleHeight     =   5520
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9000
      TabIndex        =   36
      Text            =   "Black"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   8400
      TabIndex        =   34
      Top             =   600
      Width           =   2535
   End
   Begin VB.Frame fm_charges_DC 
      BackColor       =   &H00FFFF80&
      Caption         =   "Dealer Charges"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      TabIndex        =   26
      Top             =   3720
      Width           =   3855
      Begin VB.TextBox txt_delivery_DC 
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
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txt_purchase_DC 
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
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txt_pickup_DC 
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
         Left            =   2160
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Charge"
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
         TabIndex        =   32
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Value"
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
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick-up Charge"
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
         Left            =   2160
         TabIndex        =   30
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame fm_charges_RC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reseller Charges"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      TabIndex        =   19
      Top             =   2040
      Width           =   3855
      Begin VB.TextBox txt_delivery_RC 
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
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txt_purchase_RC 
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
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txt_pickup_RC 
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
         Left            =   2160
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Charge"
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
         TabIndex        =   25
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Value"
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
         TabIndex        =   24
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick-up Charge"
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
         Left            =   2160
         TabIndex        =   23
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame fm_charges_HC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Household Charges"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      TabIndex        =   12
      Top             =   360
      Width           =   3855
      Begin VB.TextBox txt_pickup_HC 
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
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txt_purchase_HC 
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
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txt_delivery_HC 
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
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick-up Charge"
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
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Value"
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
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Charge"
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
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Height          =   1815
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   3855
      Begin VB.TextBox txt_discription 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discription:"
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
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lbl_charL 
         BackStyle       =   0  'Transparent
         Caption         =   "1/150"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc ado_additem 
      Height          =   375
      Left            =   960
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "ado_additem"
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
   Begin MSComctlLib.ImageList img_main 
      Left            =   0
      Top             =   4680
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
            Picture         =   "frm_additem.frx":7EA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_additem.frx":A483
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_additem.frx":CA5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_additem.frx":DB37
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_additem.frx":EC11
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_additem.frx":FCEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_additem.frx":122C5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      Begin VB.ComboBox cbo_pos 
         Height          =   315
         ItemData        =   "frm_additem.frx":1489F
         Left            =   2520
         List            =   "frm_additem.frx":148A9
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cbo_type 
         Height          =   315
         ItemData        =   "frm_additem.frx":148B6
         Left            =   240
         List            =   "frm_additem.frx":148D2
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txt_itemtitle 
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
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "POS Item"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Title"
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
         TabIndex        =   2
         Top             =   360
         Width           =   690
      End
   End
   Begin MSComctlLib.Toolbar tb_additem 
      Height          =   870
      Left            =   600
      TabIndex        =   7
      Top             =   4440
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1535
      ButtonWidth     =   2910
      ButtonHeight    =   1429
      TextAlignment   =   1
      ImageList       =   "img_main"
      DisabledImageList=   "img_main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save        "
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel     "
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "frm_additem.frx":1491C
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8400
      TabIndex        =   37
      Top             =   1150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2490
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QR Code"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8400
      TabIndex        =   35
      Top             =   360
      Width           =   795
   End
   Begin VB.Label lbl_addedit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADD ITEM"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   540
      Left            =   1320
      TabIndex        =   33
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frm_additem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QRGen As clsQRGen

Private Sub cbo_pos_Click()
If cbo_pos.text = "Yes" Then
    cbo_type.text = "Container"
    fm_charges_HC.Visible = True
    fm_charges_RC.Visible = True
    fm_charges_DC.Visible = True
    frm_additem.Width = 11085
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2


Else
    cbo_type.text = "Others"
    fm_charges_HC.Visible = False
    fm_charges_RC.Visible = False
    fm_charges_DC.Visible = False
    frm_additem.Width = 11085
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    
    'txt_delivery_HC.Text = "0"
    'txt_pickup_HC.Text = "0"
    'txt_purchase_HC.Text = "0"
    '
    'txt_delivery_RC.Text = "0"
    'txt_pickup_RC.Text = "0"
    'txt_purchase_RC.Text = "0"
    '
    'txt_delivery_DC.Text = "0"
    'txt_pickup_DC.Text = "0"
    'txt_purchase_DC.Text = "0"
    End If
    
    If cbo_pos.text = "No" Then
        txtQrCode.Visible = False
        Label14.Visible = False
    Else
        txtQrCode.Visible = True
        Label14.Visible = True
    End If
End Sub

Private Sub cbo_type_Click()
'/GROUP7 - not needed
'If cbo_type.ListIndex = 0 Or cbo_type.ListIndex = 1 Then
'cbo_pos.Text = "Yes"
'Else
'cbo_pos.Text = "No"
'End If

End Sub

Private Sub Form_Load()
Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
cbo_type.ListIndex = 0
cbo_pos.ListIndex = 0
Set QRGen = New clsQRGen
Combo1.AddItem "Black"
Combo1.AddItem "Red"
Combo1.AddItem "Green"
Combo1.AddItem "Yellow"
Combo1.AddItem "Blue"
Combo1.AddItem "Magenta"
Combo1.AddItem "Cyan"
Combo1.AddItem "Black"
End Sub

Private Sub Form_Terminate()
mdi_wrs.Enabled = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
mdi_wrs.Enabled = True
End Sub

Private Sub tb_additem_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rs As New ADODB.Recordset

If txt_delivery_HC.text = "" Then txt_delivery_HC.text = "0"
If txt_pickup_HC.text = "" Then txt_pickup_HC.text = "0"
If txt_purchase_HC.text = "" Then txt_purchase_HC.text = "0"

If txt_delivery_RC.text = "" Then txt_delivery_RC.text = "0"
If txt_pickup_RC.text = "" Then txt_pickup_RC.text = "0"
If txt_purchase_RC.text = "" Then txt_purchase_RC.text = "0"

If txt_delivery_DC.text = "" Then txt_delivery_DC.text = "0"
If txt_pickup_DC.text = "" Then txt_pickup_DC.text = "0"
If txt_purchase_DC.text = "" Then txt_purchase_DC.text = "0"

Select Case Button.Index
Case Is = 1
    'save
    If frm_additem.Caption = "ADD ITEM" Then
        If Trim(txt_itemtitle.text) = "" Then
            Call SetWindowPos(frm_additem.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            MsgBox "Input Item Title!", vbExclamation + vbOKOnly, "Water Refilling System"
            Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            Exit Sub
        End If
        Set rs = gvConnection.Execute("select * from tbl_item_list where qrCode =  '" & Nz(txtQrCode.text, -1) & "'")

        If rs.RecordCount > 0 Then
            Call SetWindowPos(frm_additem.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            MsgBox "QR Code is already on list!", vbExclamation + vbOKOnly, "Water Refilling System"
            Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            Exit Sub
        End If
        
        If txt_delivery_HC.text = "0" Or txt_pickup_HC.text = "0" Or txt_purchase_HC.text = "0" Then
            Call SetWindowPos(frm_additem.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            MsgBox "Delivery charge,Pick-up charge and Purchase value is 0", vbInformation
            Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        End If
        Call Conn_ado_additem
         ado_additem.RecordSource = "select * from tbl_item_list where Item_title =  '" & txt_itemtitle.text & "'"
         ado_additem.Refresh
         If ado_additem.Recordset.RecordCount = 0 Then
            ado_additem.Recordset.AddNew
            ado_additem.Recordset(0) = txt_itemtitle.text
            ado_additem.Recordset(1) = cbo_type.text
            ado_additem.Recordset(2) = cbo_pos.text
            
            ado_additem.Recordset(3) = Trim(txt_delivery_HC.text)
            ado_additem.Recordset(4) = Trim(txt_pickup_HC.text)
            ado_additem.Recordset(5) = Trim(txt_purchase_HC.text)
            
            ado_additem.Recordset(6) = Trim(txt_delivery_RC.text)
            ado_additem.Recordset(7) = Trim(txt_pickup_RC.text)
            ado_additem.Recordset(8) = Trim(txt_purchase_RC.text)
            
            ado_additem.Recordset(9) = Trim(txt_delivery_DC.text)
            ado_additem.Recordset(10) = Trim(txt_pickup_DC.text)
            ado_additem.Recordset(11) = Trim(txt_purchase_DC.text)
            
            ado_additem.Recordset(12) = txt_discription.text
            ado_additem.Recordset(16) = Nz(Trim(txtQrCode.text), RanmdonString)
            ado_additem.Recordset.Update
            ado_additem.Refresh
            
             '=====================
            Call Conn_ado_history
            frm_history.ado_history.Recordset.AddNew
            frm_history.ado_history.Recordset(0) = "Administrator"
            frm_history.ado_history.Recordset(1) = "9999"
            frm_history.ado_history.Recordset(2) = "N/A"
            frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
            frm_history.ado_history.Recordset(4) = "Added  " & txt_itemtitle.text & " to Itemlist"
            frm_history.ado_history.Recordset.Update
            frm_history.ado_history.Refresh
            '=====================
            
            Call SetWindowPos(frm_additem.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            MsgBox "Saved!", vbInformation + vbOKOnly, "Water Refilling System"
            Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            txt_itemtitle.text = ""
            
            txt_delivery_HC.text = ""
            txt_pickup_HC.text = ""
            txt_purchase_HC.text = ""
            
            txt_delivery_RC.text = ""
            txt_pickup_RC.text = ""
            txt_purchase_RC.text = ""
            
            txt_delivery_DC.text = ""
            txt_pickup_DC.text = ""
            txt_purchase_DC.text = ""
            
            txt_discription.text = ""
            mdi_wrs.Enabled = True
            Call Itemlist_ListLoad
            Unload Me
         Else
            Call SetWindowPos(frm_additem.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            MsgBox "Item already on list!", vbExclamation + vbOKOnly, "Water Refilling System"
            Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            Exit Sub
        End If
    Else ' EDIT ITEM
        Set rs = gvConnection.Execute("select * from tbl_item_list where qrCode =  '" & Nz(txtQrCode.text, -1) & "'  and Item_title != '" & txt_itemtitle.text & "'")
        If rs.RecordCount > 0 Then
            Call SetWindowPos(frm_additem.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            MsgBox "QR Code is already on list!", vbExclamation + vbOKOnly, "Water Refilling System"
            Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            Exit Sub
        End If
        Call Conn_ado_additem
        ado_additem.RecordSource = "select * from tbl_item_list where Item_title =  '" & txt_itemtitle.text & "'"
        ado_additem.Refresh
        ado_additem.Recordset(0) = txt_itemtitle.text
        ado_additem.Recordset(1) = cbo_type.text
        ado_additem.Recordset(2) = cbo_pos.text
        ado_additem.Recordset(3) = Trim(txt_delivery_HC.text)
        ado_additem.Recordset(4) = Trim(txt_pickup_HC.text)
        ado_additem.Recordset(5) = Trim(txt_purchase_HC.text)
        
        ado_additem.Recordset(6) = Trim(txt_delivery_RC.text)
        ado_additem.Recordset(7) = Trim(txt_pickup_RC.text)
        ado_additem.Recordset(8) = Trim(txt_purchase_RC.text)
        
        ado_additem.Recordset(9) = Trim(txt_delivery_DC.text)
        ado_additem.Recordset(10) = Trim(txt_pickup_DC.text)
        ado_additem.Recordset(11) = Trim(txt_purchase_DC.text)
        
        ado_additem.Recordset(16) = Nz(Trim(txtQrCode.text), RanmdonString)
        
        ado_additem.Recordset.Update
        ado_additem.Refresh
        
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = "Administrator"
        frm_history.ado_history.Recordset(1) = "9999"
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Edited   " & txt_itemtitle.text
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
        
        Call SetWindowPos(frm_additem.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        MsgBox "Update Saved!", vbInformation + vbOKOnly, "Water Refilling System"
        Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        txt_itemtitle.text = ""
        
        txt_delivery_HC.text = ""
        txt_pickup_HC.text = ""
        txt_purchase_HC.text = ""
        
        txt_delivery_RC.text = ""
        txt_pickup_RC.text = ""
        txt_purchase_RC.text = ""
        
        txt_delivery_DC.text = ""
        txt_pickup_DC.text = ""
        txt_purchase_DC.text = ""
        
        txt_discription.text = ""
        mdi_wrs.Enabled = True
        
        
        Call Itemlist_ListLoad
        Unload Me
    End If
    
    
Case Is = 2
'close
mdi_wrs.Enabled = True
Unload Me
End Select
Call frm_itemlist.lstview_itemlist_Click



End Sub

Private Sub txt_discription_Change()
lbl_charL.Caption = Len(txt_discription.text) & "/150"
If Len(txt_discription.text) >= 150 Then
Call SetWindowPos(frm_additem.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
MsgBox "Text Exceed to Limit!", vbExclamation + vbOKOnly, "Water Refilling System"
Call SetWindowPos(frm_additem.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
txt_discription.text = ""
End If

End Sub

Private Sub txt_delivery_HC_Change()
On Error Resume Next
        If IsNumeric(txt_delivery_HC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_delivery_HC.text = Trim(Left$(txt_delivery_HC.text, Len(txt_delivery_HC.text) - 1))
        End If
End Sub


Private Sub txt_pickup_HC_Change()
On Error Resume Next
        If IsNumeric(txt_pickup_HC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_pickup_HC.text = Trim(Left$(txt_pickup_HC.text, Len(txt_pickup_HC.text) - 1))
        End If

End Sub

Private Sub txt_purchase_HC_Change()
On Error Resume Next
        If IsNumeric(xt_purchase_HC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_purchase_HC.text = Trim(Left$(txt_purchase_HC.text, Len(txt_purchase_HC.text) - 1))
        End If

End Sub

Private Sub txt_delivery_RC_Change()
On Error Resume Next
        If IsNumeric(txt_delivery_RC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_delivery_RC.text = Trim(Left$(txt_delivery_RC.text, Len(txt_delivery_RC.text) - 1))
        End If
End Sub


Private Sub txt_pickup_RC_Change()
On Error Resume Next
        If IsNumeric(txt_pickup_RC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_pickup_RC.text = Trim(Left$(txt_pickup_RC.text, Len(txt_pickup_RC.text) - 1))
        End If

End Sub

Private Sub txt_purchase_RC_Change()
On Error Resume Next
        If IsNumeric(txt_purchase_RC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_purchase_RC.text = Trim(Left$(txt_purchase_RC.text, Len(txt_purchase_RC.text) - 1))
        End If

End Sub
Private Sub txt_delivery_DC_Change()
On Error Resume Next
        If IsNumeric(txt_delivery_RC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_delivery_DC.text = Trim(Left$(txt_delivery_DC.text, Len(txt_delivery_DC.text) - 1))
        End If
End Sub


Private Sub txt_pickup_DC_Change()
On Error Resume Next
        If IsNumeric(txt_pickup_DC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_pickup_DC.text = Trim(Left$(txt_pickup_DC.text, Len(txt_pickup_DC.text) - 1))
        End If

End Sub

Private Sub txt_purchase_DC_Change()
On Error Resume Next
        If IsNumeric(txt_purchase_DC.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_purchase_DC.text = Trim(Left$(txt_purchase_DC.text, Len(txt_purchase_DC.text) - 1))
        End If
    
End Sub

Private Sub txtQrCode_Change()
    Setup
    'Me.StrokeScribe1.Text = Me.txtQrCode.Text
    'Set frm_additem.Image1.Picture = QRGen.QRCodegenBarcode(Me.txtQrCode.Text)
    'Set frm_additem.Image1.Picture = QRGen.QRCodegenBarcode(txtQrCode.Text)
    Set frm_additem.Image1.Picture = QRGen.QRCodegenBarcode(txtQrCode.text)
    'Picture1. = QRCodegenConvertToData(QRCodegenBarcode(Me.txtQrCode.Text), 500, 500)
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
