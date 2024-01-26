VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_borrower 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_borrower.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc ado_borrower 
      Height          =   495
      Left            =   11640
      Top             =   8760
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
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
      Caption         =   "ado_borrower"
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
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   600
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
         TabIndex        =   4
         Top             =   120
         Width           =   1905
      End
   End
   Begin MSComctlLib.ListView lstview_borrower 
      Height          =   7335
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12938
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Number"
         Object.Width           =   4958
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer Name"
         Object.Width           =   4958
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Borrowed Item"
         Object.Width           =   4958
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   4958
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Borrowed"
         Object.Width           =   4958
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Count :"
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
      Left            =   480
      TabIndex        =   6
      Top             =   8760
      Width           =   870
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
      Left            =   1440
      TabIndex        =   5
      Top             =   8760
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BORROWER LIST"
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
      Left            =   9960
      TabIndex        =   1
      Top             =   480
      Width           =   3510
   End
End
Attribute VB_Name = "frm_borrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
lbl_count.Caption = lstview_borrower.ListItems.Count
End Sub

Private Sub Form_Load()
Call Borrower_ListLoad
Select Case mdi_wrs.WindowState
Case Is = vbNormal
mdi_wrs.Width = frm_CI.Width + 315
mdi_wrs.Height = frm_CI.Height + 955
DoCenter frm_borrower, mdi_wrs
Case Is = vbMaximized
DoCenter frm_borrower, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select
End Sub



Private Sub lbl_profit_Click()

End Sub

Private Sub lstview_borrower_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstview_borrower.SortKey = ColumnHeader.Index - 1

    If lstview_borrower.SortKey = 1 Then lstview_borrower.SortKey = 3 ' **** This column changes the key
    lstview_borrower.SortOrder = (lstview_borrower.SortOrder - 1) * -1
    lstview_borrower.Sorted = True
End Sub

Private Sub txt_search_Change()
Call Borrower_ListLoad
lbl_count.Caption = lstview_borrower.ListItems.Count
End Sub
