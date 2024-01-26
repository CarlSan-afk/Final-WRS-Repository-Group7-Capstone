VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frm_deliveryman 
   BorderStyle     =   0  'None
   Caption         =   "Deliveryman"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_deliveryman.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc ado_deliveryman 
      Height          =   570
      Left            =   6480
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1005
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
      Left            =   480
      TabIndex        =   4
      Top             =   8280
      Width           =   5175
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
      Left            =   480
      TabIndex        =   3
      Top             =   7320
      Width           =   5175
   End
   Begin VB.TextBox txt_completename 
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
      Left            =   480
      TabIndex        =   2
      Top             =   6240
      Width           =   5175
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   14295
      Begin MSComctlLib.ListView lstview_deliveryman 
         Height          =   4215
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   7435
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Deliveryman"
            Object.Width           =   8405
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Contact Number"
            Object.Width           =   8405
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Address"
            Object.Width           =   8405
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar_main 
      Height          =   870
      Left            =   7080
      TabIndex        =   1
      Top             =   7320
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   1535
      ButtonWidth     =   3016
      ButtonHeight    =   1429
      TextAlignment   =   1
      ImageList       =   "img_main"
      DisabledImageList=   "img_main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save     "
            ImageIndex      =   9
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New  "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit        "
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete      "
            ImageIndex      =   12
         EndProperty
      EndProperty
      MouseIcon       =   "frm_deliveryman.frx":DBEB
   End
   Begin MSComctlLib.ImageList img_main 
      Left            =   13560
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":1D045
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":1F61F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":21BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":241D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":267AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":27887
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":28961
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":29A3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":2C015
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":2E5EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":30BC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_deliveryman.frx":331A3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl_addedit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERYMAN LIST"
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
      Left            =   480
      TabIndex        =   9
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label3 
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
      Left            =   480
      TabIndex        =   7
      Top             =   7920
      Width           =   1110
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
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
      Left            =   465
      TabIndex        =   6
      Top             =   6840
      Width           =   1830
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complete Name"
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
      Left            =   480
      TabIndex        =   5
      Top             =   5880
      Width           =   1800
   End
End
Attribute VB_Name = "frm_deliveryman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selection As Boolean

Private Sub Form_Activate()
Call Deliveryman_ListLoad
End Sub

Private Sub Form_Load()
Call Deliveryman_ListLoad
Select Case mdi_wrs.WindowState
Case Is = vbNormal
mdi_wrs.Width = frm_deliveryman.Width + 315
mdi_wrs.Height = frm_deliveryman.Height + 955
DoCenter frm_deliveryman, mdi_wrs
Case Is = vbMaximized
DoCenter frm_deliveryman, mdi_wrs
Case Is = vbMinimized
Exit Sub
End Select
Call lstview_deliveryman_Click
End Sub



Private Sub lstview_deliveryman_Click()

On Error Resume Next
txt_completename.text = lstview_deliveryman.SelectedItem.text
txt_contact.text = lstview_deliveryman.SelectedItem.SubItems(1)
txt_address.text = lstview_deliveryman.SelectedItem.SubItems(2)

txt_completename.Enabled = False
txt_contact.Enabled = False
txt_address.Enabled = False

Call Conn_ado_deliveryman
ado_deliveryman.RecordSource = "select * from tbl_deliveryman where Deliveryman_name =  '" & txt_completename.text & "'"
ado_deliveryman.Refresh

End Sub

Public Sub Toolbar_main_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case Is = 1
'save
    If txt_completename.text = "" And txt_contact.text = "" And txt_address.text = "" Then
    answer = MsgBox("Pls Complete Details!", vbExclamation + vbOKCancel, "Water Refilling System")
        If answer = vbOK Then
        Exit Sub
        Else
        ado_deliveryman.Refresh
        txt_completename.text = ado_deliveryman.Recordset(0)
        txt_contact.text = ado_deliveryman.Recordset(1)
        txt_address.text = ado_deliveryman.Recordset(2)
        txt_completename.Enabled = False
        txt_contact.Enabled = False
        txt_address.Enabled = False
        Toolbar_main.Buttons(1).Enabled = False
        Toolbar_main.Buttons(2).Enabled = True
        Toolbar_main.Buttons(3).Enabled = True
        Toolbar_main.Buttons(4).Enabled = True
        lstview_deliveryman.Enabled = True
        Exit Sub
        End If
    End If
   '-----------------
    
If selection = True Then
        Call Conn_ado_deliveryman
        ado_deliveryman.RecordSource = "select * from tbl_deliveryman where Deliveryman_name =  '" & txt_completename.text & "'"
        ado_deliveryman.Refresh
            
            If ado_deliveryman.Recordset.RecordCount = 0 Then
            Else
            MsgBox "Name Already Exist!", vbExclamation + vbOKOnly, "Water Refilling System"
            txt_completename.text = ""
            txt_contact.text = ""
            txt_address.text = ""
            Exit Sub
            End If
        ado_deliveryman.Recordset.AddNew
        ado_deliveryman.Recordset(0) = txt_completename.text
        ado_deliveryman.Recordset(1) = txt_contact.text
        ado_deliveryman.Recordset(2) = txt_address.text
        ado_deliveryman.Recordset.Update
        ado_deliveryman.Refresh
        MsgBox "Saved!", vbInformation + vbOKOnly, "Water Refilling System"
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = "Administrator"
        frm_history.ado_history.Recordset(1) = 9999
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Added New Deliveryman - " & txt_completename.text
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
        txt_completename.Enabled = False
        txt_contact.Enabled = False
        txt_address.Enabled = False
        Call Deliveryman_ListLoad
        Toolbar_main.Buttons(1).Enabled = False
        Toolbar_main.Buttons(2).Enabled = True
        Toolbar_main.Buttons(3).Enabled = True
        Toolbar_main.Buttons(4).Enabled = True
        lstview_deliveryman.Enabled = True
        Call lstview_deliveryman_Click
Else
        ado_deliveryman.Recordset(0) = txt_completename.text
        ado_deliveryman.Recordset(1) = txt_contact.text
        ado_deliveryman.Recordset(2) = txt_address.text
        ado_deliveryman.Recordset.Update
        ado_deliveryman.Refresh
        MsgBox "Saved!", vbInformation + vbOKOnly, "Water Refilling System"
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = "Administrator"
        frm_history.ado_history.Recordset(1) = 9999
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Edited  Deliveryman Information"
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================
        txt_completename.Enabled = False
        txt_contact.Enabled = False
        txt_address.Enabled = False
        Call Deliveryman_ListLoad
        Toolbar_main.Buttons(1).Enabled = False
        Toolbar_main.Buttons(2).Enabled = True
        Toolbar_main.Buttons(3).Enabled = True
        Toolbar_main.Buttons(4).Enabled = True
        lstview_deliveryman.Enabled = True
        Call lstview_deliveryman_Click

End If


Case Is = 2
'add new
selection = True
txt_completename.Enabled = True
txt_contact.Enabled = True
txt_address.Enabled = True
txt_completename.text = ""
txt_contact.text = ""
txt_address.text = ""
Toolbar_main.Buttons(1).Enabled = True
Toolbar_main.Buttons(2).Enabled = False
Toolbar_main.Buttons(3).Enabled = False
Toolbar_main.Buttons(4).Enabled = False
lstview_deliveryman.Enabled = False
txt_completename.SetFocus

Case Is = 3
'edit
selection = False
If txt_completename.text = "N/A" Then Exit Sub
txt_completename.Enabled = True
txt_contact.Enabled = True
txt_address.Enabled = True
Toolbar_main.Buttons(1).Enabled = True
Toolbar_main.Buttons(2).Enabled = False
Toolbar_main.Buttons(3).Enabled = False
Toolbar_main.Buttons(4).Enabled = False
lstview_deliveryman.Enabled = False

Case Is = 4
'delete
On Error Resume Next
If txt_completename.text = "N/A" Then Exit Sub
ado_deliveryman.Recordset.Delete
ado_deliveryman.Recordset.Update
ado_deliveryman.Refresh
Call Conn_ado_deliveryman
Call Deliveryman_ListLoad
Call lstview_deliveryman_Click
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = "Administrator"
        frm_history.ado_history.Recordset(1) = 9999
        frm_history.ado_history.Recordset(2) = "N/A"
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Deleted Deliveryman"
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        '=====================


End Select

End Sub

Private Sub txt_address_Change()
If txt_address.text = "" Then
txt_address.BackColor = &H80FF80
Else
txt_address.BackColor = &HFFFFFF
End If
End Sub

Private Sub txt_completename_Change()
If txt_completename.text = "" Then
txt_completename.BackColor = &H80FF80
Else
txt_completename.BackColor = &HFFFFFF
End If
End Sub

Private Sub txt_contact_Change()
If txt_contact.text = "" Then
txt_contact.BackColor = &H80FF80
Else
txt_contact.BackColor = &HFFFFFF
End If
End Sub
