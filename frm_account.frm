VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_account 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Account"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7305
   Icon            =   "frm_account.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_account.frx":10CA
   ScaleHeight     =   5625
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fm_account 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Create Admin Account"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   6495
      Begin VB.TextBox txt_puser 
         Height          =   495
         Left            =   2880
         TabIndex        =   0
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txt_ppass 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txt_prpass 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "Save"
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3360
         TabIndex        =   4
         Top             =   3720
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc ado_account 
         Height          =   495
         Left            =   2400
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Enabled         =   0
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "ado_account"
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
      Begin VB.Line Line1 
         X1              =   480
         X2              =   5880
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
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
         Left            =   600
         TabIndex        =   8
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label lbl_password 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   600
         TabIndex        =   7
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repeat Password"
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
         Left            =   600
         TabIndex        =   6
         Top             =   2400
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frm_account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public username As String

Private Sub cmd_cancel_Click()
frm_login.Show
Unload Me
End Sub

Private Sub cmd_save_Click()
Call Conn_ado_account
If txt_puser.text = "" Or txt_ppass.text = "" Or txt_prpass.text = "" Then
    MsgBox " fill up all boxes", vbExclamation + vbOKOnly, "WRS Management System  "
    
    Exit Sub
    Else
    
    If cmd_save.Caption = "OK" Then
    
        'ado_account.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False"
        ado_account.ConnectionString = gvConnectString2
        ado_account.RecordSource = "select*from tbl_account where username =  '" & username & "'"
        ado_account.Refresh
        If ado_account.Recordset(0) = username Then
            If txt_ppass.text = txt_prpass.text Then
            Else
            MsgBox " Password did not Match!", vbExclamation + vbOKOnly, "WRS Management System  "
            Exit Sub
            End If
        Else
            MsgBox " Username does not exist!", vbExclamation + vbOKOnly, "WRS Management System  "
            Exit Sub
        End If
        
        frm_account.ado_account.Recordset.Delete
        frm_account.ado_account.Recordset.Update
        frm_account.ado_account.Refresh
        ado_account.Refresh
        ado_account.Recordset.AddNew
        ado_account.Recordset.Fields(0).Value = txt_puser.text
        ado_account.Recordset.Fields(1).Value = txt_ppass.text
        If fm_account.Caption = "Change Admin Password" Then
        ado_account.Recordset.Fields(3).Value = "admin"
        Else
        ado_account.Recordset.Fields(3).Value = "user"
        End If
        ado_account.Recordset.Update
        ado_account.Refresh
        ado_account.Refresh
               
        frm_login.Show
        Unload Me
        Exit Sub
    Else
    End If
    
        If txt_ppass.text = txt_prpass.text Then
            ado_account.RecordSource = "select* from tbl_account"
            ado_account.Refresh
            ado_account.Recordset.Find "username='" & txt_puser.text & "'"
        
            If Not ado_account.Recordset.EOF Then
            MsgBox " Username is already in use", vbExclamation + vbOKOnly, "WRS Management System  "
            Exit Sub
            Else
           
            End If
           
        Else
        MsgBox "Password did not match", vbExclamation + vbOKOnly, "WRS Management System  "
        Exit Sub
        End If
    End If

    If txt_puser.text = "" Or txt_ppass.text = "" Or txt_prpass.text = "" Then
    MsgBox " fill up all boxes", vbExclamation + vbOKOnly, "WRS Management System"
    Exit Sub
    Else
        
    End If
        ado_account.Refresh
        ado_account.Recordset.AddNew
        ado_account.Recordset.Fields(0).Value = txt_puser.text
        ado_account.Recordset.Fields(1).Value = txt_ppass.text
        If fm_account.Caption = "Create Admin Account" Then
        ado_account.Recordset.Fields(3).Value = "admin"
        Else
        ado_account.Recordset.Fields(3).Value = "user"
        End If
              
        ado_account.Recordset.Update
        ado_account.Refresh
        ado_account.Refresh
               
        frm_login.Show
        Unload Me
    
    MsgBox "Success!", vbExclamation + vbOKOnly, " & mdi_wrs.Caption "
    
    frm_login.Show
    Unload Me
End Sub



Private Sub Form_Load()
Call Conn_ado_account
If ado_account.Recordset.RecordCount = 0 Then
cmd_save.Caption = "Save"
Else
cmd_save.Caption = "OK"
End If

End Sub
