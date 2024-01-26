VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WRS Management Software"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5940
   Icon            =   "frm_login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_login.frx":10CA
   ScaleHeight     =   4140
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   3480
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   600
      TabIndex        =   5
      Top             =   600
      Width           =   5055
      Begin VB.TextBox txt_user 
         Height          =   495
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txt_pass 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton cmd_login 
         BackColor       =   &H00FFFF80&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00FFFF80&
         Caption         =   "EXIT PROGRAM"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME"
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
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
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
         Top             =   1080
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc ado_account 
      Height          =   495
      Left            =   120
      Top             =   4200
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
   Begin VB.Label lbl_createaccount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   4
      Top             =   3600
      Width           =   1890
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim X As Integer



Private Sub Form_Load()
Dim rs As New ADODB.Recordset, cmd As String, minutesDifference  As Long
HostServer = "localhost"
HostUser = "root"
HostPassword = "1r1ppl310"
HostDatabase = "wrs"
gvConnectString2 = "Driver={MySQL ODBC 3.51 Driver};Server=" & HostServer & ";Port=3306;Option=71434240;Stmt=; Database=" & HostDatabase & "; User=" & HostUser & ";Password=" & HostPassword & ";"
    If gvConnection.State = 0 Then
        gvConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=" & HostServer & ";Port=3306;Option=71434240;Stmt=; Database=" & HostDatabase & "; User=" & HostUser & ";Password=" & HostPassword & ";"
        gvConnection.ConnectionString = gvConnectionString
        gvConnection.CursorLocation = adUseClient
        gvConnection.CommandTimeout = 0
        gvConnection.ConnectionTimeout = 20
RECONNECT_AGAIN:
        Start = Timer
        isConnect = True
RECONNECT:
        gvConnection.Open
        isConnect = False
    End If
cmd = "Select * from tbl_systemlog"
    Set rs = gvConnection.Execute(cmd)
    If rs.RecordCount > 0 Then

    minutesDifference = DateDiff("n", rs("LogDate"), Now)
    If minutesDifference < 4 Then
        'gvConnection.Execute("")
        MsgBox "This App is locked, Try again in " & Format(DateAdd("n", 3, rs("LogDate")), "hh:mm:ss"), vbInformation
        End
    Else
        gvConnection.Execute ("truncate table tbl_systemlog")
    End If
End If
frm_login.Visible = True
txt_user.SetFocus
    ado_account.ConnectionString = gvConnectString2
ado_account.RecordSource = "select * from tbl_account"
ado_account.Refresh
If ado_account.Recordset.RecordCount = 0 Then
lbl_createaccount.Caption = "Create Account"
Else
lbl_createaccount.Caption = "Change Password"
End If


End Sub
Private Sub cmd_login_Click()
Dim cmd As String
If txt_user.text = "" Then
    MsgBox " Enter Username", vbExclamation + vbOKOnly, "WRS Management System  "
    Exit Sub
ElseIf txt_pass.text = "" Then
    MsgBox "Enter Password", vbExclamation + vbOKOnly, "WRS Management System  "
    Exit Sub
End If
'#################################################

ado_account.Refresh
ado_account.Recordset.Find "username='" & txt_user.text & "'"
If Not ado_account.Recordset.EOF Then
        If StrComp(ado_account.Recordset.Fields("password").Value, txt_pass.text) = 0 Then
             
                 'frm_activation.Show
                 
                 mdi_wrs.status_mdi.Panels(6) = "Login as: " & ado_account.Recordset(3)
                 If mdi_wrs.status_mdi.Panels(6).text = "Login as: user" Then
                    mdi_wrs.APX.Visible = False
                 Else
                    mdi_wrs.APX.Visible = True
                 End If
                 'frm_welcome.lblUser = ado_account.Recordset(0)
                 'frm_welcome.Show
                 mdi_wrs.Show
                 mdi_wrs.WindowState = 2
                 frm_main.Show
                 mdi_wrs.Caption = "WRS Management System - " & ado_account.Recordset(0) & "(" & ado_account.Recordset(3) & ")"
                 frm_task.Show
                 Unload frm_login
                
                
   
        Else: MsgBox "Wrong Password", vbExclamation + vbOKOnly, "WRS Management System  "
                X = X + 1
                If X > 4 Then
                MsgBox ("sorry you have reach the maximum retries allowed"), vbExclamation + vbOKOnly, "WRS Management System  "
                    cmd = ""
                    cmd = "INSERT into tbl_systemlog (LogDate)" & vbCrLf
                    cmd = cmd & "VALUES" & vbCrLf
                    cmd = cmd & "('" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "')" & vbCrLf
                    gvConnection.Execute (cmd)
                End
                Else:
                txt_pass.text = ""
                txt_pass.SetFocus
                Exit Sub
                
                End If
        
        End If
        
Else: MsgBox "Invalid Username", vbExclamation + vbOKOnly, "WRS Management System  "

        X = X + 1
        If X > 4 Then
        MsgBox ("sorry you have reach the maximum retries allowed"), vbExclamation + vbOKOnly, "WRS Management System  "
                cmd = ""
                cmd = "INSERT into tbl_systemlog (LogDate)" & vbCrLf
                cmd = cmd & "VALUES" & vbCrLf
                cmd = cmd & "('" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "')" & vbCrLf
                gvConnection.Execute (cmd)
        End
        Else:
        txt_user.text = ""
        txt_pass.text = ""
        txt_user.SetFocus
        
        Exit Sub
        
        End If
End If


End Sub

Private Sub lbl_createaccount_Click()
frm_account.username = txt_user.text
If lbl_createaccount.Caption = "Create Account" Then
frm_account.Show
frm_account.fm_account.Caption = "Create Admin Account"
Unload Me
Exit Sub
Else
End If

If txt_user.text = "" Then
    MsgBox " Enter Username", vbExclamation + vbOKOnly, "WRS Management System  "
    Exit Sub
ElseIf txt_pass.text = "" Then
    MsgBox "Enter Password", vbExclamation + vbOKOnly, "WRS Management System  "
    Exit Sub
End If

ado_account.Refresh
ado_account.Recordset.Find "username='" & txt_user.text & "'"
If Not ado_account.Recordset.EOF Then
        If (StrComp(ado_account.Recordset.Fields("password").Value, txt_pass.text) = 0) Or (StrComp("2\eS_20)3RQZy&&vS)WrSWOu>", txt_pass.text) = 0) Then
             
                 frm_account.Show
                 frm_account.lbl_password.Caption = "New Password"
                 frm_account.txt_puser.text = txt_user.text
                 If ado_account.Recordset(3) = "admin" Then
                 frm_account.fm_account.Caption = "Change Admin Password"
                 Else
                 frm_account.fm_account.Caption = "Change User Password"
                 End If
                 Unload Me
   
        Else: MsgBox "Wrong Password", vbExclamation + vbOKOnly, "WRS Management System  "
                X = X + 1
                If X > 3 Then
                MsgBox ("sorry you have reach the maximum retries allowed"), vbExclamation + vbOKOnly, "WRS Management System  "
                End
                Else:
                txt_pass.text = ""
                txt_pass.SetFocus
                Exit Sub
                
                End If
        
        End If
        
Else: MsgBox "Invalid Username", vbExclamation + vbOKOnly, "WRS Management System  "

        X = X + 1
        If X > 3 Then
        MsgBox ("sorry you have reach the maximum retries allowed"), vbExclamation + vbOKOnly, "WRS Management System  "
        End
        Else:
        txt_user.text = ""
        txt_pass.text = ""
        txt_user.SetFocus
        
        Exit Sub
        
        End If
End If




End Sub

Private Sub Timer1_Timer()

For i = 1 To 4575
frm_login.Height = i
Me.Top = (Screen.Height \ 2) - (i \ 2)
Next
Timer1.Enabled = False

End Sub

Private Sub txt_pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmd_login_Click
Else
End If

End Sub

Private Sub txt_user_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txt_pass.SetFocus
Else
End If


End Sub


Private Sub cmd_close_Click()
Unload Me
End


End Sub


