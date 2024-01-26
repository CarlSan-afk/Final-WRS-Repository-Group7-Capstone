VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.MDIForm mdi_wrs 
   BackColor       =   &H00FFFFC0&
   Caption         =   "WRS Management System"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   15195
   Icon            =   "mdi_wrs.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdi_wrs.frx":10CA
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   5055
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   15135
      TabIndex        =   1
      Top             =   2250
      Visible         =   0   'False
      Width           =   15195
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   2820
         Left            =   11160
         ScaleHeight     =   2820
         ScaleWidth      =   3015
         TabIndex        =   3
         Top             =   960
         Width           =   3015
      End
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   16200
         Left            =   0
         Picture         =   "mdi_wrs.frx":1C750
         ScaleHeight     =   16200
         ScaleWidth      =   28800
         TabIndex        =   2
         Top             =   0
         Width           =   28800
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   8040
   End
   Begin MSAdodcLib.Adodc ado_sales 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
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
   Begin MSAdodcLib.Adodc ado_ci 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
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
   Begin MSComctlLib.StatusBar status_mdi 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            TextSave        =   "01/26/2024"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   3360
            Picture         =   "mdi_wrs.frx":37DD6
            Text            =   "No of Customer"
            TextSave        =   "No of Customer"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   4683
            Picture         =   "mdi_wrs.frx":3A3B0
            Text            =   "Refilled Gallon for the Day"
            TextSave        =   "Refilled Gallon for the Day"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   3572
            Picture         =   "mdi_wrs.frx":3C98A
            Text            =   "Sales for the Day"
            TextSave        =   "Sales for the Day"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   9181
            Text            =   "Created by: MC Abarca FB Account : https://www.facebook.com/mackee2018/"
            TextSave        =   "Created by: MC Abarca FB Account : https://www.facebook.com/mackee2018/"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Text            =   "Login as: "
            TextSave        =   "Login as: "
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Text            =   "Logout"
            TextSave        =   "Logout"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado_itemlist 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   1125
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
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
   Begin MSAdodcLib.Adodc ado_deliveryman 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   375
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
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
   Begin MSAdodcLib.Adodc ado_refilled 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   1875
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
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
      Caption         =   "ado_refilled"
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
   Begin MSAdodcLib.Adodc ado_activation 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
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
      Caption         =   "ado_activation"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu pos 
      Caption         =   "POS"
   End
   Begin VB.Menu customer 
      Caption         =   "Customer"
      Begin VB.Menu customerlist 
         Caption         =   "Customer List"
         Shortcut        =   ^L
      End
      Begin VB.Menu CI 
         Caption         =   "Customer's Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu borrower 
         Caption         =   "Borrower"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu delivery 
      Caption         =   "Delivery"
   End
   Begin VB.Menu APX 
      Caption         =   "Admin Panel"
      Begin VB.Menu cash 
         Caption         =   "Cash"
         Begin VB.Menu credit 
            Caption         =   "Credit"
         End
         Begin VB.Menu sales 
            Caption         =   "Sales and Expenses Log"
            Shortcut        =   ^M
         End
         Begin VB.Menu SAE 
            Caption         =   "Stocks and Expenses"
         End
      End
      Begin VB.Menu deliveryman 
         Caption         =   "Deliveryman"
      End
      Begin VB.Menu cua 
         Caption         =   "Create User Account"
      End
      Begin VB.Menu backup 
         Caption         =   "BackUp Database"
      End
   End
   Begin VB.Menu history 
      Caption         =   "History"
   End
   Begin VB.Menu taskcalendar 
      Caption         =   "Task Calendar"
   End
End
Attribute VB_Name = "mdi_wrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type




Private Sub activation_Click()
'If activation.Caption = "Activated" Then Exit Sub
'frm_activation.Show
'frm_activation.Caption = "Activate Full Version"
'frm_activation.cmd_trial.Visible = False
'Unload Me
End Sub

Private Sub backup_Click()
'AppPath As Variant
On Error GoTo ErrorTrap
Dim sFilename, sTemp1 As String, mvfiletitle As String, command As String
Dim FileName As String
Dim FileOpen As Variant, mysqldumpPath As String
'mysqldumpPath = "C:\Program Files (x86)\MySQL\MySQL Server 5.5\bin\mysql.exe"
'    With CommonDialog1
'        .Filter = "Export Files (*.sql)"
'        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
'        .DialogTitle = "File Save"
'        .FileName = sFilename
'        .CancelError = True
'        .ShowSave
'        On Error Resume Next
'    End With
'    sFilename = CommonDialog1.FileName
'    mvfiletitle = CommonDialog1.FileTitle
'    If mvfiletitle = "" Then Exit Sub
'    If Dir(sFilename) <> "" Then Kill sFilename
'    Set FileOpen = CreateObject("WSCript.shell")
'    If mvfiletitle <> "" Then
'        'FileOpen.run "cmd /C:\Program Files\MySQL\MySQL Server 5.1\bin\mysqldump" & " -u" & HostUser & " -p" & HostPassword & " -h" & HostServer & " " & HostDatabase & " > C:\" & mvfiletitle & ".sql", 0, True
'        'command
'        FileOpen.Run "cmd /c mysqldump -u" & HostUser & " -p" & HostPassword & " -h" & HostServer & " " & HostDatabase & " > C:\" & mvfiletitle & ".sql", 0, True
'    Else
'        MsgBox "Unable to backup database please enter file name", vbExclamation
'    End If
'    'Move
'    FileCopy "C:\" & mvfiletitle & ".sql", sFilename & ".sql"
'    Kill "C:\" & mvfiletitle & ".sql"
'    'Message
    'mvfiletitle = mvfiletitle & "db.sql"
    'sFilename = sFilename & ".sql"
    'AppPath = GetShortPath(App.path)
    Select Case Format(Now, "ddd")
        Case "Mon", "Wed", "Fri": FileName = "MWF"
        Case "Sun": FileName = "Sunday"
        Case Else: FileName = "TTHS"
    End Select
    FileName = FileName & "_" & "wrs" & "db.sql"
    Open AppendPath(AppPath, "\mysqlbackup.bat") For Output As #119
    Print #119, "echo off"
    Print #119, Replace("cd '" & AppPath & "'", "'", """")
    Print #119, "cls"
    Print #119, "mysqldump.exe -u root -p1r1ppl310 wrs > " & AppendPath(AppPath, "\" & FileName)
    'Print #119, "mysqldump.exe -u root -p1r1ppl310 wrs > " & sFilename
    Print #119, "break>" & AppendPath(AppPath, "\FIN.txt")
    Close #119
    
    Shell AppendPath("""" & AppPath, "\mysqlbackup.bat" & """")
    If Dir(AppendPath(AppPath, "\" & FileName)) <> "" Then MsgBox "The backup of database" & FileName & "completed successfully.", vbInformation
ErrorTrap:
    Exit Sub
End Sub

Private Sub borrower_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    frm.Visible = False
  End If
Next
frm_borrower.Show
Call Status_Sales_Refilled

End Sub

Private Sub CI_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    frm.Visible = False
  End If
Next
frm_CI.Show
frm_CI.txt_search.text = frm_main.txt_customername.text
Call Status_Sales_Refilled

End Sub

Private Sub credit_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    frm.Visible = False
    'Unload frm
  End If
Next
frm_credit.Show
Call Status_Sales_Refilled

End Sub

Private Sub cua_Click()
frm_account.Show
frm_account.fm_account.Caption = "Create User Account"
frm_account.cmd_save.Caption = "Save"
Unload Me

End Sub

Private Sub customerlist_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    frm.Visible = False
  End If
Next
frm_customer.Show
Call Status_Sales_Refilled

End Sub

Private Sub delivery_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    frm.Visible = False
  End If
Next
frm_Deliver.Show
Call Status_Sales_Refilled
End Sub

Private Sub deliveryman_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
   frm.Visible = False
  End If
Next
frm_deliveryman.Show
Call Status_Sales_Refilled
End Sub



Private Sub history_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    frm.Visible = False
  End If
Next
frm_history.Show
Call Status_Sales_Refilled
End Sub


Private Sub MDIForm_Activate()
MDIForm_Resize
Call Status_Sales_Refilled
Call Conn_MDI
If ado_itemlist.Recordset.RecordCount = 0 Then
ado_itemlist.Recordset.AddNew
ado_itemlist.Recordset(0) = "Slim Container with Faucet"
ado_itemlist.Recordset(1) = "Container"
ado_itemlist.Recordset(2) = "Yes"
ado_itemlist.Recordset(3) = "25.00"
ado_itemlist.Recordset(4) = "20.00"
ado_itemlist.Recordset(5) = "200.00"

ado_itemlist.Recordset(6) = "25.00"
ado_itemlist.Recordset(7) = "20.00"
ado_itemlist.Recordset(8) = "200.00"

ado_itemlist.Recordset(9) = "20.00"
ado_itemlist.Recordset(10) = "20.00"
ado_itemlist.Recordset(11) = "200.00"

ado_itemlist.Recordset(12) = "Default Value"
ado_itemlist.Recordset(13) = "0"
ado_itemlist.Recordset(14) = "0"
ado_itemlist.Recordset.Update
ado_itemlist.Refresh
Else
End If

If ado_deliveryman.Recordset.RecordCount = 0 Then
ado_deliveryman.Recordset.AddNew
ado_deliveryman.Recordset(0) = "N/A"
ado_deliveryman.Recordset(1) = "N/A"
ado_deliveryman.Recordset(2) = "N/A"
ado_deliveryman.Recordset.Update
ado_deliveryman.Refresh
Else
End If


Call Status_Sales_Refilled

End Sub




Private Sub MDIForm_Resize()
closeform
On Error Resume Next
Dim client_rect As RECT
Dim client_hwnd As Long

    picStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight


    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight

    ' Set the MDI form's picture.
    Picture = picStretched.Image

    ' Invalidate the picture.
    client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1
 
'Dim frm As Form
'For Each frm In Forms
'    '/GROUP7-2023 - keep unloading forms enable resizing of current frm
'  If frm.Name <> Me.Name And frm.Name <> "frm_login" Then
'    Unload frm
'    frm.Show
'  End If
'Next
Call Status_Sales_Refilled



End Sub

Private Sub pos_Click()
'Unload frm_main
closeform
If frm_deliveryman.Toolbar_main.Buttons(1).Enabled = False Then
Else
    Call frm_deliveryman.Toolbar_main_ButtonClick(frm_deliveryman.Toolbar_main.Buttons(1))
    Exit Sub
End If

If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
    Else
    Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
    Exit Sub
End If
Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    'frm.Visible = False
    Unload frm
  End If
Next

frm_main.Show
Call Conn_MDI
Call Status_Sales_Refilled


End Sub

Private Sub SAE_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
    Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
    Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    Unload frm
  End If
Next
Unload frm_itemlist
frm_itemlist.Show
Call Status_Sales_Refilled
End Sub

Private Sub sales_Click()
'Unload frm_main
closeform
If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
Else
Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
Exit Sub
End If

Dim frm As Form
For Each frm In Forms
  If frm.Name <> Me.Name Then
    frm.Visible = False
  End If
Next
frm_sales.Show
Call Status_Sales_Refilled
End Sub


Private Sub MDIForm_Load()

Call Conn_MDI
Call GetMacAddress
'/GROUP7 removed FUNCTION


End Sub



Private Sub status_mdi_PanelClick(ByVal Panel As MSComctlLib.Panel)
Dim frm As Form
Dim xpanel As Variant
xpanel = Left(Panel, 4)

Select Case xpanel
    Case "Logo":
        frm_login.Show
        mdi_wrs.Hide
    Case "Numb":
        closeform
        If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
        Else
        Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
        Exit Sub
        End If
        
        
        For Each frm In Forms
          If frm.Name <> Me.Name Then
            frm.Visible = False
          End If
        Next
        frm_customer.Show
        Call Status_Sales_Refilled
        
    Case "Sale":
        closeform
        If frm_CI.tb_mainbutton.Buttons(1).Enabled = False Then
        Else
        Call frm_CI.tb_mainbutton_ButtonClick(frm_CI.tb_mainbutton.Buttons(1))
        Exit Sub
        End If
        
        'Dim frm As Form
        For Each frm In Forms
          If frm.Name <> Me.Name Then
            frm.Visible = False
          End If
        Next
        frm_sales.Show
        Call Status_Sales_Refilled
End Select

'If Left(Panel, 4) = "Logo" Then
'    frm_login.Show
'    mdi_wrs.Hide
'ElseIf Left(Panel, 4) Then
'End If
'frm_login.Show
'mdi_wrs.Hide
End Sub

'Private Sub status_mdi_PanelClick(ByVal Panel As MSComctlLib.Panel)


'Dim myelement As Variant
'myelement = status_mdi(1)
'status_mdi.Panels


'If status_mdi.Index(1) = 1 Then
'    MsgBox ""
'End If
'frm_login.Show
'mdi_wrs.Hide

'End Sub

'Private Sub status_mdi_PanelClick(ByVal Panel As MSComctlLib.Panel)
'Select Case Panel.Index
'Case Is = 5
'On Error Resume Next
'Dim sURL As String

'sURL = " https://www.facebook.com/mackee2018/ "
'Shell "C:\Program Files\Google\Chrome\Application\chrome.exe" & sURL, vbMaximizedFocus
'Call SetWindowPos(mdi_wrs.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'End Select


'End Sub

Private Sub taskcalendar_Click()
frm_task.Show
End Sub

Private Sub Timer1_Timer()
    status_mdi.Panels(5).text = Mid(status_mdi.Panels(5).text, 2) & Left(status_mdi.Panels(5).text, 1)
End Sub
