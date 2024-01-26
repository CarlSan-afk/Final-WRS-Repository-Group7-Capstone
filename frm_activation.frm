VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_activation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activation"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_activation.frx":0000
   ScaleHeight     =   3795
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   240
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_activation.frx":3AEC8
      Height          =   2895
      Left            =   6840
      TabIndex        =   5
      Top             =   600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado_activation 
      Height          =   330
      Left            =   3360
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5895
      Begin VB.CommandButton cmd_trial 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Use Trial"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton cmd_activate 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Activate"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_activation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activation Code : "
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frm_activation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_activate_Click()
'Call Conn_ado_activation
If ado_activation.Recordset.RecordCount = 0 Then
ado_activation.Recordset.AddNew
ado_activation.Recordset(0) = GetMacAddress
ado_activation.Recordset(1) = FormatDateTime(Now, vbShortDate)
ado_activation.Recordset(2) = DateAdd("d", 7, Date)
ado_activation.Recordset(3) = "WRS-0528-2018-MAC-IES"
ado_activation.Recordset.Update
ado_activation.Refresh
Else
    If ado_activation.Recordset(3) = txt_activation.Text Then
    ado_activation.Recordset(4) = "WRS-0528-2018-MAC-IES"
    ado_activation.Recordset.Update
    ado_activation.Refresh
    MsgBox "Activated, Re-Run the Software", vbInformation, "WRS Management System"
    Unload Me
    Else
    MsgBox "Invalid Activation Code", vbExclamation, "WRS Management System"
    End If
End If

End Sub

Private Sub cmd_trial_Click()
'Call Conn_ado_activation
If ado_activation.Recordset.RecordCount = 1 Then Exit Sub
ado_activation.Recordset.AddNew
ado_activation.Recordset(0) = GetMacAddress
ado_activation.Recordset(1) = FormatDateTime(Now, vbShortDate)
ado_activation.Recordset(2) = DateAdd("d", 7, Date)
ado_activation.Recordset(3) = "WRS-0528-2018-MAC-IES"
ado_activation.Recordset.Update
ado_activation.Refresh
MsgBox "Trial Version Initialized, Re-Run the Software", vbinfromation, "WRS Management System"
Unload Me

End Sub


Private Sub Form_Activate()
If frm_activation.Caption = "Activate Full Version" Then
Else
    'Call Conn_ado_activation
    If ado_activation.Recordset.RecordCount = 0 Then
    Exit Sub
    Else
    frm_main.Show
    Unload Me
    End If
End If


End Sub


Private Sub Form_Load()
'Call Conn_ado_activation
'If ado_activation.Recordset.RecordCount = 0 Then
'ado_activation.Recordset.AddNew
'ado_activation.Recordset(0) = GetMacAddress
'ado_activation.Recordset(1) = FormatDateTime(Now, vbShortDate)
'ado_activation.Recordset(2) = DateAdd("d", 7, Date)
'ado_activation.Recordset(3) = "WRS-0528-2018-MAC-IES"
'ado_activation.Recordset.Update
'ado_activation.Refresh
'Else
'Exit Sub
'End If
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Timer1_Timer()


For i = 1 To 4230
Me.Height = i
Me.Top = (Screen.Height \ 2) - (i \ 2)
Next
Timer1.Enabled = False


End Sub
