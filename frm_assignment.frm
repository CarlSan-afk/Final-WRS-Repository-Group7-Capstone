VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_assignment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task Assignment"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6225
   Icon            =   "frm_assignment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_assignment.frx":10CA
   ScaleHeight     =   3870
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_resched 
      Caption         =   "Resched Next Day"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   2520
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "Delete Task"
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
      Left            =   3840
      TabIndex        =   6
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ComboBox cbo_status 
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
      ItemData        =   "frm_assignment.frx":6286
      Left            =   1560
      List            =   "frm_assignment.frx":6290
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "Add To Calendar"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txt_task 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1560
      TabIndex        =   1
      Top             =   510
      Width           =   4455
   End
   Begin MSAdodcLib.Adodc ado_task 
      Height          =   375
      Left            =   2760
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "ado_task"
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
   Begin VB.Label lbl_task 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   3360
      Width           =   555
   End
   Begin VB.Label Status 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status : "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Task :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   645
   End
End
Attribute VB_Name = "frm_assignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_add_Click()
If txt_task.text = "" Then Exit Sub
ado_task.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
If cmd_add.Caption = "Add to Calendar" Then
    Call Conn_ado_task
    frm_task.ado_task.Recordset.AddNew
    frm_task.ado_task.Recordset(1) = txt_task.text
    frm_task.ado_task.Recordset(2) = cbo_status.text
    frm_task.ado_task.Recordset(3) = frm_task.cal_task.Value
    frm_task.ado_task.Recordset.Update
    frm_task.ado_task.Refresh
    
    frm_task.Enabled = True
    Call Task_ListLoad
    Unload Me
Else
    frm_task.ado_task.RecordSource = "select * from tbl_task  where id = " & frm_task.lbl_id.Caption
    frm_task.ado_task.Refresh
    frm_task.ado_task.Recordset(1) = txt_task.text
    frm_task.ado_task.Recordset(2) = cbo_status.text
    frm_task.ado_task.Recordset(3) = frm_task.lbl_date.Caption
    frm_task.ado_task.Recordset.Update
    frm_task.ado_task.Refresh
    
    frm_task.Enabled = True
    Call Task_ListLoad
    Unload Me
End If

End Sub

Private Sub cmd_delete_Click()
frm_task.ado_task.RecordSource = "select * from tbl_task  where id = " & frm_task.lbl_id.Caption
frm_task.ado_task.Refresh
frm_task.ado_task.Recordset.Delete
frm_task.ado_task.Recordset.Update
frm_task.ado_task.Refresh
frm_task.Enabled = True
Call Task_ListLoad
Unload Me
End Sub

Private Sub cmd_resched_Click()
frm_task.ado_task.RecordSource = "select * from tbl_task  where id = " & frm_task.lbl_id.Caption
frm_task.ado_task.Refresh
frm_task.ado_task.Recordset(3) = DateAdd("d", 1, frm_task.ado_task.Recordset(3))
frm_task.ado_task.Recordset.Update
frm_task.ado_task.Refresh
frm_task.Enabled = True
Call Task_ListLoad
Unload Me
End Sub

Private Sub Form_Load()
cbo_status.ListIndex = 0
End Sub


Private Sub Form_Terminate()
frm_task.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_task.Enabled = True
End Sub

Private Sub Timer1_Timer()
For i = 1 To 3630
Me.Height = i
Me.Top = (Screen.Height \ 2) - (i \ 2)
Next
Timer1.Enabled = False

End Sub
