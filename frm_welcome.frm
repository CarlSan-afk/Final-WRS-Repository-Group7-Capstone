VERSION 5.00
Begin VB.Form frm_welcome 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11595
   ClipControls    =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frm_welcome.frx":0000
   ScaleHeight     =   4590
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblUser 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "test"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   2280
      TabIndex        =   0
      Top             =   3360
      Width           =   8775
   End
End
Attribute VB_Name = "frm_welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
