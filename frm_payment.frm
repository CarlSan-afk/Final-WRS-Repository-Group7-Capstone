VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_payment 
   Caption         =   "Pay Order"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   Icon            =   "frm_payment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_payment.frx":10CA
   ScaleHeight     =   4215
   ScaleWidth      =   4425
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc ado_payment 
      Height          =   375
      Left            =   5040
      Top             =   3600
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "ado_payment"
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
   Begin VB.CommandButton cmd_proceed 
      BackColor       =   &H00FFFF80&
      Caption         =   "Proceed"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmd_cancel 
      BackColor       =   &H00FFFF80&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txt_change 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2400
      TabIndex        =   7
      Text            =   "0"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txt_payment 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2400
      TabIndex        =   1
      Text            =   "0"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txt_amounttopay 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2400
      TabIndex        =   6
      Text            =   "0"
      Top             =   360
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc ado_sales 
      Height          =   375
      Left            =   5040
      Top             =   3960
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
   Begin VB.Label lbl_sold_item 
      Caption         =   "sold item"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lbl_deliveryman 
      Caption         =   "deliveryman"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lbl_id_number 
      Caption         =   "id number"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lbl_classification 
      Caption         =   "classification"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lbl_customer_name 
      Caption         =   "customer name"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lbl_amount 
      Caption         =   "amount"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lbl_change 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment : "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount to Pay : "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2145
   End
End
Attribute VB_Name = "frm_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_cancel_Click()
If frm_payment.Caption = "Pay Order" Then
'Pay order
frm_order.Enabled = True
Unload Me
Else
'Pay Balance
mdi_wrs.Enabled = True
Unload Me
End If
End Sub

Private Sub cmd_proceed_Click()
Dim refilledgallon As Integer, mvchange As Currency, mvAmountTopay As Currency, mvAmountTendered As Currency
refilledgallon = Val(frm_order.lbl_gallon.Caption)
mvchange = Nz(Val(txt_payment.text), 0) - Nz(Val(txt_amounttopay.text), 0)
mvAmountTopay = Nz(Val(txt_payment.text), 0) - Nz(Val(Me.txt_change.text), 0)
mvAmountTendered = Nz(Val(txt_payment.text), 0)
'mvchange = Nz(Val(txt_payment.Text), 0) - Nz(Val(txt_amounttopay.Text), 0)
If frm_payment.Caption = "Pay Order" Then
' Pay Order============================================================

    If Nz(Val(txt_payment.text), 0) > Nz(Val(txt_amounttopay.text), 0) Then
    mvchange = Nz(Val(txt_payment.text), 0) - Nz(Val(txt_amounttopay.text), 0)
    txt_payment.text = txt_amounttopay.text
    End If
    
    If Val(txt_change.text) < 0 Then
     If frm_main.txt_customername = "Guest" Then
        MsgBox "The transaction cannot be completed as Guess does not offer credit services.", vbInformation
        Exit Sub
     End If
    answer = MsgBox("PAYMENT  is not enough! Do you want to PROCEED?", vbExclamation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        Call frm_order.Toolbar_main_ButtonClick(frm_order.Toolbar_main.Buttons(2))
        Else
        Exit Sub
        End If
     End If
     updateQTY
     Call frm_order.Toolbar_main_ButtonClick(frm_order.Toolbar_main.Buttons(2))
     
     Call Conn_ado_refilledMDI
     mdi_wrs.ado_refilled.Recordset.AddNew
     mdi_wrs.ado_refilled.Recordset(0) = refilledgallon
     mdi_wrs.ado_refilled.Recordset(1) = FormatDateTime(Now, vbShortDate)
     mdi_wrs.ado_refilled.Recordset(2) = frm_main.txt_idnumber.text
     mdi_wrs.ado_refilled.Recordset(3) = frm_main.txt_customername.text
     mdi_wrs.ado_refilled.Recordset.Update
     mdi_wrs.ado_refilled.Refresh
     mdi_wrs.ado_refilled.Visible = False
     Call Status_Sales_Refilled
             
     
 ElseIf frm_payment.Caption = "Pay Balance" Then
 'Pay Balance============================================================
    If txt_payment.text = "0" Or txt_payment.text = "" Then
    Unload frm_payment
    Exit Sub
    End If
    
    If Val(txt_payment.text) > Val(txt_amounttopay.text) Then
    txt_payment.text = txt_amounttopay.text
    End If
    
    If Val(txt_change.text) < 0 Then
        answer = MsgBox("PAYMENT  is not enough! Do you want to PROCEED?", vbExclamation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        ' code to pay
        
        Conn_ado_delivery
        frm_main.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
        frm_main.ado_delivery.Refresh
        frm_main.ado_delivery.Recordset(0) = txt_change.text
        frm_main.ado_delivery.Recordset(1) = Val(frm_main.ado_delivery.Recordset(1)) - Val(txt_payment.text)
        frm_main.ado_delivery.Recordset.Update
        frm_main.ado_delivery.Visible = False
        frm_main.ado_delivery.Refresh
        '---------------
        Call Conn_ado_sales
        If txt_payment.text = "" Or txt_payment.text = 0 Then Exit Sub
        frm_sales.ado_sales.Recordset.AddNew
        frm_sales.ado_sales.Recordset(0) = txt_payment.text
        frm_sales.ado_sales.Recordset(1) = frm_main.txt_customername
        frm_sales.ado_sales.Recordset(2) = frm_main.txt_classification
        frm_sales.ado_sales.Recordset(3) = frm_main.txt_idnumber
        frm_sales.ado_sales.Recordset(4) = FormatDateTime(Now, vbShortDate)
        frm_main.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
        frm_main.ado_delivery.Refresh
        frm_sales.ado_sales.Recordset(5) = frm_main.ado_delivery.Recordset(6)
        frm_sales.ado_sales.Recordset(6) = "Pay Credit - " & frm_main.ado_delivery.Recordset(7)
        frm_sales.ado_sales.Recordset.Update
        frm_sales.ado_sales.Refresh
        frm_sales.ado_sales.Visible = False
        
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = frm_main.txt_customername
        frm_history.ado_history.Recordset(1) = frm_main.txt_idnumber
        frm_history.ado_history.Recordset(2) = frm_main.txt_classification
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        'frm_history.ado_history.Recordset(4) = "Pay Credit - " & txt_payment.text & " " & "Change~" & (mvchange) & " " & frm_main.ado_delivery.Recordset(7)
        frm_history.ado_history.Recordset(4) = "Pay Credit - " & Nz(mvAmountTopay, txt_payment.text) & " /" & "Amount Tendered - " & mvAmountTendered & "/ " & "Changed" & (mvchange) & " " & frm_main.ado_delivery.Recordset(7)
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        frm_history.ado_history.Visible = False
        '=====================
        Call Conn_ado_delivery
        Call Delivery_ListLoad
        Call Status_Sales_Refilled
        'Call Conn_ado_sales
        'Call Sales_ListLoad
        mdi_wrs.Enabled = True
        Call Status_Sales_Refilled
        Unload Me
        Exit Sub
    Else
        Exit Sub
        End If
    End If
     ' code to pay
         Conn_ado_delivery
        frm_main.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
        frm_main.ado_delivery.Refresh
        frm_main.ado_delivery.Recordset(0) = txt_change.text
        frm_main.ado_delivery.Recordset(1) = Val(frm_main.ado_delivery.Recordset(1)) - Val(txt_payment.text)
        frm_main.ado_delivery.Recordset.Update
        frm_main.ado_delivery.Refresh
        frm_main.ado_delivery.Visible = False
        '---------------
        Call Conn_ado_sales
        If txt_payment.text = "" Or txt_payment.text = 0 Then Exit Sub
        frm_sales.ado_sales.Recordset.AddNew
        frm_sales.ado_sales.Recordset(0) = txt_payment.text
        frm_sales.ado_sales.Recordset(1) = frm_main.txt_customername
        frm_sales.ado_sales.Recordset(2) = frm_main.txt_classification
        frm_sales.ado_sales.Recordset(3) = frm_main.txt_idnumber
        frm_sales.ado_sales.Recordset(4) = FormatDateTime(Now, vbShortDate)
        frm_main.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
        frm_main.ado_delivery.Refresh
        frm_sales.ado_sales.Recordset(5) = frm_main.ado_delivery.Recordset(6)
        frm_sales.ado_sales.Recordset(6) = "Pay Credit" & frm_main.ado_delivery.Recordset(7)
        frm_sales.ado_sales.Recordset.Update
        frm_sales.ado_sales.Refresh
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = frm_main.txt_customername
        frm_history.ado_history.Recordset(1) = frm_main.txt_idnumber
        frm_history.ado_history.Recordset(2) = frm_main.txt_classification
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        '"Pay Credit- " & txt_payment.Text & " " & frm_main.ado_delivery.Recordset(7)
        frm_history.ado_history.Recordset(4) = "Pay Credit - " & Nz(mvAmountTopay, txt_payment.text) & " / " & "Changed - " & (mvchange) & " / " & "Amount Tendered -" & mvAmountTendered & "Change~" & frm_main.ado_delivery.Recordset(7)
        'frm_history.ado_history.Recordset(4) = "Pay Credit- " & txt_payment.Text & " " & frm_main.ado_delivery.Recordset(7)
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        
        

        '=====================
        Call Conn_ado_delivery
        Call Delivery_ListLoad
        Call Status_Sales_Refilled
        'Call Conn_ado_sales
        'Call Sales_ListLoad
        Call Status_Sales_Refilled
        mdi_wrs.Enabled = True
        Unload Me
        
        frm_sales.ado_sales.Visible = False
        frm_main.ado_delivery.Visible = False
        frm_sales.ado_sales.Visible = False
        frm_history.ado_history.Visible = False
 ElseIf frm_payment.Caption = "Pay Credit" Then
 'Pay Credit============================================================
    If txt_payment.text = "" Or txt_payment.text = "0" Then Exit Sub
    Call Conn_ado_salesPAYMENT
    ado_sales.RecordSource = "select * from tbl_sales where Id_number like '%" & lbl_id_number.Caption & "%'"
    ado_sales.Refresh
    ado_sales.Recordset.AddNew
    ado_sales.Recordset(0) = txt_payment.text
    ado_sales.Recordset(1) = lbl_customer_name.Caption
    ado_sales.Recordset(2) = lbl_classification.Caption
    ado_sales.Recordset(3) = lbl_id_number.Caption
    ado_sales.Recordset(4) = FormatDateTime(Now, vbShortDate)
    ado_sales.Recordset(5) = lbl_deliveryman.Caption
    ado_sales.Recordset(6) = lbl_sold_item.Caption
    ado_sales.Recordset.Update
    ado_sales.Refresh
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = lbl_customer_name.Caption
        frm_history.ado_history.Recordset(1) = lbl_id_number.Caption
        frm_history.ado_history.Recordset(2) = lbl_classification.Caption
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        'frm_history.ado_history.Recordset(4) = lbl_sold_item.Caption & "/Pay Credit " & mvAmountTopay
        frm_history.ado_history.Recordset(4) = "Pay Credit - " & Nz(mvAmountTopay, txt_payment.text) & " /  " & "Amount Tendered - " & mvAmountTendered & " / " & "Changed - " & (mvchange) & " " & lbl_sold_item.Caption
        '"Pay Credit - " & Nz(mvAmountTopay, txt_payment.text) & " /  " & "Amount Tendered - " & mvAmountTendered & " / " & "Balance - " & (mvchange) & " " & frm_Deliver.ado_delivery.Recordset(7)
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        
        '=====================
    
    Call Conn_ado_payment
    ado_payment.RecordSource = "select * from tbl_credit where Id_number like '%" & lbl_id_number.Caption & "%'"
    ado_payment.Refresh
    
    If txt_change.text >= 0 Then
        ado_payment.Recordset.Delete
        ado_payment.Recordset.Update
        ado_payment.Refresh
    Else
        ado_payment.Recordset(0) = txt_change.text
        ado_payment.Recordset.Update
        ado_payment.Refresh
    End If
    Call Sales_ListLoad
    Call Credit_ListLoad
    mdi_wrs.Enabled = True
    
    
    Unload Me
    
    Else
    'Pay Balance. =================================================================================
    If txt_payment.text = "0" Or txt_payment.text = "" Then
    Unload frm_payment
    Exit Sub
    End If
    If Val(txt_payment.text) > Val(txt_amounttopay.text) Then
    txt_payment.text = txt_amounttopay.text
    End If
    
    If Val(txt_change.text) < 0 Then
        answer = MsgBox("PAYMENT  is not enough! Do you want to PROCEED?", vbExclamation + vbOKCancel, "Confirm")
        If answer = vbOK Then
        ' code to pay
        
        Conn_ado_FORdelivery
        '/GROUP7-2023/11/22-error yung idnumber
        frm_Deliver.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_Deliver.lstview_order.SelectedItem.SubItems(4) & "%'"
        frm_Deliver.ado_delivery.Refresh
        frm_Deliver.ado_delivery.Recordset(0) = Val(txt_change.text)
        frm_Deliver.ado_delivery.Recordset(1) = Val(frm_Deliver.ado_delivery.Recordset(1)) - Val(txt_payment.text)
        frm_Deliver.ado_delivery.Recordset.Update
        frm_Deliver.ado_delivery.Visible = False
        frm_Deliver.ado_delivery.Refresh
        
        '---------------
        Call Conn_ado_sales
        If txt_payment.text = "" Or txt_payment.text = 0 Then Exit Sub
        '/GROUP7-2023/11/22-error yung idnumber + plus 1 subitem
        frm_sales.ado_sales.Recordset.AddNew
        frm_sales.ado_sales.Recordset(0) = txt_payment.text
        frm_sales.ado_sales.Recordset(1) = frm_Deliver.lstview_order.SelectedItem.SubItems(2)
        frm_sales.ado_sales.Recordset(2) = frm_Deliver.lstview_order.SelectedItem.SubItems(3)
        '/GROUP7-2023/11/22-error yung idnumber
        frm_sales.ado_sales.Recordset(3) = frm_Deliver.lstview_order.SelectedItem.SubItems(4)
        frm_sales.ado_sales.Recordset(4) = FormatDateTime(Now, vbShortDate)
        '/GROUP7-2023/11/22-error yung idnumber
        frm_Deliver.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_Deliver.lstview_order.SelectedItem.SubItems(4) & "%'"
        frm_Deliver.ado_delivery.Refresh
        frm_sales.ado_sales.Recordset(5) = frm_Deliver.ado_delivery.Recordset(6)
        frm_sales.ado_sales.Recordset(6) = "Pay Credit - " & frm_Deliver.ado_delivery.Recordset(7)
        frm_sales.ado_sales.Recordset.Update
        frm_sales.ado_sales.Refresh
         '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = frm_Deliver.lstview_order.SelectedItem.SubItems(2)
        frm_history.ado_history.Recordset(1) = frm_Deliver.lstview_order.SelectedItem.SubItems(4)
        frm_history.ado_history.Recordset(2) = frm_Deliver.lstview_order.SelectedItem.SubItems(3)
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Pay Credit - " & txt_payment.text & " " & frm_Deliver.ado_delivery.Recordset(7)
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        
        '=====================
        Call Conn_ado_FORdelivery
        Call FORDelivery_ListLoad
        Call Conn_ado_sales
        Call Sales_ListLoad
        Call Credit_ListLoad
        Call Status_Sales_Refilled
        frm_sales.ado_sales.Visible = False
        frm_main.ado_delivery.Visible = False
        frm_sales.ado_sales.Visible = False
        frm_history.ado_history.Visible = False
        mdi_wrs.Enabled = True
        Unload Me
        Exit Sub
    Else
        Exit Sub
        End If
    End If
     ' code to pay
        Conn_ado_FORdelivery
        frm_Deliver.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_Deliver.lstview_order.SelectedItem.SubItems(4) & "%'"
        frm_Deliver.ado_delivery.Refresh
        frm_Deliver.ado_delivery.Recordset(0) = CInt(txt_change.text)
        frm_Deliver.ado_delivery.Recordset(1) = Val(frm_Deliver.ado_delivery.Recordset(1)) - Val(txt_payment.text)
        frm_Deliver.ado_delivery.Recordset.Update
        frm_Deliver.ado_delivery.Visible = False
        frm_Deliver.ado_delivery.Refresh
        '---------------
        Call Conn_ado_sales
        If txt_payment.text = "" Or txt_payment.text = 0 Then Exit Sub
        frm_sales.ado_sales.Recordset.AddNew
        frm_sales.ado_sales.Recordset(0) = txt_payment.text
        '/GROUP7 fixed wrong customername
        frm_sales.ado_sales.Recordset(1) = frm_Deliver.lstview_order.SelectedItem.SubItems(2)
        frm_sales.ado_sales.Recordset(2) = frm_Deliver.lstview_order.SelectedItem.SubItems(3)
        frm_sales.ado_sales.Recordset(3) = frm_Deliver.lstview_order.SelectedItem.SubItems(4)
        frm_sales.ado_sales.Recordset(4) = FormatDateTime(Now, vbShortDate)
        frm_Deliver.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_Deliver.lstview_order.SelectedItem.SubItems(4) & "%'"
        frm_Deliver.ado_delivery.Refresh
        frm_sales.ado_sales.Recordset(5) = frm_Deliver.ado_delivery.Recordset(6)
        frm_sales.ado_sales.Recordset(6) = "Pay Credit- " & frm_Deliver.ado_delivery.Recordset(7)
        frm_sales.ado_sales.Recordset.Update
        frm_sales.ado_sales.Refresh
        '=====================
        Call Conn_ado_history
        frm_history.ado_history.Recordset.AddNew
        frm_history.ado_history.Recordset(0) = frm_Deliver.lstview_order.SelectedItem.SubItems(2)
        frm_history.ado_history.Recordset(1) = frm_Deliver.lstview_order.SelectedItem.SubItems(4)
        frm_history.ado_history.Recordset(2) = frm_Deliver.lstview_order.SelectedItem.SubItems(3)
        frm_history.ado_history.Recordset(3) = FormatDateTime(Now, vbGeneralDate)
        frm_history.ado_history.Recordset(4) = "Pay Credit - " & Nz(mvAmountTopay, txt_payment.text) & " /  " & "Amount Tendered - " & mvAmountTendered & " / " & "Changed - " & (mvchange) & " " & frm_Deliver.ado_delivery.Recordset(7)
        frm_history.ado_history.Recordset.Update
        frm_history.ado_history.Refresh
        
        '=====================
        Call Conn_ado_FORdelivery
        Call FORDelivery_ListLoad
        Call Conn_ado_sales
        Call Sales_ListLoad
        Call Credit_ListLoad
        Call Status_Sales_Refilled
        mdi_wrs.Enabled = True
        Unload Me
        frm_sales.ado_sales.Visible = False
        frm_main.ado_delivery.Visible = False
        frm_sales.ado_sales.Visible = False
        frm_history.ado_history.Visible = False
        frm_credit.Visible = False
 End If
        frm_sales.ado_sales.Visible = False
        frm_main.ado_delivery.Visible = False
        frm_sales.ado_sales.Visible = False
        frm_history.ado_history.Visible = False
        frm_credit.Visible = False
        frm_history.Visible = False
        frm_sales.Visible = False
End Sub

Private Sub Form_Activate()

txt_change.text = Val(txt_payment.text) - Val(txt_amounttopay.text)
txt_payment.text = ""
txt_payment.SetFocus
End Sub

Private Sub Form_Terminate()
If frm_payment.Caption = "Pay Order" Then
'Pay order
frm_order.Enabled = True
Else
'Pay Balance
mdi_wrs.Enabled = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If frm_payment.Caption = "Pay Order" Then
'Pay Order
frm_order.Enabled = True
Else
'Pay Balance
mdi_wrs.Enabled = True
End If

End Sub

Private Sub txt_change_Change()
Select Case Val(txt_change.text)
Case Is = 0
lbl_change.Caption = "Exact Amount"
Case Is < 0
lbl_change.Caption = "Balance"
Case Is > 0
lbl_change.Caption = "Change"
End Select
End Sub

Private Sub txt_payment_Change()
On Error Resume Next
        If IsNumeric(txt_payment.text) Then
        Else
        'MsgBox "Invalid Input!", vbExclamation + vbOKOnly, "Water Refilling System"
        txt_payment.text = Trim(Left$(txt_payment.text, Len(txt_payment.text) - 1))
        End If
txt_change.text = Val(txt_payment.text) - Val(txt_amounttopay.text)
End Sub

Private Sub txt_payment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmd_proceed_Click
ElseIf KeyAscii = 32 Then
Call cmd_proceed_Click
End If
End Sub
