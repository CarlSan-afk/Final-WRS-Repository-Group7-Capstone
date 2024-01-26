Attribute VB_Name = "Listview_control"
Option Explicit
Dim list As ListItem


Sub CI_ListLoad()


Select Case frm_CI.cbo_search.text
Case Is = "All"

    Call Conn_ado_ci
    frm_CI.lstview_ci.ListItems.Clear
    
    Do Until frm_CI.ado_ci.Recordset.EOF
    Set list = frm_CI.lstview_ci.ListItems.Add(, , frm_CI.ado_ci.Recordset(0) & "")
    
    list.SubItems(1) = frm_CI.ado_ci.Recordset(1) & ""
    list.SubItems(2) = frm_CI.ado_ci.Recordset(2) & ""
    list.SubItems(3) = frm_CI.ado_ci.Recordset(3) & ""
    list.SubItems(4) = frm_CI.ado_ci.Recordset(4) & ""
    list.SubItems(5) = frm_CI.ado_ci.Recordset(5) & ""
    list.SubItems(6) = frm_CI.ado_ci.Recordset(6) & ""
    list.SubItems(7) = frm_CI.ado_ci.Recordset(7) & ""
    list.SubItems(8) = frm_CI.ado_ci.Recordset(8) & ""
    frm_CI.ado_ci.Recordset.MoveNext
        Loop
    
    With frm_CI.lstview_ci
    
    gvCmdCustomer = frm_CI.ado_ci.RecordSource
    
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
    
Case Is = "ID Number"
    frm_CI.ado_ci.RecordSource = "select * from tbl_customer_info where ID_number like  '%" & frm_CI.txt_search.text & "%' order by date_of_last_buy desc"
    frm_CI.ado_ci.Refresh
    frm_CI.lstview_ci.ListItems.Clear
    
    Do Until frm_CI.ado_ci.Recordset.EOF
    Set list = frm_CI.lstview_ci.ListItems.Add(, , frm_CI.ado_ci.Recordset(0) & "")
    
    list.SubItems(1) = frm_CI.ado_ci.Recordset(1) & ""
    list.SubItems(2) = frm_CI.ado_ci.Recordset(2) & ""
    list.SubItems(3) = frm_CI.ado_ci.Recordset(3) & ""
    list.SubItems(4) = frm_CI.ado_ci.Recordset(4) & ""
    list.SubItems(5) = frm_CI.ado_ci.Recordset(5) & ""
    list.SubItems(6) = frm_CI.ado_ci.Recordset(6) & ""
    list.SubItems(7) = frm_CI.ado_ci.Recordset(7) & ""
    list.SubItems(8) = frm_CI.ado_ci.Recordset(8) & ""
    frm_CI.ado_ci.Recordset.MoveNext
        Loop
    
    With frm_CI.lstview_ci
    gvCmdCustomer = frm_CI.ado_ci.RecordSource
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
 
Case Is = "Customer Name"
    Call Conn_ado_ci
    frm_CI.ado_ci.RecordSource = "select * from tbl_customer_info where Customer_Name like  '%" & frm_CI.txt_search.text & "%'order by date_of_last_buy desc"
    frm_CI.ado_ci.Refresh
    frm_CI.lstview_ci.ListItems.Clear
    
    Do Until frm_CI.ado_ci.Recordset.EOF
    Set list = frm_CI.lstview_ci.ListItems.Add(, , frm_CI.ado_ci.Recordset(0) & "")
    
    list.SubItems(1) = frm_CI.ado_ci.Recordset(1) & ""
    list.SubItems(2) = frm_CI.ado_ci.Recordset(2) & ""
    list.SubItems(3) = frm_CI.ado_ci.Recordset(3) & ""
    list.SubItems(4) = frm_CI.ado_ci.Recordset(4) & ""
    list.SubItems(5) = frm_CI.ado_ci.Recordset(5) & ""
    list.SubItems(6) = frm_CI.ado_ci.Recordset(6) & ""
    list.SubItems(7) = frm_CI.ado_ci.Recordset(7) & ""
    list.SubItems(8) = frm_CI.ado_ci.Recordset(8) & ""
    frm_CI.ado_ci.Recordset.MoveNext
        Loop
    
    With frm_CI.lstview_ci
    gvCmdCustomer = frm_CI.ado_ci.RecordSource
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
    
Case Is = "Classification"
    frm_CI.ado_ci.RecordSource = "select * from tbl_customer_info where Classification like  '%" & frm_CI.txt_search.text & "%' order by date_of_last_buy desc"
    frm_CI.ado_ci.Refresh
    frm_CI.lstview_ci.ListItems.Clear
    
    Do Until frm_CI.ado_ci.Recordset.EOF
    Set list = frm_CI.lstview_ci.ListItems.Add(, , frm_CI.ado_ci.Recordset(0) & "")
    
    list.SubItems(1) = frm_CI.ado_ci.Recordset(1) & ""
    list.SubItems(2) = frm_CI.ado_ci.Recordset(2) & ""
    list.SubItems(3) = frm_CI.ado_ci.Recordset(3) & ""
    list.SubItems(4) = frm_CI.ado_ci.Recordset(4) & ""
    list.SubItems(5) = frm_CI.ado_ci.Recordset(5) & ""
    list.SubItems(6) = frm_CI.ado_ci.Recordset(6) & ""
    list.SubItems(7) = frm_CI.ado_ci.Recordset(7) & ""
    list.SubItems(8) = frm_CI.ado_ci.Recordset(8) & ""
    frm_CI.ado_ci.Recordset.MoveNext
        Loop
    
    With frm_CI.lstview_ci
    gvCmdCustomer = frm_CI.ado_ci.RecordSource
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
    
Case Is = "Address"
    frm_CI.ado_ci.RecordSource = "select * from tbl_customer_info where Address like  '%" & frm_CI.txt_search.text & "%' order by date_of_last_buy desc"
    frm_CI.ado_ci.Refresh
    frm_CI.lstview_ci.ListItems.Clear
    
    Do Until frm_CI.ado_ci.Recordset.EOF
    Set list = frm_CI.lstview_ci.ListItems.Add(, , frm_CI.ado_ci.Recordset(0) & "")
    
    list.SubItems(1) = frm_CI.ado_ci.Recordset(1) & ""
    list.SubItems(2) = frm_CI.ado_ci.Recordset(2) & ""
    list.SubItems(3) = frm_CI.ado_ci.Recordset(3) & ""
    list.SubItems(4) = frm_CI.ado_ci.Recordset(4) & ""
    list.SubItems(5) = frm_CI.ado_ci.Recordset(5) & ""
    list.SubItems(6) = frm_CI.ado_ci.Recordset(6) & ""
    list.SubItems(7) = frm_CI.ado_ci.Recordset(7) & ""
    list.SubItems(8) = frm_CI.ado_ci.Recordset(8) & ""
    frm_CI.ado_ci.Recordset.MoveNext
        Loop
    
    With frm_CI.lstview_ci
    gvCmdCustomer = frm_CI.ado_ci.RecordSource
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
    
End Select
End Sub

Sub Itemlist_ListLoad()
Call Conn_ado_itemlist
    frm_itemlist.lstview_itemlist.ListItems.Clear
    
    Do Until frm_itemlist.ado_itemlist.Recordset.EOF
    Set list = frm_itemlist.lstview_itemlist.ListItems.Add(, , frm_itemlist.ado_itemlist.Recordset(0) & "")
    
    list.SubItems(1) = frm_itemlist.ado_itemlist.Recordset(1) & ""
    list.SubItems(2) = frm_itemlist.ado_itemlist.Recordset(2) & ""
        
    frm_itemlist.ado_itemlist.Recordset.MoveNext
        Loop
    
    With frm_itemlist.lstview_itemlist
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub

Sub Customer_ListLoad()
Select Case frm_customer.cbo_search.text

Case Is = "All"
    Call Conn_ado_customer
    frm_customer.lstview_customer.ListItems.Clear
    
    Do Until frm_customer.ado_customer.Recordset.EOF
    Set list = frm_customer.lstview_customer.ListItems.Add(, , frm_customer.ado_customer.Recordset(0) & "")
    
    list.SubItems(1) = frm_customer.ado_customer.Recordset(1) & ""
    list.SubItems(2) = frm_customer.ado_customer.Recordset(2) & ""
    list.SubItems(3) = frm_customer.ado_customer.Recordset(5) & ""
        
    frm_customer.ado_customer.Recordset.MoveNext
        Loop
    
    With frm_customer.lstview_customer
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
Case Is = "Customer"
    Call Conn_ado_customer
    frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where Customer_Name like  '%" & frm_customer.txt_search.text & "%'order by date_of_last_buy desc"
    frm_customer.ado_customer.Refresh
    frm_customer.lstview_customer.ListItems.Clear
    Do Until frm_customer.ado_customer.Recordset.EOF
    Set list = frm_customer.lstview_customer.ListItems.Add(, , frm_customer.ado_customer.Recordset(0) & "")
    
    list.SubItems(1) = frm_customer.ado_customer.Recordset(1) & ""
    list.SubItems(2) = frm_customer.ado_customer.Recordset(2) & ""
    list.SubItems(3) = frm_customer.ado_customer.Recordset(5) & ""
    list.SubItems(4) = frm_customer.ado_customer.Recordset(3) & ""
        
    frm_customer.ado_customer.Recordset.MoveNext
        Loop
    
    With frm_customer.lstview_customer
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
Case Is = "Classification"
    Call Conn_ado_customer
    frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where Classification like  '%" & frm_customer.txt_search.text & "%'order by date_of_last_buy desc"
    frm_customer.ado_customer.Refresh
    frm_customer.lstview_customer.ListItems.Clear
    Do Until frm_customer.ado_customer.Recordset.EOF
    Set list = frm_customer.lstview_customer.ListItems.Add(, , frm_customer.ado_customer.Recordset(0) & "")
    
    list.SubItems(1) = frm_customer.ado_customer.Recordset(1) & ""
    list.SubItems(2) = frm_customer.ado_customer.Recordset(2) & ""
    list.SubItems(3) = frm_customer.ado_customer.Recordset(5) & ""
    list.SubItems(4) = frm_customer.ado_customer.Recordset(3) & ""
        
    frm_customer.ado_customer.Recordset.MoveNext
        Loop
    
    With frm_customer.lstview_customer
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
Case Is = "Address"
    Call Conn_ado_customer
    frm_customer.ado_customer.RecordSource = "select * from tbl_customer_info where Address like  '%" & frm_customer.txt_search.text & "%'order by date_of_last_buy desc"
    frm_customer.ado_customer.Refresh
    frm_customer.lstview_customer.ListItems.Clear
    Do Until frm_customer.ado_customer.Recordset.EOF
    Set list = frm_customer.lstview_customer.ListItems.Add(, , frm_customer.ado_customer.Recordset(0) & "")
    
    list.SubItems(1) = frm_customer.ado_customer.Recordset(1) & ""
    list.SubItems(2) = frm_customer.ado_customer.Recordset(2) & ""
    list.SubItems(3) = frm_customer.ado_customer.Recordset(5) & ""
    list.SubItems(4) = frm_customer.ado_customer.Recordset(3) & ""
        
    frm_customer.ado_customer.Recordset.MoveNext
        Loop
    
    With frm_customer.lstview_customer
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Select
End Sub

Sub CustomerItem_ListLoad()
Call Conn_ado_customeritem
    frm_main.ado_customeritem.RecordSource = "select * from tbl_customer_item where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
    frm_main.ado_customeritem.Refresh
    
    frm_main.lstview_ci.ListItems.Clear
    
    Do Until frm_main.ado_customeritem.Recordset.EOF
    Set list = frm_main.lstview_ci.ListItems.Add(, , frm_main.ado_customeritem.Recordset(2) & "")
    
    list.SubItems(1) = frm_main.ado_customeritem.Recordset(3) & ""
    frm_main.ado_customeritem.Recordset.MoveNext
    Loop
    
    With frm_main.lstview_ci
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub

Sub Delivery_ListLoad()
Call Conn_ado_delivery
    frm_main.ado_delivery.RecordSource = "select * from tbl_delivery where ID_number like  '%" & frm_main.txt_idnumber.text & "%'"
    frm_main.ado_delivery.Refresh
    frm_main.lstview_order.ListItems.Clear
    
    Do Until frm_main.ado_delivery.Recordset.EOF
'    Set list = frm_main.lstview_order.ListItems.Add(, , frm_main.ado_delivery.Recordset(0) & "")
'
'    list.SubItems(1) = frm_main.ado_delivery.Recordset(2) & ""
'    list.SubItems(2) = frm_main.ado_delivery.Recordset(3) & ""
'    list.SubItems(3) = frm_main.ado_delivery.Recordset(4) & ""
'    list.SubItems(4) = frm_main.ado_delivery.Recordset(5) & ""
'    list.SubItems(5) = frm_main.ado_delivery.Recordset(6) & ""
'    list.SubItems(6) = frm_main.ado_delivery.Recordset(7) & ""
    Set list = frm_main.lstview_order.ListItems.Add(, , frm_main.ado_delivery.Recordset(8) & "")
    
    list.SubItems(1) = frm_main.ado_delivery.Recordset(0) & ""
    list.SubItems(2) = frm_main.ado_delivery.Recordset(2) & ""
    list.SubItems(3) = frm_main.ado_delivery.Recordset(3) & ""
    list.SubItems(4) = frm_main.ado_delivery.Recordset(4) & ""
    list.SubItems(5) = frm_main.ado_delivery.Recordset(5) & ""
    list.SubItems(6) = frm_main.ado_delivery.Recordset(6) & ""
    list.SubItems(7) = frm_main.ado_delivery.Recordset(7) & ""
          
    frm_main.ado_delivery.Recordset.MoveNext
        Loop
    
    With frm_main.lstview_order
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub

Sub CustomerItemCI_ListLoad()
Call Conn_ado_customeritemCI
    frm_CI.ado_customeritem.RecordSource = "select * from tbl_customer_item where ID_number like  '%" & frm_CI.txt_idnumber.text & "%'"
    frm_CI.ado_customeritem.Refresh
    
    frm_CI.lstview_Clientitem.ListItems.Clear
    
    Do Until frm_CI.ado_customeritem.Recordset.EOF
    Set list = frm_CI.lstview_Clientitem.ListItems.Add(, , frm_CI.ado_customeritem.Recordset(2) & "")
    
    list.SubItems(1) = frm_CI.ado_customeritem.Recordset(3) & ""
    frm_CI.ado_customeritem.Recordset.MoveNext
    Loop
    
    With frm_CI.lstview_Clientitem
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub

Sub Deliveryman_ListLoad()
Call Conn_ado_deliveryman
    frm_deliveryman.lstview_deliveryman.ListItems.Clear
    
    Do Until frm_deliveryman.ado_deliveryman.Recordset.EOF
    Set list = frm_deliveryman.lstview_deliveryman.ListItems.Add(, , frm_deliveryman.ado_deliveryman.Recordset(0) & "")
    
    list.SubItems(1) = frm_deliveryman.ado_deliveryman.Recordset(1) & ""
    list.SubItems(2) = frm_deliveryman.ado_deliveryman.Recordset(2) & ""
        
    frm_deliveryman.ado_deliveryman.Recordset.MoveNext
        Loop
    
    With frm_deliveryman.lstview_deliveryman
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub

Sub FORDelivery_ListLoad()
Call Conn_ado_FORdelivery
    frm_Deliver.lstview_order.ListItems.Clear
    
    Do Until frm_Deliver.ado_delivery.Recordset.EOF
    Set list = frm_Deliver.lstview_order.ListItems.Add(, , frm_Deliver.ado_delivery.Recordset(8) & "")
    
    list.SubItems(1) = frm_Deliver.ado_delivery.Recordset(0) & ""
    list.SubItems(2) = frm_Deliver.ado_delivery.Recordset(2) & ""
    list.SubItems(3) = frm_Deliver.ado_delivery.Recordset(3) & ""
    list.SubItems(4) = frm_Deliver.ado_delivery.Recordset(4) & ""
    list.SubItems(5) = frm_Deliver.ado_delivery.Recordset(5) & ""
    list.SubItems(6) = frm_Deliver.ado_delivery.Recordset(6) & ""
    list.SubItems(7) = frm_Deliver.ado_delivery.Recordset(7) & ""
          
    frm_Deliver.ado_delivery.Recordset.MoveNext
        Loop
    
    With frm_Deliver.lstview_order
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
    
End Sub

Sub Sales_ListLoad()
Call Conn_ado_sales
    frm_sales.lstview_sales.ListItems.Clear
    
    Do Until frm_sales.ado_sales.Recordset.EOF
    Set list = frm_sales.lstview_sales.ListItems.Add(, , frm_sales.ado_sales.Recordset(0) & "")
    
    list.SubItems(1) = frm_sales.ado_sales.Recordset(1) & ""
    list.SubItems(2) = frm_sales.ado_sales.Recordset(2) & ""
    list.SubItems(3) = frm_sales.ado_sales.Recordset(3) & ""
    list.SubItems(4) = frm_sales.ado_sales.Recordset(4) & ""
    list.SubItems(5) = frm_sales.ado_sales.Recordset(5) & ""
    list.SubItems(6) = frm_sales.ado_sales.Recordset(6) & ""
    list.SubItems(7) = frm_sales.ado_sales.Recordset(7) & ""
    
        
    frm_sales.ado_sales.Recordset.MoveNext
        Loop
    
    With frm_sales.lstview_sales
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
    gvCmd = frm_sales.ado_sales.RecordSource
End Sub

Sub Credit_ListLoad()
Call Conn_ado_credit
    frm_credit.lstview_credit.ListItems.Clear
    
    Do Until frm_credit.ado_credit.Recordset.EOF
    Set list = frm_credit.lstview_credit.ListItems.Add(, , frm_credit.ado_credit.Recordset(0) & "")
    
    list.SubItems(1) = frm_credit.ado_credit.Recordset(2) & ""
    list.SubItems(2) = frm_credit.ado_credit.Recordset(3) & ""
    list.SubItems(3) = frm_credit.ado_credit.Recordset(4) & ""
    list.SubItems(4) = frm_credit.ado_credit.Recordset(5) & ""
    list.SubItems(5) = frm_credit.ado_credit.Recordset(6) & ""
    list.SubItems(6) = frm_credit.ado_credit.Recordset(7) & ""
    
    
        
    frm_credit.ado_credit.Recordset.MoveNext
        Loop
    gvCmdCredit = frm_credit.ado_credit.RecordSource
    With frm_credit.lstview_credit
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub
Sub Credit_ListLoadS()
'Call Conn_ado_credit
    frm_credit.lstview_credit.ListItems.Clear
    
    Do Until frm_credit.ado_credit.Recordset.EOF
    Set list = frm_credit.lstview_credit.ListItems.Add(, , frm_credit.ado_credit.Recordset(0) & "")
    
    list.SubItems(1) = frm_credit.ado_credit.Recordset(2) & ""
    list.SubItems(2) = frm_credit.ado_credit.Recordset(3) & ""
    list.SubItems(3) = frm_credit.ado_credit.Recordset(4) & ""
    list.SubItems(4) = frm_credit.ado_credit.Recordset(5) & ""
    list.SubItems(5) = frm_credit.ado_credit.Recordset(6) & ""
    list.SubItems(6) = frm_credit.ado_credit.Recordset(7) & ""
    
    
        
    frm_credit.ado_credit.Recordset.MoveNext
        Loop
    
    With frm_credit.lstview_credit
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub

Sub Borrower_ListLoad()
Call Conn_ado_borrower
    frm_borrower.lstview_borrower.ListItems.Clear
    
    Do Until frm_borrower.ado_borrower.Recordset.EOF
    Set list = frm_borrower.lstview_borrower.ListItems.Add(, , frm_borrower.ado_borrower.Recordset(0) & "")
    
    list.SubItems(1) = frm_borrower.ado_borrower.Recordset(1) & ""
    list.SubItems(2) = frm_borrower.ado_borrower.Recordset(2) & ""
    list.SubItems(3) = frm_borrower.ado_borrower.Recordset(3) & ""
    list.SubItems(4) = frm_borrower.ado_borrower.Recordset(4) & ""
    
    
        
    frm_borrower.ado_borrower.Recordset.MoveNext
        Loop
    
    With frm_borrower.lstview_borrower
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub

Sub History_ListLoad()
Call Conn_ado_history
    frm_history.lstview_history.ListItems.Clear
    
    Do Until frm_history.ado_history.Recordset.EOF
    Set list = frm_history.lstview_history.ListItems.Add(, , frm_history.ado_history.Recordset(0) & "")
    
    list.SubItems(1) = frm_history.ado_history.Recordset(1) & ""
    list.SubItems(2) = frm_history.ado_history.Recordset(2) & ""
    list.SubItems(3) = frm_history.ado_history.Recordset(3) & ""
    list.SubItems(4) = frm_history.ado_history.Recordset(4) & ""
    
    
        
    frm_history.ado_history.Recordset.MoveNext
        Loop
    
    With frm_history.lstview_history
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub

Sub Task_ListLoad()
Dim i As Integer
Call Conn_ado_task
    frm_task.lstview_task.ListItems.Clear
    
    Do Until frm_task.ado_task.Recordset.EOF
    Set list = frm_task.lstview_task.ListItems.Add(, , frm_task.ado_task.Recordset(1) & "")
    
    list.SubItems(1) = frm_task.ado_task.Recordset(2) & ""
    list.SubItems(2) = frm_task.ado_task.Recordset(0) & ""
      
    If frm_task.ado_task.Recordset(2) = "Assigned" Then
    For i = 1 To 2
        list.ForeColor = vbRed
        list.ListSubItems(i).ForeColor = vbRed
    Next i
    End If
    frm_task.ado_task.Recordset.MoveNext
        Loop
    
    With frm_task.lstview_task
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    End If
    End With
    Set list = Nothing
End Sub



