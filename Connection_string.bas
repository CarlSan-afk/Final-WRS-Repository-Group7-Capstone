Attribute VB_Name = "Connection_string"
Public HostServer As String
Public HostUser As String
Public HostPassword As String
Public HostDatabase As String
Public gvConnectString As String            ' Database Connection String
Public gvConnectString2 As String
Public gvConnection As New ADODB.Connection
Sub Conn_ado_ci()
With frm_CI
.ado_ci.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_ci.RecordSource = "select * from tbl_customer_info order by date_of_last_buy desc"
.ado_ci.Refresh
End With
End Sub

Sub Conn_ado_additem()
With frm_additem
.ado_additem.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_additem.RecordSource = "select * from tbl_item_list"
.ado_additem.Refresh
End With
End Sub

Sub Conn_ado_itemlist()
With frm_itemlist
.ado_itemlist.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_itemlist.RecordSource = "select * from tbl_item_list"
.ado_itemlist.Refresh
End With
End Sub

Sub Conn_ado_expenses()
With frm_addexpenses
.ado_expenses.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_expenses.RecordSource = "select * from tbl_expenses"
.ado_expenses.Refresh
End With
End Sub

Sub Conn_ado_addstocks()
With frm_addstocks
.ado_addstocks.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_addstocks.RecordSource = "select * from tbl_item_list"
.ado_addstocks.Refresh
End With
End Sub
Sub Conn_ado_customer()
With frm_customer
.ado_customer.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_customer.RecordSource = "select * from tbl_customer_info order by date_of_last_buy desc "
.ado_customer.Refresh
End With
End Sub

Sub Conn_ado_main()
With frm_main
.ado_main.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_main.RecordSource = "select * from tbl_customer_info order by date_of_last_buy desc "
.ado_main.Refresh
End With
End Sub

Sub Conn_ado_customeritem()
With frm_main
.ado_customeritem.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_customeritem.RecordSource = "select * from tbl_customer_item"
.ado_customeritem.Refresh
End With
End Sub

Sub Conn_ado_delivery()
With frm_main
.ado_delivery.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_delivery.RecordSource = "select * from tbl_delivery"
.ado_delivery.Refresh
End With
End Sub

Sub Conn_ado_FORdelivery()
With frm_Deliver
.ado_delivery.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_delivery.RecordSource = "select * from tbl_delivery"
.ado_delivery.Refresh
End With
End Sub

Sub Conn_ado_customeritemCI()
With frm_CI
.ado_customeritem.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_customeritem.RecordSource = "select * from tbl_customer_item"
.ado_customeritem.Refresh
End With
End Sub

Sub Conn_ado_customeritemBR()
With frm_borrow
.ado_customeritem.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_customeritem.RecordSource = "select * from tbl_customer_item"
.ado_customeritem.Refresh
End With
End Sub

Sub Conn_ado_itemlistBR()
With frm_borrow
.ado_itemlist.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_itemlist.RecordSource = "select * from tbl_item_list"
.ado_itemlist.Refresh
End With
End Sub

Sub Conn_ado_customeritemRT()
With frm_return
.ado_customeritem.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_customeritem.RecordSource = "select * from tbl_customer_item"
.ado_customeritem.Refresh
End With
End Sub

Sub Conn_ado_itemlistRT()
With frm_return
.ado_itemlist.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_itemlist.RecordSource = "select * from tbl_item_list"
.ado_itemlist.Refresh
End With
End Sub

Sub Conn_MDI()


With mdi_wrs
.ado_ci.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_ci.RecordSource = "select * from tbl_customer_info order by date_of_last_buy desc "
.ado_ci.Refresh


.ado_itemlist.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_itemlist.RecordSource = "select * from tbl_item_list"
.ado_itemlist.Refresh

.ado_deliveryman.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_deliveryman.RecordSource = "select * from tbl_deliveryman"
.ado_deliveryman.Refresh

'.ado_activation.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Activation.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
'.ado_activation.RecordSource = "select * from tbl_activation"
'.ado_activation.Refresh


.status_mdi.Panels(2).Text = "Number of Customer: " & Val(.ado_ci.Recordset.RecordCount) - 1
End With
End Sub

Sub Conn_ado_deliveryman()
With frm_deliveryman
.ado_deliveryman.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_deliveryman.RecordSource = "select * from tbl_deliveryman"
.ado_deliveryman.Refresh
End With
End Sub

Sub Conn_ado_sales()
With frm_sales
.ado_sales.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_sales.RecordSource = "select * from tbl_sales"
.ado_sales.Refresh
End With
End Sub

Sub Conn_ado_salesMAIN()
With frm_main
.ado_sales.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_sales.RecordSource = "select * from tbl_sales"
.ado_sales.Refresh
End With
End Sub

Sub Conn_ado_creditMAIN()
With frm_main
.ado_credit.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_credit.RecordSource = "select * from tbl_credit"
.ado_credit.Refresh
End With
End Sub

Sub Conn_ado_credit()
With frm_credit
.ado_credit.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_credit.RecordSource = "select * from tbl_credit"
.ado_credit.Refresh
End With
End Sub

Sub Conn_ado_salesDELIVER()
With frm_Deliver
.ado_sales.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_sales.RecordSource = "select * from tbl_sales"
.ado_sales.Refresh
End With
End Sub

Sub Conn_ado_creditDELIVER()
With frm_Deliver
.ado_credit.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_credit.RecordSource = "select * from tbl_credit"
.ado_credit.Refresh
End With
End Sub


Sub Conn_ado_salesMDI()
With mdi_wrs
.ado_sales.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_sales.RecordSource = "select * from tbl_sales where date_of_sale between  '" & Format(Now, "yyyy/MM/dd") & "'  and  '" & Format(Now, "yyyy/MM/dd") & "'  "
.ado_sales.Refresh
'gvCmd = .ado_sales.RecordSource
End With

'Set gvRs = mdi_wrs.ado_sales.Clone
End Sub

Sub Conn_ado_refilledMDI()
With mdi_wrs
.ado_refilled.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_refilled.RecordSource = "select * from tbl_refilled where Date_refilled between  '" & Format(Now, "yyyy/MM/dd") & "'  and  '" & Format(Now, "yyyy/MM/dd") & "'  "
.ado_refilled.Refresh
End With
End Sub

Sub Conn_ado_borrower()
With frm_borrower
.ado_borrower.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_borrower.RecordSource = "select * from tbl_customer_item where Customer_name  like  '%" & frm_borrower.txt_search.Text & "%'"
.ado_borrower.Refresh
End With
End Sub
Sub Conn_ado_payment()
With frm_payment
.ado_payment.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_payment.RecordSource = "select * from tbl_credit"
.ado_payment.Refresh
End With
End Sub

Sub Conn_ado_salesPAYMENT()
With frm_payment
.ado_sales.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_sales.RecordSource = "select * from tbl_sales"
.ado_sales.Refresh
End With
End Sub

Sub Conn_ado_history()
With frm_history
.ado_history.ConnectionString = gvConnectString2 ' " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_history.RecordSource = "select * from tbl_history"
.ado_history.Refresh
End With
End Sub

Sub Conn_ado_task()
With frm_task
.ado_task.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_task.RecordSource = "select * from tbl_task  where Task_Date like  '%" & Format(.cal_task.Value, "yyyy-MM-dd") & "%'"
.ado_task.Refresh
End With
End Sub


Sub Conn_ado_account()
With frm_account
.ado_account.ConnectionString = gvConnectString2 ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_account.RecordSource = "select * from tbl_account"
.ado_account.Refresh
End With
End Sub

Sub Conn_ado_taskCUS()
With frm_customer
.ado_task.ConnectionString = gvConnectString2 '" Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
.ado_task.RecordSource = "select * from tbl_task "
.ado_task.Refresh
End With
End Sub










