Attribute VB_Name = "SalesFiltering"
Option Explicit
    Private CN As ADODB.Connection
    Private rs As ADODB.Recordset
    Private cmd As ADODB.Command
    Dim SQL As String
    Dim list As MSComctlLib.ListItem

Sub ConnDataBaseSales()
    Set CN = New ADODB.Connection
    'CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
    CN.ConnectionString = gvConnectString2
    CN.Open
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = CN
    cmd.CommandType = adCmdText
    End Sub
    
    
    
   Sub List2LOADSales()
    frm_sales.lbl_sales.Caption = 0
    
      SQL = " Select*  from tbl_sales where Customer_name like  '%" & frm_sales.txt_search.text & "%'  and date_of_sale between  '" & Format(frm_sales.DTPicker1.Value, "yyyy/MM/dd") & "'  and  '" & Format(frm_sales.DTPicker2.Value, "yyyy/MM/dd") & "'  "
   
    cmd.CommandText = SQL
    Set rs = cmd.Execute
    frm_sales.lstview_sales.ListItems.Clear
   ' On Error GoTo err
    With rs
    
        Do Until .EOF
    Set list = frm_sales.lstview_sales.ListItems.Add(, , !amount & "")
    
    list.SubItems(1) = !Customer_name & ""
    list.SubItems(2) = !classification & ""
    list.SubItems(3) = !Id_number & ""
    list.SubItems(4) = !date_of_sale & ""
    list.SubItems(5) = !delivered_by & ""
    list.SubItems(6) = !Sold_item & ""
    list.SubItems(7) = !TN & ""
        
      frm_sales.lbl_sales.Caption = Val(frm_sales.lbl_sales.Caption) + !amount & ""
    .MoveNext
    
        Loop
    End With
    
    With frm_sales.lstview_sales
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    
    End If
    End With
    gvCmd = rs.Source
    
    Set list = Nothing
    Set rs = Nothing
    End Sub
    
Sub ConnDataBaseExpenses()
    Set CN = New ADODB.Connection
    'CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\WRSv3.mdb;Persist Security Info=False;Jet OLEDB:Database Password=macie2018"
    CN.ConnectionString = gvConnectString2
    CN.Open
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = CN
    cmd.CommandType = adCmdText
    End Sub
    
Sub List2LOADExpenses()
    frm_sales.lbl_expenses.Caption = 0
    
      SQL = " Select* from tbl_expenses where date_of_expenses between  '" & Format(frm_sales.DTPicker1.Value, "yyyy/MM/dd") & "'  and  '" & Format(frm_sales.DTPicker2.Value, "yyyy/MM/dd") & "'  "
      
   cmd.CommandText = SQL
    Set rs = cmd.Execute
    frm_sales.lstview_expenses.ListItems.Clear
   ' On Error GoTo err
    With rs
    
        Do Until .EOF
    Set list = frm_sales.lstview_expenses.ListItems.Add(, , !Item_title & "")
    
    list.SubItems(1) = !Item_cost & ""
    list.SubItems(2) = !date_of_expenses & ""
    list.SubItems(3) = !TN & ""
           
      frm_sales.lbl_expenses.Caption = Val(frm_sales.lbl_expenses.Caption) + !Item_cost & ""
    .MoveNext
    
        Loop
    End With
    
    'frmPreview.CommandQuery = rs.Source
    With frm_sales.lstview_expenses
    If .ListItems.Count > 0 Then
    Set .SelectedItem = .ListItems(1)
    
    End If
    End With
    gvCmdExpenses = rs.Source
    Set list = Nothing
    Set rs = Nothing
    End Sub
