Attribute VB_Name = "Math_function"
Sub matinik()
Dim lngIndex As Long
    Dim lngTot As Long
    
    For lngIndex = 1 To frm_order.lstview_order.ListItems.Count
        lngTot = lngTot + frm_order.lstview_order.ListItems(lngIndex).SubItems(2)
    Next
    
    frm_order.lbl_total.Caption = lngTot

End Sub


Sub Creditmath()
Dim CreditIndex As Long
    Dim lngTotal As Long
    
    For CreditIndex = 1 To frm_credit.lstview_credit.ListItems.Count
        lngTotal = lngTotal + frm_credit.lstview_credit.ListItems(CreditIndex)
    Next
    
    frm_credit.lbl_total_credit.Caption = lngTotal
End Sub

Sub updateQTY()
    If frm_order.cbo_serviceoption.Text = "Buy" Then
        Dim lngIndex As Long
        Dim lngTot As Long, cmd As String
        'Dim rs As New ADODB.Recordset,
        For lngIndex = 1 To frm_order.lstview_order.ListItems.Count
            cmd = ""
            cmd = "update tbl_item_list set stocks = stocks - " & frm_order.lstview_order.ListItems(lngIndex).SubItems(1) & " where Item_title = '" & frm_order.lstview_order.ListItems(lngIndex).Text & "'"
            gvConnection.Execute (cmd)
        Next
    End If
End Sub
