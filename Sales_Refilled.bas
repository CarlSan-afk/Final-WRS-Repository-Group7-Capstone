Attribute VB_Name = "Sales_Refilled"

Sub Status_Sales_Refilled()
Dim salesday As Long
Dim refilled As Long
'=======================
Call Conn_ado_salesMDI
Do Until mdi_wrs.ado_sales.Recordset.EOF
salesday = salesday + Val(mdi_wrs.ado_sales.Recordset(0))
mdi_wrs.ado_sales.Recordset.MoveNext
Loop
mdi_wrs.status_mdi.Panels(4) = "Sales for Today: " & salesday

'=======================
Call Conn_ado_refilledMDI
Do Until mdi_wrs.ado_refilled.Recordset.EOF
refilled = refilled + Val(mdi_wrs.ado_refilled.Recordset(0))
mdi_wrs.ado_refilled.Recordset.MoveNext
Loop
mdi_wrs.status_mdi.Panels(3) = "Total Refilled for Today: " & refilled

End Sub





