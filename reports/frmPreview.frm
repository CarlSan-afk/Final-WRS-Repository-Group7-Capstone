VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmPreview 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public Report      As Object
Public customerReportCmd As String


'Dim report As New rptItemSalesProfitReport


Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Execute_Report
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
CRViewer1.Zoom 100

End Sub

Public Sub Execute_Report()
Dim rs As ADODB.Recordset
Dim cmd As String
Dim lngRowsAffected As Long
    If gvReportValue = 1 Then
        Set Report = New rptDelivery
    ElseIf gvReportValue = 2 Then
        'CommandQuery = gvCmd
    End If
    cmd = CommandQuery
    
    Set rs = gvConnection.Execute(cmd, lngRowsAffected)
    
    Do While rs.State = 0
        Set rs = rs.NextRecordset
    Loop


    Set rs = gvConnection.Execute(cmd)
    Do While Not rs.State <> 0
        Set rs = rs.NextRecordset
    Loop
    If gvReportValue = 1 Then
        Set Report = New rptDelivery
    ElseIf gvReportValue = 2 Then
        Set Report = New rptSales
    ElseIf gvReportValue = 3 Then
        Set Report = New RptHistory
    ElseIf gvReportValue = 4 Then
        Set Report = New rptQRCode
    ElseIf gvReportValue = 5 Then
        Set Report = New rptExpenses
    ElseIf gvReportValue = 6 Then
        Set Report = New RptCredit
    ElseIf gvReportValue = 7 Then
        Set Report = New rptItemList
    ElseIf gvReportValue = 8 Then
        Set Report = New rptCustomerReport
    End If
'    ElseIf gvReportValue = 3 Then
'        Set Report = New rptItemSalesProfitReport
'    End If
'    If gvReportValue = 1 Then
'        Dim mvtotalamount As Currency, mvGrandtotal As Currency
'                mvtotalamount = 0
'                mvGrandtotal = 0
'                cmd = ""
'                If rs.RecordCount > 0 Then
'                    While Not rs.EOF
'                        mvtotalamount = Nz(rs("totalpayment"), 0)
'                        If rs("transcode") = "REC" Then
'                            mvGrandtotal = mvGrandtotal - mvtotalamount
'                        Else
'                            mvGrandtotal = mvGrandtotal + mvtotalamount
'                        End If
'                    rs.MoveNext
'                    Wend
'                End If
'         Report.Totalpayment.SetText Format(mvGrandtotal, "#,##0.00")
'         Report.txtDate.SetText "From: " & Format(dlgReportParameter.dtpfromdate.Value, "MMMM dd, yyyy") & " To " & Format(dlgReportParameter.dtpfromdate.Value, "MMMM dd, yyyy")
'    ElseIf gvReportValue = 3 Then
'        Dim mvprofit As Currency, mvtotalprofit As Currency, mvquantity As Currency, mvtotalquantity As Currency, mvnetprice As Currency, mvtotalnetprice As Currency, mvdiscount As Currency, mvtotaldiscount As Currency, mvGrossPrice As Currency, mvtotalGrossPrice As Currency, mvtotalcost As Currency, mvsumtotalcost As Currency
'        If rs.RecordCount > 0 Then
'            While Not rs.EOF
'                mvprofit = Nz(rs("profit"), 0)
'                mvquantity = Nz(rs("quantity"), 0)
'                mvnetprice = Nz(rs("NetPrice"), 0)
'                mvdiscount = Nz(rs("discount"), 0)
'                mvGrossPrice = Nz(rs("GrossPrice"), 0)
'                mvtotalcost = Nz(rs("totalcost"))
'                If rs("transcode") = "REC" Then
'                    mvtotalprofit = mvtotalprofit - mvprofit
'                    mvtotalquantity = mvtotalquantity - mvquantity
'                    mvtotalnetprice = mvtotalnetprice - mvnetprice
'                    mvtotaldiscount = mvtotaldiscount - mvdiscount
'                    mvtotalGrossPrice = mvtotalGrossPrice - mvGrossPrice
'                    mvsumtotalcost = mvsumtotalcost - mvtotalcost
'                Else
'                    mvtotalprofit = mvtotalprofit + mvprofit
'                    mvtotalquantity = mvtotalquantity + mvquantity
'                    mvtotalnetprice = mvtotalnetprice + mvnetprice
'                    mvtotaldiscount = mvtotaldiscount + mvdiscount
'                    mvtotalGrossPrice = mvtotalGrossPrice + mvGrossPrice
'                    mvsumtotalcost = mvsumtotalcost + mvtotalcost
'                End If
'            rs.MoveNext
'            Wend
'        End If
'        'Report.totalprofit.SetText mvtotalprofit
'        Report.totalcost.SetText Format(mvsumtotalcost, "#,##0.00")
'        Report.totalgross.SetText Format(mvtotalGrossPrice, "#,##0.00")
'        Report.totalnet.SetText Format(mvtotalnetprice, "#,##0.00")
'        Report.totalquantity.SetText mvtotalquantity
'        Report.totaldiscount.SetText Format(mvtotaldiscount, "#,##0.00")
'        Report.totalprofit.SetText Format(mvtotalprofit, "#,##0.00")
'        Report.txtDate.SetText "From: " & Format(dlgReportParameter.dtpfromdate.Value, "MMMM dd, yyyy") & " To " & Format(dlgReportParameter.dtpfromdate.Value, "MMMM dd, yyyy")
'    ElseIf gvReportValue = 4 Then
'        Dim payment As Currency, Totalpayment As Currency
'        If rs.RecordCount > 0 Then
'            While Not rs.EOF
'                payment = Nz(rs("paymentAmount"), 0)
'                If rs("transcode") = "REC" Then
'                    Totalpayment = Totalpayment - payment
'                Else
'                    Totalpayment = Totalpayment + payment
'                End If
'            rs.MoveNext
'            Wend
'        End If
'        Report.Totalpayment.SetText Format(Totalpayment, "#,##0.00")
'        Report.txtDate.SetText "From: " & Format(dlgReportParameter.dtpfromdate.Value, "MMMM dd, yyyy") & " To " & Format(dlgReportParameter.dtpfromdate.Value, "MMMM dd, yyyy")
'    ElseIf gvReportValue = 5 Or gvReportValue = 6 Then
'        Report.txtDate.SetText "From: " & Format(dlgReportParameter.dtpfromdate.Value, "MMMM dd, yyyy") & " To " & Format(dlgReportParameter.dtpfromdate.Value, "MMMM dd, yyyy")
'    End If
'
'    If gvTranscode = "REC" And gvTypecode = "POS" Then
'        If gvReportValue <> 8 Then
'            Report.txtDocheader.SetText ""
'        End If
'    Else
'        If gvReportValue = 9 Or gvReportValue = 8 Then
'            Report.txtDocheader.SetText ""
'        End If
'    End If
'    If gvReportValue = 10 Then
'        If Nz(rs("transcode"), "REC") = "REL" Then
'            Report.txtDocheader.SetText "Stock OUT"
'        Else
'            Report.txtDocheader.SetText "Stock IN"
'        End If
'    End If
    'Unload dlgReportParameter
    'Unload dlgCustomerReportParameter
    Report.Database.SetDataSource rs
    DoEvents
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    DoEvents
    Screen.MousePointer = vbDefault
End Sub

Private Function CommandQuery() As String
Dim cmd As String
On Error GoTo CommandQuery_Error
    If gvReportValue = 1 Then
        cmd = ""
        'cmd = cmd & "select balance,Customer_name,classification,date_of_sale,delivered_by,Sold_item from tbl_delivery;" & vbCrLf
        cmd = cmd & "select D.balance,d.Customer_name,c.address,d.date_of_sale,d.delivered_by,d.Sold_item,c.contact from tbl_delivery D left join tbl_customer_info C on C.customer_name=D.customer_name;" & vbCrLf
        
    ElseIf gvReportValue = 2 Then
        cmd = ""
        cmd = gvCmd
    ElseIf gvReportValue = 3 Then
        cmd = ""
        'cmd = dlgCustomerReportParameter.mvcustomerReportCmd
        cmd = gvCmd1
    ElseIf gvReportValue = 4 Then
        cmd = cmd & "select Item_title,qrCode from tbl_item_list where Item_title = '" & frm_itemlist.txt_itemtitle & "'" & vbCrLf
    ElseIf gvReportValue = 5 Then
        cmd = gvCmdExpenses
    ElseIf gvReportValue = 6 Then
        cmd = ""
        'cmd = cmd & "select * from tbl_credit" & vbCrLf
        cmd = gvCmdCredit
    ElseIf gvReportValue = 7 Then
        cmd = ""
        cmd = cmd & "select Item_title,qrCode,`type`,Discription,stocks,borrowed,damaged from tbl_item_list where pos_item='Yes'" & vbCrLf
    ElseIf gvReportValue = 8 Then
        cmd = ""
        cmd = gvCmdCustomer
    End If
        
    CommandQuery = cmd
    Exit Function
CommandQuery_Error:
    MsgBox "Error Executing Query at " & cmd
    Exit Function
    Resume
End Function
