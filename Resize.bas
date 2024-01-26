Attribute VB_Name = "Resize"
Sub DoCenter(child As Form, parent As Form)
Dim mTop As Integer
Dim mLeft As Integer

If parent.WindowState = vbNormal Then
mTop = 0 '((parent.Height - child.Height) \ 2)
mLeft = 0 '((parent.Width - child.Width) \ 2)
child.Move mLeft, mTop
ElseIf parent.WindowState = vbMaximized Then
mTop = ((parent.Height - child.Height) \ 2)
mLeft = ((parent.Width - child.Width) \ 2)
child.Move mLeft, mTop
Else
Exit Sub

End If

End Sub
Public Function Nz(ByVal testValue As Variant, Optional returnValue As Variant) As Variant
    If IsNull(testValue) Or testValue = "" Then
        If IsMissing(returnValue) Then
            Nz = ""
        Else
            Nz = returnValue
        End If
    Else
        Nz = testValue
    End If
End Function

Public Function closeform()
'   Unload frm_account
'   Unload frm_borrower
'   Unload frm_CI
'   Unload frm_credit
'   Unload frm_customer
'   Unload frm_Deliver
'   Unload frm_deliveryman
'   Unload frm_history
'   Unload frm_itemlist
'   'Unload frm_main
'   Unload frm_sales
End Function
