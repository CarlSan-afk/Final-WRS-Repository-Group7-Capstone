VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rptQRCode 
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15825
   OleObjectBlob   =   "rptQRCode.dsx":0000
End
Attribute VB_Name = "rptQRCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QRGen As clsQRGen
Private Sub Section10_Format(ByVal pFormattingInfo As Object)

Dim data As String
Dim nfile As Long
Dim bmpHold As StdPicture
    If Dir(AppendPath(App.path, "\QRCodeimages\" & "qrcode.jpg")) <> "" Then
        Kill (AppendPath(App.path, "\QRCodeimages\" & "qrcode.jpg"))
    End If
    SavePicture frm_itemlist.Image1.Picture, App.path & "\QRCodeimages\qrcode.jpg"
    Set bmpHold = LoadPicture(AppendPath(App.path, "\QRCodeimages\" & "qrcode.jpg"))
    Set Picture2.FormattedPicture = bmpHold

End Sub
