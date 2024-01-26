Attribute VB_Name = "CommonBas"
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public gvRs As New ADODB.Recordset
Public gvRs2 As New ADODB.Recordset
Public gvCmd As String
Public gvCmd1 As String
Public gvCmdCustomer As String
Public gvCmdExpenses As String
Public gvCmdCredit As String

Public Function RanmdonString()
    Dim rs As New ADODB.Recordset
    
Retry_RamdomQR:

    Randomize
    Dim RamdomQR As String, RawRamdomQR As String
    RamdomQR = RamdomQR & UCase(RamdomQR) & "0123456789"
    Dim i As Long
        For i = 1 To 13
        RawRamdomQR = RawRamdomQR & Mid$(RamdomQR, Int(Rnd() * Len(RamdomQR) + 1), 1)
    Next i
    
    Set rs = gvConnection.Execute("Select * from tbl_item_list where qrCode = '& RawRamdomQR &'")
    If rs.RecordCount > 0 Then
        RawRamdomQR = ""
        GoTo Retry_RamdomQR
    Else
        RanmdonString = RawRamdomQR
    End If
End Function
Public Function AppendPath(ByVal path As String, ByVal folder As String) As String
    If BeginsWith(folder, "\") Then
        folder = RemoveFirstLetter(folder)
    End If

    If Right(path, 1) = "\" Then
        AppendPath = path & folder
    Else
        AppendPath = path & "\" & folder
    End If
End Function
Public Function BeginsWith(text As String, beginWithText As String) As Boolean
    BeginsWith = Left(text, Len(beginWithText)) = beginWithText
End Function
Public Function RemoveFirstLetter(s As String)
    RemoveFirstLetter = Right(s, Len(s) - 1)
End Function
Public Function CreateFolder(folderName As String) As Long
On Error GoTo CreateFolder_Error

    CreateFolder = 0  'always return zero
    MkDir folderName
  
    Exit Function
  
CreateFolder_Error:
End Function
Public Function AppPath() As String
    AppPath = App.path & IIf(Right(App.path, 1) = "\", "", "\")
End Function
Public Function GetShortPath(pathName As String) As String
Dim shortPath As String
Dim length As Long

    shortPath = String(1024, " ")
    length = GetShortPathName(pathName, shortPath, 1024)
    shortPath = Left(shortPath, length)

    ' Add slash if there is none
    If Right(shortPath, 1) <> "\" Then
        shortPath = shortPath & "\"
    End If
    GetShortPath = shortPath
End Function
