Attribute VB_Name = "MAC_Address"
Option Explicit
 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long
 
Public Function GetMacAddress() As String
    Const OFFSET_LENGTH As Long = 400
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim lIdx            As Long
    Dim sRetVal         As String
    
    Call GetAdaptersInfo(ByVal 0, lSize)
    If lSize <> 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call GetAdaptersInfo(baBuffer(0), lSize)
        Call CopyMemory(lSize, baBuffer(OFFSET_LENGTH), 4)
        For lIdx = OFFSET_LENGTH + 4 To OFFSET_LENGTH + 4 + lSize - 1
            sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & ":", vbNullString) & Right$("0" & Hex$(baBuffer(lIdx)), 2)
        Next
    End If
    sRetVal = Left$(sRetVal, Len(sRetVal) - 3)
    GetMacAddress = sRetVal
    
End Function
Public Sub NumberOnly(KeyAscii As Integer)
    Select Case KeyAscii
    Case Asc("0") To Asc("9")
    'Case Asc("-")
    Case Str(8)
    Case Str(13)
    Case Asc(".")
    Case Else
        KeyAscii = 0
    End Select
End Sub
