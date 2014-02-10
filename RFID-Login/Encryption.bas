Attribute VB_Name = "Encryption"
Public Function EncodeString(ByVal strToEncode As String, ByVal strPassword As String) As String
Dim strResult As String                                 'Verschlüsseln
Dim i As Long
Dim cfc() As Integer
  ReDim cfc(1 To Len(strPassword))
  For i = 1 To UBound(cfc)
    cfc(i) = Asc(Right(strPassword, Len(strPassword) - i + 1))
  Next i
  For i = 1 To Len(strToEncode)
    strResult = strResult & Chr(addToIndex(Asc(Right(strToEncode, Len(strToEncode) - i + 1)), VirtPos(i, cfc)))
  Next i
  EncodeString = strResult
End Function

Public Function DecodeString(ByVal strToDecode As String, ByVal strPassword As String) As String
Dim strResult As String                                 'Entschlüsseln
Dim i As Long
Dim cfc() As Integer
  ReDim cfc(1 To Len(strPassword))
  ReDim ttc(1 To Len(strToDecode))
  For i = 1 To UBound(cfc)
    cfc(i) = Asc(Right(strPassword, Len(strPassword) - i + 1))
  Next i
  For i = 1 To Len(strToDecode)
    strResult = strResult & Chr(GetOfIndex(Asc(Right(strToDecode, Len(strToDecode) - i + 1)), VirtPos(i, cfc)))
  Next i
  DecodeString = strResult
End Function

Private Function VirtPos(i As Long, a() As Integer) As Integer
  If i > UBound(a) Then
    VirtPos = VirtPos(i - UBound(a), a)
  Else
    VirtPos = a(i)
  End If
End Function

Private Function addToIndex(i As Integer, j As Integer) As Integer
  If i + j > 255 Then
    addToIndex = i + j - 255
  Else
    addToIndex = i + j
  End If
End Function

Private Function GetOfIndex(i As Integer, j As Integer) As Integer
  If i - j < 0 Then
    GetOfIndex = i - j + 255
  Else
    GetOfIndex = i - j
  End If
End Function
