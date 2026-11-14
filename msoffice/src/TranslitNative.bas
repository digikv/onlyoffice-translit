Attribute VB_Name = "TranslitNative"
' Translit - NFC normalizacija preko Windows API-ja (slovo + kombinujuci znak -> jedno slovo).
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function NormalizeString Lib "Normaliz.dll" (ByVal NormForm As Long, _
    ByVal lpSrcString As LongPtr, ByVal cwSrcLength As Long, ByVal lpDstString As LongPtr, ByVal cwDstLength As Long) As Long
#Else
Private Declare Function NormalizeString Lib "Normaliz.dll" (ByVal NormForm As Long, _
    ByVal lpSrcString As Long, ByVal cwSrcLength As Long, ByVal lpDstString As Long, ByVal cwDstLength As Long) As Long
#End If

Public Function TrNormalizeCluster(ByVal s As String) As String
    Dim buf As String, n As Long
    TrNormalizeCluster = s
    If Len(s) = 0 Then Exit Function
    On Error GoTo failed
    buf = String(Len(s) * 3 + 8, vbNullChar)
    n = NormalizeString(1, StrPtr(s), Len(s), StrPtr(buf), Len(buf))
    If n > 0 Then TrNormalizeCluster = Left(buf, n)
    Exit Function
failed:
    TrNormalizeCluster = s
End Function
