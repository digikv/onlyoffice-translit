Attribute VB_Name = "TranslitShapes"
' Translit - tekst u oblicima (PowerPoint, i tekst-polja u Excel-u), preko Office TextRange2.
Option Explicit

Public TrShapeChanged As Long

' Pasus po pasus, od kraja; svaka izmena zadrzava formatiranje slova koje zamenjuje
Public Sub TrTextRange2(ByVal tr As Object, ByVal mode As String)
    Dim i As Long, p As Object, t As String, n As Long, k As Long
    For i = tr.Paragraphs.Count To 1 Step -1
        Set p = tr.Paragraphs(i)
        t = p.Text
        n = TrSpans(t, mode)
        For k = n - 1 To 0 Step -1
            p.Characters(TrA(k) + 1, TrB(k) - TrA(k)).Text = TrOut(k)
        Next
        TrShapeChanged = TrShapeChanged + n
    Next
End Sub

Public Sub TrShape(ByVal shp As Object, ByVal mode As String)
    Dim i As Long, r As Long, c As Long, isGroup As Boolean, hasTbl As Boolean, hasText As Boolean
    ' svojstvo koje oblik nema baca gresku - zato se svako cita posebno
    On Error Resume Next
    isGroup = False
    isGroup = (shp.Type = 6)
    If isGroup Then
        For i = 1 To shp.GroupItems.Count
            TrShape shp.GroupItems(i), mode
        Next
        Exit Sub
    End If
    hasTbl = False
    hasTbl = shp.HasTable
    If hasTbl Then
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                TrTextRange2 shp.Table.Cell(r, c).Shape.TextFrame2.TextRange, mode
            Next
        Next
        Exit Sub
    End If
    ' (Excel oblici nemaju HasTextFrame; oblik bez teksta baca gresku na TextFrame2)
    hasText = False
    hasText = shp.TextFrame2.HasText
    If hasText Then TrTextRange2 shp.TextFrame2.TextRange, mode
End Sub
