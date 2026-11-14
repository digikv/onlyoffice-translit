Attribute VB_Name = "TranslitExcel"
' Translit za Microsoft Excel.
'
' Preslovljava oznacene celije sa tekstom (formule i brojevi se preskacu), a uz opciju
' "Uvek ceo dokument" sve listove radne sveske. Celije sa razlicitim formatiranjem unutar
' teksta menjaju se slovo po slovo, pa formatiranje ostaje. Ctrl+Z (Ponisti Translit)
' vraca prethodni tekst.
Option Explicit

Private gChanged As Long
Private gUndoCount As Long
Private gUndoBook As String
Private gUndoSheet() As String
Private gUndoAddr() As String
Private gUndoValue() As String
Private gUndoPrefix() As Boolean


' ------------------------------------------------------------ kartica i makroi

Public Sub Translit_Run(control As IRibbonControl)
    TranslitExcelRun control.Tag
End Sub

Public Sub TranslitLatinicaUCirilicu()
    TranslitExcelRun "cyr"
End Sub

Public Sub TranslitCirilicaULatinicu()
    TranslitExcelRun "lat"
End Sub

Public Sub TranslitYusciiULatinicu()
    TranslitExcelRun "yuLat"
End Sub

Public Sub TranslitYusciiUCirilicu()
    TranslitExcelRun "yuCyr"
End Sub


' ------------------------------------------------------------ pokretanje

Public Sub TranslitExcelRun(ByVal mode As String)
    Dim ws As Object, calc As Long, failed As Boolean, errText As String, sr As Object, calcSaved As Boolean
    If ActiveWorkbook Is Nothing Then Exit Sub
    gChanged = 0
    TrShapeChanged = 0
    gUndoCount = 0
    gUndoBook = ActiveWorkbook.Name
    ReDim gUndoSheet(255)
    ReDim gUndoAddr(255)
    ReDim gUndoValue(255)
    ReDim gUndoPrefix(255)
    On Error GoTo fail
    TrApplyOptions
    Application.ScreenUpdating = False
    calc = Application.Calculation
    calcSaved = True
    Application.Calculation = xlCalculationManual

    If TrOption("scopeAll") Then
        For Each ws In ActiveWorkbook.Worksheets
            ExcelCells ws.UsedRange, mode
            ExcelShapes ws.Shapes, mode
        Next
    ElseIf TypeName(Selection) = "Range" Then
        ExcelCells Selection, mode
    Else
        Set sr = Nothing
        On Error Resume Next
        Set sr = Selection.ShapeRange
        On Error GoTo fail
        If Not sr Is Nothing Then ExcelShapes sr, mode
    End If

finish:
    On Error Resume Next
    If calcSaved Then Application.Calculation = calc
    Application.ScreenUpdating = True
    If failed Then
        TrShowError errText
    ElseIf gChanged + TrShapeChanged = 0 Then
        TrShowNothing
    End If
    If gUndoCount > 0 Then Application.OnUndo "Translit", "'" & ThisWorkbook.Name & "'!TranslitExcelUndo"
    Exit Sub
fail:
    failed = True
    errText = Err.Description
    Resume finish
End Sub

Private Sub ExcelCells(ByVal rng As Object, ByVal mode As String)
    Dim cells As Object, c As Object
    If rng.Cells.CountLarge = 1 Then
        ExcelCell rng.Cells(1, 1), mode
        Exit Sub
    End If
    Set cells = Nothing
    On Error Resume Next
    Set cells = rng.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If cells Is Nothing Then Exit Sub
    For Each c In cells.Cells
        ExcelCell c, mode
    Next
End Sub

Private Sub ExcelShapes(ByVal shapes As Object, ByVal mode As String)
    Dim shp As Object
    For Each shp In shapes
        TrShape shp, mode
    Next
End Sub

Private Sub ExcelCell(ByVal c As Object, ByVal mode As String)
    Dim v As Variant, t As String, n As Long, nt As String, i As Long, rich As Boolean
    If c.HasFormula Then Exit Sub
    v = c.Value
    If VarType(v) <> vbString Then Exit Sub
    t = v
    n = TrSpans(t, mode)
    If n = 0 Then Exit Sub
    nt = TrApplySpans(t)
    RememberUndo c, t
    rich = IsRichText(c)
    If rich And Len(t) <= 255 Then
        For i = n - 1 To 0 Step -1
            c.Characters(TrA(i) + 1, TrB(i) - TrA(i)).Text = TrOut(i)
        Next
    ElseIf NeedsTextPrefix(c, nt) Then
        c.Value = "'" & nt
    Else
        c.Value = nt
    End If
    gChanged = gChanged + 1
End Sub

' Celija ima vise formata unutar teksta (Excel tada vraca Null)
Private Function IsRichText(ByVal c As Object) As Boolean
    Dim f As Object
    Set f = c.Font
    IsRichText = IsNull(f.Name) Or IsNull(f.Size) Or IsNull(f.Bold) Or IsNull(f.Italic) Or _
        IsNull(f.Underline) Or IsNull(f.Color) Or IsNull(f.Strikethrough) Or _
        IsNull(f.Superscript) Or IsNull(f.Subscript)
End Function

' Excel bi tekst koji lici na broj, datum ili formulu pretvorio - zato apostrof ispred
Private Function NeedsTextPrefix(ByVal c As Object, ByVal nt As String) As Boolean
    NeedsTextPrefix = False
    If c.NumberFormat = "@" Then Exit Function
    If c.PrefixCharacter = "'" Then
        NeedsTextPrefix = True
    ElseIf InStr("=+-@", Left(nt, 1)) > 0 And Len(nt) > 0 Then
        NeedsTextPrefix = True
    ElseIf IsNumeric(nt) Or IsDate(nt) Then
        NeedsTextPrefix = True
    End If
End Function


' ------------------------------------------------------------ ponistavanje

Private Sub RememberUndo(ByVal c As Object, ByVal oldText As String)
    If gUndoCount > UBound(gUndoAddr) Then
        ReDim Preserve gUndoSheet(gUndoCount * 2)
        ReDim Preserve gUndoAddr(gUndoCount * 2)
        ReDim Preserve gUndoValue(gUndoCount * 2)
        ReDim Preserve gUndoPrefix(gUndoCount * 2)
    End If
    gUndoSheet(gUndoCount) = c.Worksheet.Name
    gUndoAddr(gUndoCount) = c.Address
    gUndoValue(gUndoCount) = oldText
    gUndoPrefix(gUndoCount) = (c.PrefixCharacter = "'")
    gUndoCount = gUndoCount + 1
End Sub

Public Sub TranslitExcelUndo()
    Dim i As Long, wb As Object, c As Object
    On Error Resume Next
    Set wb = Workbooks(gUndoBook)
    If wb Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    For i = gUndoCount - 1 To 0 Step -1
        Set c = wb.Worksheets(gUndoSheet(i)).Range(gUndoAddr(i))
        If gUndoPrefix(i) Then
            c.Value = "'" & gUndoValue(i)
        Else
            c.Value = gUndoValue(i)
        End If
    Next
    Application.ScreenUpdating = True
    gUndoCount = 0
End Sub
