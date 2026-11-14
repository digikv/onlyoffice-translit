Attribute VB_Name = "TranslitWord"
' Translit za Microsoft Word.
'
' Preslovljava oznaceni tekst, a ako nista nije oznaceno - ceo dokument (telo, tabele,
' zaglavlja i podnozja, fusnote, tekst-polja). Formatiranje se cuva: Word novom tekstu
' daje formatiranje prvog znaka koji zamenjuje, a izmene se prave samo unutar delova
' sa istim formatiranjem. Celo preslovljavanje se ponistava jednim Ctrl+Z.
Option Explicit

Private gChanged As Long
Private gHasRevisions As Boolean


' ------------------------------------------------------------ kartica i makroi

Public Sub Translit_Run(control As IRibbonControl)
    TranslitWordRun control.Tag
End Sub

' Makroi za precice sa tastature ili traku za brzi pristup
Public Sub TranslitLatinicaUCirilicu()
    TranslitWordRun "cyr"
End Sub

Public Sub TranslitCirilicaULatinicu()
    TranslitWordRun "lat"
End Sub

Public Sub TranslitYusciiULatinicu()
    TranslitWordRun "yuLat"
End Sub

Public Sub TranslitYusciiUCirilicu()
    TranslitWordRun "yuCyr"
End Sub


' ------------------------------------------------------------ pokretanje

Public Sub TranslitWordRun(ByVal mode As String)
    Dim doc As Document, sel As Selection, shp As Shape, undoOpen As Boolean, failed As Boolean, errText As String
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    Set sel = Selection
    gChanged = 0
    On Error GoTo fail
    TrApplyOptions
    gHasRevisions = (doc.Revisions.Count > 0)
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Translit"
    undoOpen = True

    If TrOption("scopeAll") Then
        WordAllStories doc, mode
    ElseIf sel.Type = wdSelectionShape Then
        For Each shp In sel.ShapeRange
            WordShape shp, mode
        Next
    ElseIf (sel.Type = wdSelectionNormal Or sel.Type = wdSelectionFrame Or sel.Type = wdSelectionColumn _
            Or sel.Type = wdSelectionRow Or sel.Type = wdSelectionBlock) And sel.Start < sel.End Then
        WordRange sel.Range, mode
    Else
        WordAllStories doc, mode
    End If

finish:
    On Error Resume Next
    If undoOpen Then Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    If failed Then
        TrShowError errText
    ElseIf gChanged = 0 Then
        TrShowNothing
    End If
    Exit Sub
fail:
    failed = True
    errText = Err.Description
    Resume finish
End Sub

Private Sub WordAllStories(ByVal doc As Document, ByVal mode As String)
    Dim story As Range, shp As Shape, junk As Long
    ' Word ponekad ne nabroji zaglavlja dok im se jednom ne pristupi
    junk = doc.Sections(1).Headers(1).Range.StoryType
    For Each story In doc.StoryRanges
        Do While Not story Is Nothing
            WordRange story, mode
            Select Case story.StoryType
                Case wdEvenPagesHeaderStory, wdPrimaryHeaderStory, wdEvenPagesFooterStory, _
                     wdPrimaryFooterStory, wdFirstPageHeaderStory, wdFirstPageFooterStory
                    If story.ShapeRange.Count > 0 Then
                        For Each shp In story.ShapeRange
                            WordShape shp, mode
                        Next
                    End If
            End Select
            Set story = story.NextStoryRange
        Loop
    Next
End Sub

Private Sub WordShape(ByVal shp As Shape, ByVal mode As String)
    Dim i As Long, isGroup As Boolean, hasText As Boolean
    On Error Resume Next
    isGroup = False
    isGroup = (shp.Type = msoGroup)
    If isGroup Then
        For i = 1 To shp.GroupItems.Count
            WordShape shp.GroupItems(i), mode
        Next
        Exit Sub
    End If
    hasText = False
    hasText = shp.TextFrame.HasText
    On Error GoTo 0
    If hasText Then WordRange shp.TextFrame.TextRange, mode
End Sub


' ------------------------------------------------------------ pasusi

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

' Pasusi opsega, od poslednjeg ka prvom (izmene ne pomeraju pasuse ispred)
Private Sub WordRange(ByVal rng As Range, ByVal mode As String)
    Dim p As Paragraph, starts() As Long, ends() As Long, n As Long, i As Long, r As Range, rs As Long, re As Long
    rs = rng.Start
    re = rng.End
    If rs >= re Then Exit Sub
    n = 0
    ReDim starts(63)
    ReDim ends(63)
    For Each p In rng.Paragraphs
        If n > UBound(starts) Then
            ReDim Preserve starts(n * 2)
            ReDim Preserve ends(n * 2)
        End If
        starts(n) = p.Range.Start
        ends(n) = p.Range.End
        n = n + 1
    Next
    For i = n - 1 To 0 Step -1
        Set r = rng.Duplicate
        r.SetRange MaxL(starts(i), rs), MinL(ends(i), re)
        If r.Start < r.End Then WordParagraph r, mode
    Next
End Sub

Private Sub WordParagraph(ByVal r As Range, ByVal mode As String)
    Dim t As String, n As Long
    r.TextRetrievalMode.IncludeFieldCodes = True
    r.TextRetrievalMode.IncludeHiddenText = True
    t = r.Text
    If Len(t) <> r.End - r.Start Then
        ' tekst se ne poklapa sa polozajima (npr. kontrole sadrzaja) - rec po rec
        WordByWords r, mode
        Exit Sub
    End If
    t = MaskFieldCodes(t)
    n = TrSpans(t, mode)
    If n > 0 Then ApplySpans r, r.Start, t
End Sub

Private Sub WordByWords(ByVal r As Range, ByVal mode As String)
    Dim w As Range, starts() As Long, ends() As Long, n As Long, i As Long, part As Range, t As String
    n = 0
    ReDim starts(63)
    ReDim ends(63)
    For Each w In r.Words
        If n > UBound(starts) Then
            ReDim Preserve starts(n * 2)
            ReDim Preserve ends(n * 2)
        End If
        starts(n) = w.Start
        ends(n) = w.End
        n = n + 1
    Next
    For i = n - 1 To 0 Step -1
        Set part = r.Duplicate
        part.SetRange starts(i), ends(i)
        part.TextRetrievalMode.IncludeFieldCodes = True
        part.TextRetrievalMode.IncludeHiddenText = True
        t = part.Text
        If Len(t) = part.End - part.Start Then
            t = MaskFieldCodes(t)
            If TrSpans(t, mode) > 0 Then ApplySpans part, part.Start, t
        End If
    Next
End Sub

' Kodovi polja (izmedju znakova 19 i 20) se ne preslovljavaju: zamene se znakom koji nije slovo
Private Function MaskFieldCodes(ByVal t As String) As String
    Dim i As Long, ch As String, stack As String, r As String, c As Long
    If InStr(t, ChrW(19)) = 0 Then
        MaskFieldCodes = t
        Exit Function
    End If
    stack = ""
    r = ""
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        c = AscW(ch)
        If c = 19 Then
            stack = stack & "C"
            r = r & ch
        ElseIf c = 20 Then
            If Len(stack) > 0 Then stack = Left(stack, Len(stack) - 1) & "R"
            r = r & ch
        ElseIf c = 21 Then
            If Len(stack) > 0 Then stack = Left(stack, Len(stack) - 1)
            r = r & ch
        ElseIf InStr(stack, "C") > 0 Then
            r = r & ChrW(&HFFFC)
        Else
            r = r & ch
        End If
    Next
    MaskFieldCodes = r
End Function

' Izmedju dve izmene su samo obicni znaci (ne polja, slike, fusnote, kraj celije...)
Private Function SafeGap(ByVal t As String, ByVal fromPos As Long, ByVal toPos As Long) As Boolean
    Dim i As Long, c As Long
    SafeGap = False
    For i = fromPos + 1 To toPos
        c = AscW(Mid(t, i, 1))
        If c < 32 And c >= 0 Then Exit Function
        If c = &HFFFC Or c = -4 Then Exit Function
    Next
    SafeGap = True
End Function

' Primenjuje izmene iz TrSpans na pasus koji pocinje na polozaju base, od kraja ka pocetku.
' Susedne izmene se spajaju u jednu zamenu kada je formatiranje tog dela isto.
Private Sub ApplySpans(ByVal r As Range, ByVal base As Long, ByVal t As String)
    Dim i As Long, j As Long, k As Long, outStr As String, part As Range
    i = TrCount - 1
    Do While i >= 0
        j = i
        If Not gHasRevisions Then
            Do While j > 0
                If Not SafeGap(t, TrB(j - 1), TrA(j)) Then Exit Do
                j = j - 1
            Loop
        End If
        If j < i Then
            Set part = r.Duplicate
            part.SetRange base + TrA(j), base + TrB(i)
            If IsUniform(part) Then
                outStr = ""
                For k = j To i
                    outStr = outStr & TrOut(k)
                    If k < i Then outStr = outStr & Mid(t, TrB(k) + 1, TrA(k + 1) - TrB(k))
                Next
                part.Text = outStr
                gChanged = gChanged + 1
            Else
                For k = i To j Step -1
                    ReplaceSpan r, base + TrA(k), base + TrB(k), TrOut(k)
                Next
            End If
        Else
            ReplaceSpan r, base + TrA(i), base + TrB(i), TrOut(i)
        End If
        i = j - 1
    Loop
End Sub

Private Sub ReplaceSpan(ByVal r As Range, ByVal s As Long, ByVal e As Long, ByVal outStr As String)
    Dim part As Range, rv As Revision
    Set part = r.Duplicate
    part.SetRange s, e
    If gHasRevisions Then
        ' obrisan tekst (pracenje izmena) se ne dira
        For Each rv In part.Revisions
            If rv.Type = wdRevisionDelete Then Exit Sub
        Next
    End If
    part.Text = outStr
    gChanged = gChanged + 1
End Sub

' Isto formatiranje u celom delu (Word vraca wdUndefined / "" kada je mesovito)
Private Function IsUniform(ByVal part As Range) As Boolean
    Dim f As Font
    IsUniform = False
    On Error GoTo done
    If part.Bookmarks.Count > 0 Or part.Hyperlinks.Count > 0 Or part.ContentControls.Count > 0 Then Exit Function
    Set f = part.Font
    If f.Name = "" Or f.Size = wdUndefined Then Exit Function
    If f.Bold = wdUndefined Or f.Italic = wdUndefined Or f.Underline = wdUndefined Then Exit Function
    If f.Color = wdUndefined Or f.StrikeThrough = wdUndefined Or f.DoubleStrikeThrough = wdUndefined Then Exit Function
    If f.Superscript = wdUndefined Or f.Subscript = wdUndefined Or f.Hidden = wdUndefined Then Exit Function
    If f.SmallCaps = wdUndefined Or f.AllCaps = wdUndefined Then Exit Function
    If f.Shading.BackgroundPatternColor = wdUndefined Then Exit Function
    If part.HighlightColorIndex = wdUndefined Then Exit Function
    IsUniform = True
done:
End Function
