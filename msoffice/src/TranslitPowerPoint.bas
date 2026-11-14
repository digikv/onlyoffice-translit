Attribute VB_Name = "TranslitPowerPoint"
' Translit za Microsoft PowerPoint.
'
' Preslovljava oznaceni tekst ili oznacene oblike, a ako nista nije oznaceno - celu
' prezentaciju (svi slajdovi, tabele, grupe i beleske). Formatiranje se cuva.
Option Explicit


' ------------------------------------------------------------ kartica i makroi

Public Sub Translit_Run(control As IRibbonControl)
    TranslitPptRun control.Tag
End Sub

Public Sub TranslitLatinicaUCirilicu()
    TranslitPptRun "cyr"
End Sub

Public Sub TranslitCirilicaULatinicu()
    TranslitPptRun "lat"
End Sub

Public Sub TranslitYusciiULatinicu()
    TranslitPptRun "yuLat"
End Sub

Public Sub TranslitYusciiUCirilicu()
    TranslitPptRun "yuCyr"
End Sub


' ------------------------------------------------------------ pokretanje

Public Sub TranslitPptRun(ByVal mode As String)
    Dim sel As Object, selType As Long, sld As Object, shp As Object, done As Boolean
    Dim failed As Boolean, errText As String, selText As String
    If Presentations.Count = 0 Then Exit Sub
    TrShapeChanged = 0
    done = False
    On Error GoTo fail
    TrApplyOptions

    If Not TrOption("scopeAll") Then
        selType = 0
        selText = ""
        On Error Resume Next
        Set sel = ActiveWindow.Selection
        selType = sel.Type
        If selType = ppSelectionText Then selText = sel.TextRange2.Text
        On Error GoTo fail
        If selType = ppSelectionText And Len(selText) > 0 Then
            TrTextRange2 sel.TextRange2, mode
            done = True
        ElseIf selType = ppSelectionShapes Then
            For Each shp In sel.ShapeRange
                TrShape shp, mode
            Next
            done = True
        End If
    End If

    If Not done Then
        For Each sld In ActivePresentation.Slides
            For Each shp In sld.Shapes
                TrShape shp, mode
            Next
            For Each shp In sld.NotesPage.Shapes
                TrShape shp, mode
            Next
        Next
    End If

finish:
    On Error Resume Next
    If failed Then
        TrShowError errText
    ElseIf TrShapeChanged = 0 Then
        TrShowNothing
    End If
    Exit Sub
fail:
    failed = True
    errText = Err.Description
    Resume finish
End Sub
