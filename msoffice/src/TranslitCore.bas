Attribute VB_Name = "TranslitCore"
' Translit - pravila preslovljavanja latinica <-> cirilica i YUSCII.
'
' VBA verzija translit/scripts/translit.js i libreoffice/pythonpath/srtranslit.py.
' Kada menjate pravila, promenite sve tri i pokrenite tests/vba_core_test.py.
'
' Izvorni kod je samo ASCII (VBA ne cuva Unicode u kodu): slova se prave sa ChrW,
' a u regularnim izrazima se pisu kao \uXXXX.
' Kod mora da radi i kao VBScript (tako se testira bez Office-a): bez IIf, Like,
' GoTo, $ funkcija i opcionih parametara; "As Tip" test uklanja.
Option Explicit

Private gInit As Boolean
Private gLat2Cyr As Object
Private gCyr2Lat As Object
Private gDigraphs As Object
Private gYuLat As Object
Private gYuCyr As Object
Private gExceptions As Variant
Private gDjExceptions As Variant
Private gYuSymbols As String
Private gSrLatin As String

Private reRoman As Object
Private reForeign As Object
Private reWord As Object
Private reSrLatin As Object
Private reSentenceEnd As Object
Private reSpaceOnly As Object
Private reDigitUs As Object
Private reLetter As Object
Private reExcLetter As Object
Private reCombining As Object
Private reUrl As Object
Private reUnit1 As Object
Private reUnit2 As Object
Private reUnit3 As Object
Private reExc As Object

Private optSkipForeign As Boolean
Private optSkipRoman As Boolean
Private optSkipUrls As Boolean
Private optSkipUnits As Boolean
Private optDjToDj As Boolean
Private gExcKey As String
Private gHasExc As Boolean

' stanje jednog preslovljavanja
Private gMode As String
Private gOrig As String
Private gNorm As String
Private gPos As Long
Private gUseOffs As Boolean
Private gOffs() As Long
Private gProt() As String
Private gProtCount As Long
Private gToks() As String
Private gTokCount As Long

' Rezultat TrSpans: izmene (pocetak, kraj, novi tekst), polozaji od 0 u originalnom tekstu
Public TrCount As Long
Public TrA() As Long
Public TrB() As Long
Public TrOut() As String


' ------------------------------------------------------------ pomocno

' Niz Unicode znakova iz heksadecimalnih kodova: TrU("10D 107") = "cc" sa kvacicom/akcentom
Public Function TrU(ByVal codes As String) As String
    Dim parts As Variant, i As Long, s As String
    parts = Split(codes, " ")
    s = ""
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then s = s & ChrW(CLng("&H" & parts(i)))
    Next
    TrU = s
End Function

Private Function NewRe(ByVal pattern As String, ByVal isGlobal As Boolean) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.Global = isGlobal
    re.IgnoreCase = False
    re.MultiLine = False
    Set NewRe = re
End Function

Private Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    Set NewDict = d
End Function

Private Function IsUpper(ByVal ch As String) As Boolean
    IsUpper = False
    If Len(ch) = 0 Then Exit Function
    If ch <> LCase(ch) Then
        If ch = UCase(ch) Then IsUpper = True
    End If
End Function

Private Function IsLetter(ByVal ch As String) As Boolean
    IsLetter = False
    If Len(ch) = 0 Then Exit Function
    If LCase(ch) <> UCase(ch) Then
        IsLetter = True
    ElseIf ch = ChrW(&HB5) Or ch = ChrW(&HDF) Then
        ' mikro i sharp s: Windows ih ne menja u velika slova, a JavaScript/Python ih smatraju slovima
        IsLetter = True
    End If
End Function

Private Function ApplyCase(ByVal src As String, ByVal dst As String) As String
    If IsUpper(src) Then
        ApplyCase = UCase(dst)
    Else
        ApplyCase = dst
    End If
End Function

' Isto kao escapeRe u translit.js
Private Function EscapeRe(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    r = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If InStr(".*+?^${}()|[]\/", ch) > 0 Then r = r & "\"
        r = r & ch
    Next
    EscapeRe = r
End Function

' trim() iz JavaScript-a: skida i tabulatore, nove redove, NBSP...
Private Function TrimWs(ByVal s As String) As String
    Dim ws As String, a As Long, b As Long
    ws = vbTab & vbLf & ChrW(11) & ChrW(12) & vbCr & " " & TrU("A0 1680 2000 2001 2002 2003 2004 2005 2006 2007 2008 2009 200A 2028 2029 202F 205F 3000 FEFF")
    a = 1
    b = Len(s)
    Do While a <= b
        If InStr(ws, Mid(s, a, 1)) = 0 Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If InStr(ws, Mid(s, b, 1)) = 0 Then Exit Do
        b = b - 1
    Loop
    TrimWs = Mid(s, a, b - a + 1)
End Function


' ------------------------------------------------------------ pocetno podesavanje

Private Sub InitCore()
    Dim lat As String, cyr As String, i As Long, k As String, zh As String
    Dim latL As String, jsSpace As String, letter As String, u As String, unit As String, unitExpr As String, sup As String
    Dim yuKeys As String, yuVals As String, symFrom As String, symTo As String
    If gInit Then Exit Sub

    zh = TrU("17E")
    lat = "abc" & TrU("10D 107") & "d" & TrU("111") & "efghijklmnoprs" & TrU("161") & "tuvz" & zh
    cyr = TrU("430 431 446 447 45B 434 452 435 444 433 445 438 458 43A 43B 43C 43D 43E 43F 440 441 448 442 443 432 437 436")
    Set gLat2Cyr = NewDict()
    Set gCyr2Lat = NewDict()
    For i = 1 To Len(lat)
        gLat2Cyr.Add Mid(lat, i, 1), Mid(cyr, i, 1)
        gCyr2Lat.Add Mid(cyr, i, 1), Mid(lat, i, 1)
    Next
    gCyr2Lat.Add TrU("459"), "lj"
    gCyr2Lat.Add TrU("45A"), "nj"
    gCyr2Lat.Add TrU("45F"), "d" & zh

    Set gDigraphs = NewDict()
    gDigraphs.Add "lj", TrU("459")
    gDigraphs.Add "nj", TrU("45A")
    gDigraphs.Add "d" & zh, TrU("45F")

    ' "|" = mesto razdvajanja, "^" = samo na pocetku reci, "Z" = z sa kvacicom
    gExceptions = Split("in|jek in|junk kon|jug kon|junk kon|jukt van|jez van|jug tran|jug an|jon " & _
        "^nad|Zanj ^nad|Zelj ^nad|Ziv ^nad|Znje ^nad|Zup ^pod|Zanr ^pod|Zelud ^pod|Zup ^pod|Ziv ^pod|Znje " & _
        "^od|Ziv ^od|Zvaka", " ")
    For i = 0 To UBound(gExceptions)
        gExceptions(i) = Replace(gExceptions(i), "Z", zh)
    Next
    gDjExceptions = Split("^od|j ^nad|j ^pod|j ^pred|j ^ispod|j ^iznad|j", " ")

    ' YUSCII: JUS I.B1.002 (latinica) i JUS I.B1.003 (cirilica)
    Set gYuLat = NewDict()
    symFrom = "@`[{\|]}^~"
    symTo = TrU("17D 17E 160 161 110 111 106 107 10C 10D")
    For i = 1 To Len(symFrom)
        gYuLat.Add Mid(symFrom, i, 1), Mid(symTo, i, 1)
    Next
    Set gYuCyr = NewDict()
    yuKeys = "ABCDEFGHIJKLMNOPQRSTUVWXZ@[\]^"
    yuVals = TrU("410 411 426 414 415 424 413 425 418 408 41A 41B 41C 41D 41E 41F 409 420 421 422 423 412 40A 40F 417 416 428 402 40B 427")
    For i = 1 To Len(yuKeys)
        gYuCyr.Add Mid(yuKeys, i, 1), Mid(yuVals, i, 1)
    Next
    symFrom = "@[\]^"
    symTo = "`{|}~"
    For i = 1 To Len(yuKeys)
        k = Mid(yuKeys, i, 1)
        If InStr(symFrom, k) > 0 Then
            gYuCyr.Add Mid(symTo, InStr(symFrom, k), 1), LCase(Mid(yuVals, i, 1))
        Else
            gYuCyr.Add LCase(k), LCase(Mid(yuVals, i, 1))
        End If
    Next
    gYuSymbols = "@`[]{}\|^~"

    latL = "A-Za-z\u010C\u010D\u0106\u0107\u0110\u0111\u0160\u0161\u017D\u017E\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u00FF"
    jsSpace = "\u0009\u000A\u000B\u000C\u000D \u00A0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\uFEFF"
    letter = "[" & latL & "\u0400-\u04FF\u00B5\u03BC\u03A9]"
    sup = "(?:[\u00B2\u00B323])?"

    Set reRoman = NewRe("^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$", False)
    Set reForeign = NewRe("[qwxyQWXY]", False)
    Set reWord = NewRe("[" & latL & "\u0400-\u04FF0-9_]+|[^" & latL & "\u0400-\u04FF0-9_]+", True)
    Set reSrLatin = NewRe("[A-Za-z\u010C\u010D\u0106\u0107\u0110\u0111\u0160\u0161\u017D\u017E]", False)
    Set reSentenceEnd = NewRe("[.!?\u2026:;\u000A]", False)
    Set reSpaceOnly = NewRe("^[ \u00A0\u202F\u0009]+$", False)
    Set reDigitUs = NewRe("[0-9_]", False)
    Set reLetter = NewRe(letter, False)
    ' izuzeci se porede na tekstu pretvorenom u mala slova: LCase od omega je mala omega
    Set reExcLetter = NewRe("[" & latL & "\u0400-\u04FF\u00B5\u03BC\u03A9\u03C9]", False)
    Set reCombining = NewRe("[\u0300-\u036F]", False)

    Set reUrl = NewRe("(?:https?://|ftp://|www\.)[^" & jsSpace & "<>""]+" & _
        "|[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}" & _
        "|\b[A-Za-z0-9\-]+\.(?:com|org|net|rs|ba|me|hr|info|edu|gov|io)(?:/[^" & jsSpace & "<>""]*)?\b", True)

    ' Merne jedinice, najduze prvo (isti redosled kao UNITS u srtranslit.py)
    u = ""
    u = u & "kVAr|kcal|mbar|mmol|kbps|Mbps|Gbps|dag|min|kHz|MHz|GHz|THz|rpm|kWh|MWh|GWh|kVA|MVA|VAr|"
    u = u & "mAh|cal|keV|MeV|hPa|kPa|MPa|GPa|bar|atm|psi|dBm|mol|ppm|bps|dpi|nm|\u00B5m|\u03BCm|mm|"
    u = u & "cm|dm|km|ft|ha|ml|cl|dl|hl|cc|\u00B5g|mg|kg|lb|oz|ct|ns|\u00B5s|ms|Hz|mA|kA|mV|kV|MV|kW|"
    u = u & "MW|GW|mW|Wh|VA|Ah|k\u03A9|M\u03A9|\u00B5F|nF|pF|mH|Wb|kJ|MJ|eV|KS|hp|kN|MN|Pa|cd|lm|lx|"
    u = u & "dB|pH|kb|kB|KB|Mb|MB|Gb|GB|TB|PB|px|pt|m|l|L|g|t|s|h|A|V|W|\u03A9|F|H|C|T|S|J|N|K|b|B"
    unit = "(?:" & u & ")"
    unitExpr = unit & sup & "(?:[" & jsSpace & "]?[/\u00B7\u22C5*][" & jsSpace & "]?" & unit & sup & ")*"
    ' iza broja: "5 km", "10kg", "3,5 m2", "100 km/h"
    Set reUnit1 = NewRe("[0-9](?:[" & jsSpace & "]|\u00A0|\u202F)?" & unitExpr & "(?!" & letter & ")", False)
    ' slozene jedinice bilo gde: "km/h" (ispred ne sme biti slovo - proverava se posebno)
    Set reUnit2 = NewRe(unit & sup & "/" & unit & sup & "(?!" & letter & ")", False)
    ' stepeni: "C sa kruzicem"
    Set reUnit3 = NewRe("\u00B0[" & jsSpace & "]?[CFK](?!" & letter & ")", False)

    optSkipForeign = True
    optSkipRoman = True
    optSkipUrls = True
    optSkipUnits = True
    optDjToDj = False
    gExcKey = ""
    gHasExc = False
    gInit = True
End Sub


' ------------------------------------------------------------ opcije

Public Sub TrSetOptions(ByVal skipForeign As Boolean, ByVal skipRoman As Boolean, ByVal skipUrls As Boolean, _
                        ByVal skipUnits As Boolean, ByVal djToDj As Boolean, ByVal exceptions As String)
    InitCore
    optSkipForeign = skipForeign
    optSkipRoman = skipRoman
    optSkipUrls = skipUrls
    optSkipUnits = skipUnits
    optDjToDj = djToDj
    If exceptions <> gExcKey Then
        gExcKey = exceptions
        BuildExceptionsRe exceptions
    End If
End Sub

' Korisnicki izuzeci: jedna rec ili izraz po redu, "*" na kraju = bilo koji nastavak reci.
' Poredi se bez obzira na velika/mala slova (oba teksta se pretvore u mala slova).
Private Sub BuildExceptionsRe(ByVal text As String)
    Dim lines As Variant, parts() As String, n As Long, i As Long, j As Long, s As String, star As Boolean, tmp As String
    Dim letter As String
    gHasExc = False
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    ReDim parts(UBound(lines) + 1)
    n = 0
    letter = reExcLetter.Pattern
    For i = 0 To UBound(lines)
        s = TrimWs(lines(i))
        If s <> "" And s <> "*" Then
            star = (Right(s, 1) = "*")
            If star Then s = Left(s, Len(s) - 1)
            s = EscapeRe(LCase(s))
            If star Then s = s & letter & "*"
            parts(n) = s
            n = n + 1
        End If
    Next
    If n = 0 Then Exit Sub
    ' stabilno sortiranje po duzini, najduze prvo
    For i = 1 To n - 1
        tmp = parts(i)
        j = i - 1
        Do While j >= 0
            If Len(parts(j)) >= Len(tmp) Then Exit Do
            parts(j + 1) = parts(j)
            j = j - 1
        Loop
        parts(j + 1) = tmp
    Next
    s = parts(0)
    For i = 1 To n - 1
        s = s & "|" & parts(i)
    Next
    Set reExc = NewRe("(?:" & s & ")(?!" & letter & ")", False)
    gHasExc = True
End Sub


' ------------------------------------------------------------ rimski brojevi u kontekstu

Private Function IsRomanForm(ByVal w As String) As Boolean
    IsRomanForm = False
    If Len(w) > 0 Then IsRomanForm = reRoman.Test(w)
End Function

' "" znaci da prethodne reci nema (pocetak teksta ili recenice)
Private Function PrevWord(ByVal i As Long) As String
    PrevWord = ""
    If i < 2 Then Exit Function
    If reSentenceEnd.Test(gToks(i - 1)) Then Exit Function
    PrevWord = gToks(i - 2)
End Function

Private Function NextWord(ByVal i As Long) As String
    NextWord = ""
    If i + 2 >= gTokCount Then Exit Function
    If Not reSpaceOnly.Test(gToks(i + 1)) Then Exit Function
    NextWord = gToks(i + 2)
End Function

Private Function IsCapsWord(ByVal w As String) As Boolean
    IsCapsWord = False
    If Len(w) < 2 Then Exit Function
    If reDigitUs.Test(w) Then Exit Function
    If w <> UCase(w) Then Exit Function
    If w = LCase(w) Then Exit Function
    If IsRomanForm(w) Then Exit Function
    IsCapsWord = True
End Function

' Rimski broj, ili rec koja samo tako izgleda ("I tako...", "MAMA I TATA", "MI SMO")?
Private Function IsRoman(ByVal i As Long) As Boolean
    Dim w As String, nxt As String, prv As String, ln As String
    w = gToks(i)
    IsRoman = False
    If Not IsRomanForm(w) Then Exit Function
    IsRoman = True
    nxt = NextWord(i)
    If w = "I" Then
        If nxt = "" Then Exit Function
        If Not IsLetter(Left(nxt, 1)) Then Exit Function
        ln = LCase(nxt)
        If ln = "i" Or ln = "ili" Or ln = "do" Then
            If IsRomanForm(NextWord(i + 2)) Then Exit Function
        End If
        prv = PrevWord(i)
        If prv = "" Then
            IsRoman = False
        Else
            IsRoman = Not IsCapsWord(prv)
        End If
        Exit Function
    End If
    If w = "MI" Or w = "VI" Or w = "LI" Or w = "DIV" Then IsRoman = Not IsCapsWord(nxt)
End Function


' ------------------------------------------------------------ latinica -> cirilica

' Polozaji (od 0) ispred kojih se digraf razdvaja, kao "|3|5|"
Private Function SplitPositions(ByVal lw As String, ByVal lst As Variant) As String
    Dim r As String, j As Long, pat As String, anchored As Boolean, k As Long, plain As String, pos As Long
    r = "|"
    For j = 0 To UBound(lst)
        pat = lst(j)
        anchored = (Left(pat, 1) = "^")
        If anchored Then pat = Mid(pat, 2)
        k = InStr(pat, "|") - 1
        plain = Replace(pat, "|", "")
        If anchored Then
            If Left(lw, Len(plain)) = plain Then r = r & k & "|"
        Else
            pos = InStr(1, lw, plain, vbBinaryCompare)
            Do While pos > 0
                r = r & (pos - 1 + k) & "|"
                pos = InStr(pos + 1, lw, plain, vbBinaryCompare)
            Loop
        End If
    Next
    SplitPositions = r
End Function

Private Sub WordToCyrillic(ByVal idx As Long)
    Dim w As String, lw As String, n As Long, i As Long, ch As String, pair As String, lc As String
    Dim splits As String, djSplits As String, done As Boolean
    w = gToks(idx)
    If optSkipForeign Then
        If reForeign.Test(w) Then
            Emit Len(w), w
            Exit Sub
        End If
    End If
    If optSkipRoman Then
        If IsRoman(idx) Then
            Emit Len(w), w
            Exit Sub
        End If
    End If
    If Not reSrLatin.Test(w) Then
        Emit Len(w), w
        Exit Sub
    End If

    lw = LCase(w)
    splits = SplitPositions(lw, gExceptions)
    djSplits = "|"
    If optDjToDj Then djSplits = SplitPositions(lw, gDjExceptions)
    n = Len(w)
    i = 1
    Do While i <= n
        ch = Mid(w, i, 1)
        pair = Mid(lw, i, 2)
        done = False
        If Len(pair) = 2 Then
            If gDigraphs.Exists(pair) Then
                If InStr(splits, "|" & i & "|") = 0 Then
                    Emit 2, ApplyCase(ch, gDigraphs(pair))
                    i = i + 2
                    done = True
                End If
            End If
            If Not done And optDjToDj And pair = "dj" Then
                If InStr(djSplits, "|" & i & "|") = 0 Then
                    Emit 2, ApplyCase(ch, TrU("452"))
                    i = i + 2
                    done = True
                End If
            End If
        End If
        If Not done Then
            lc = Mid(lw, i, 1)
            If gLat2Cyr.Exists(lc) Then
                Emit 1, ApplyCase(ch, gLat2Cyr(lc))
            Else
                Emit 1, ch
            End If
            i = i + 1
        End If
    Loop
End Sub

Private Sub CyrChunk(ByVal chunk As String)
    Dim ms As Object, i As Long
    Set ms = reWord.Execute(chunk)
    gTokCount = ms.Count
    If gTokCount = 0 Then Exit Sub
    ReDim gToks(gTokCount - 1)
    For i = 0 To gTokCount - 1
        gToks(i) = ms(i).Value
    Next
    For i = 0 To gTokCount - 1
        WordToCyrillic i
    Next
End Sub


' ------------------------------------------------------------ cirilica -> latinica

Private Sub LatinChunk(ByVal t As String)
    Dim i As Long, n As Long, ch As String, lc As String, lat As String, nx As String, pv As String, allCaps As Boolean
    n = Len(t)
    For i = 1 To n
        ch = Mid(t, i, 1)
        lc = LCase(ch)
        If gCyr2Lat.Exists(lc) Then
            lat = gCyr2Lat(lc)
            If Not IsUpper(ch) Then
                Emit 1, lat
            ElseIf Len(lat) = 1 Then
                Emit 1, UCase(lat)
            Else
                ' Lj, Nj, Dz: verzalom ako je i susedno slovo veliko
                nx = ""
                pv = ""
                If i < n Then nx = Mid(t, i + 1, 1)
                If i > 1 Then pv = Mid(t, i - 1, 1)
                If IsLetter(nx) Then
                    allCaps = IsUpper(nx)
                Else
                    allCaps = IsLetter(pv) And IsUpper(pv)
                End If
                If allCaps Then
                    Emit 1, UCase(lat)
                Else
                    Emit 1, UCase(Left(lat, 1)) & Mid(lat, 2)
                End If
            End If
        Else
            Emit 1, ch
        End If
    Next
End Sub


' ------------------------------------------------------------ YUSCII

Private Function CharAt(ByVal t As String, ByVal k As Long) As String
    CharAt = ""
    If k >= 1 And k <= Len(t) Then CharAt = Mid(t, k, 1)
End Function

Private Function IsYuSymbol(ByVal c As String) As Boolean
    IsYuSymbol = False
    If Len(c) = 1 Then IsYuSymbol = (InStr(gYuSymbols, c) > 0)
End Function

Private Function Wordish(ByVal c As String) As Boolean
    Wordish = IsLetter(c)
    If Not Wordish Then Wordish = IsYuSymbol(c)
End Function

' YUSCII znak (npr. "[" ili "@") menja se samo kada je deo reci
Private Function InWord(ByVal t As String, ByVal i As Long) As Boolean
    Dim j As Long, k As Long
    j = i - 1
    Do While j >= 1
        If Not IsYuSymbol(Mid(t, j, 1)) Then Exit Do
        j = j - 1
    Loop
    k = i + 1
    Do While k <= Len(t)
        If Not IsYuSymbol(Mid(t, k, 1)) Then Exit Do
        k = k + 1
    Loop
    InWord = IsLetter(CharAt(t, j)) Or IsLetter(CharAt(t, k))
    If Not InWord Then InWord = Wordish(CharAt(t, i - 1)) And Wordish(CharAt(t, i + 1))
End Function

Private Sub YuChunk(ByVal t As String, ByVal tbl As Object)
    Dim i As Long, ch As String
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If Not tbl.Exists(ch) Then
            Emit 1, ch
        ElseIf IsYuSymbol(ch) Then
            If InWord(t, i) Then
                Emit 1, tbl(ch)
            Else
                Emit 1, ch
            End If
        Else
            Emit 1, tbl(ch)
        End If
    Next
End Sub


' ------------------------------------------------------------ zasticeni delovi teksta

' Prvo pojavljivanje izraza u searchIn od polozaja pos (od 0). Ako je zadat letterBefore,
' ispred pogotka ne sme biti slovo (lookbehind, koji VBScript nema) - inace se trazi dalje.
Private Function FindFrom(ByVal re As Object, ByVal searchIn As String, ByVal pos As Long, _
                          ByVal letterBefore As Object, ByRef mStart As Long, ByRef mLen As Long) As Boolean
    Dim ms As Object, s As Long, a As Long
    FindFrom = False
    s = pos
    Do While s < Len(searchIn)
        Set ms = re.Execute(Mid(searchIn, s + 1))
        If ms.Count = 0 Then Exit Function
        a = s + ms(0).FirstIndex
        If letterBefore Is Nothing Or a = 0 Then
            mStart = a
            mLen = ms(0).Length
            FindFrom = True
            Exit Function
        End If
        If Not letterBefore.Test(Mid(searchIn, a, 1)) Then
            mStart = a
            mLen = ms(0).Length
            FindFrom = True
            Exit Function
        End If
        s = a + 1
    Loop
End Function

Private Function FindProtected(ByVal kind As String, ByVal t As String, ByVal pos As Long, _
                               ByRef mStart As Long, ByRef mLen As Long) As Boolean
    Dim ms As Object, m As Object
    Dim s0 As Long, s1 As Long, s2 As Long, l0 As Long, l1 As Long, l2 As Long
    Dim ok0 As Boolean, ok1 As Boolean, ok2 As Boolean
    FindProtected = False
    If kind = "url" Then
        Set ms = reUrl.Execute(t)
        For Each m In ms
            If m.FirstIndex >= pos Then
                mStart = m.FirstIndex
                mLen = m.Length
                FindProtected = True
                Exit Function
            End If
        Next
    ElseIf kind = "exc" Then
        FindProtected = FindFrom(reExc, LCase(t), pos, reExcLetter, mStart, mLen)
    Else
        ' tri dela izraza za jedinice; pobedjuje onaj koji pocinje najranije, pa po redu
        ok0 = FindFrom(reUnit1, t, pos, Nothing, s0, l0)
        ok1 = FindFrom(reUnit2, t, pos, reLetter, s1, l1)
        ok2 = FindFrom(reUnit3, t, pos, Nothing, s2, l2)
        If ok1 Then
            If ok0 Then
                If s1 < s0 Then
                    s0 = s1
                    l0 = l1
                End If
            Else
                ok0 = True
                s0 = s1
                l0 = l1
            End If
        End If
        If ok2 Then
            If ok0 Then
                If s2 < s0 Then
                    s0 = s2
                    l0 = l2
                End If
            Else
                ok0 = True
                s0 = s2
                l0 = l2
            End If
        End If
        If ok0 Then
            mStart = s0
            mLen = l0
            FindProtected = True
        End If
    End If
End Function

' Delovi teksta koje hvata zastita ostaju netaknuti, ostatak se preslovljava.
Private Sub MapUnprotected(ByVal t As String, ByVal level As Long)
    Dim pos As Long, lastEnd As Long, mStart As Long, mLen As Long
    If Len(t) = 0 Then Exit Sub
    If level >= gProtCount Then
        ProcessChunk t
        Exit Sub
    End If
    pos = 0
    lastEnd = 0
    Do While pos < Len(t)
        If Not FindProtected(gProt(level), t, pos, mStart, mLen) Then Exit Do
        If mLen = 0 Then Exit Do
        MapUnprotected Mid(t, lastEnd + 1, mStart - lastEnd), level + 1
        Emit mLen, Mid(t, mStart + 1, mLen)
        lastEnd = mStart + mLen
        pos = lastEnd
    Loop
    MapUnprotected Mid(t, lastEnd + 1), level + 1
End Sub

Private Sub ProcessChunk(ByVal t As String)
    If gMode = "cyr" Then
        CyrChunk t
    ElseIf gMode = "lat" Then
        LatinChunk t
    ElseIf gMode = "yuLat" Then
        YuChunk t, gYuLat
    Else
        YuChunk t, gYuCyr
    End If
End Sub

Private Sub AddProt(ByVal kind As String)
    ReDim Preserve gProt(gProtCount)
    gProt(gProtCount) = kind
    gProtCount = gProtCount + 1
End Sub

Private Sub BuildProtectors()
    gProtCount = 0
    ReDim gProt(0)
    If gHasExc Then AddProt "exc"
    If gMode = "cyr" Then
        If optSkipUrls Then AddProt "url"
        If optSkipUnits Then AddProt "units"
    ElseIf gMode = "yuLat" Or gMode = "yuCyr" Then
        If optSkipUrls Then AddProt "url"
    End If
End Sub


' ------------------------------------------------------------ izmene (spans)

Private Sub AddSpan(ByVal a As Long, ByVal b As Long, ByVal outStr As String)
    If TrCount > UBound(TrA) Then
        ReDim Preserve TrA(TrCount * 2 + 16)
        ReDim Preserve TrB(TrCount * 2 + 16)
        ReDim Preserve TrOut(TrCount * 2 + 16)
    End If
    TrA(TrCount) = a
    TrB(TrCount) = b
    TrOut(TrCount) = outStr
    TrCount = TrCount + 1
End Sub

' srcLen znakova (normalizovanog) teksta postaje outStr
Private Sub Emit(ByVal srcLen As Long, ByVal outStr As String)
    Dim a As Long, b As Long
    a = gPos
    b = gPos + srcLen
    gPos = b
    If gUseOffs Then
        a = gOffs(a)
        b = gOffs(b)
    End If
    If Mid(gOrig, a + 1, b - a) <> outStr Then AddSpan a, b, outStr
End Sub

' NFC oblik (spaja slovo i kombinujuce znake koji slede: c + kvacica -> c sa kvacicom)
' i mapa polozaja u original.
Private Sub NfcWithOffsets(ByVal t As String)
    Dim i As Long, j As Long, n As Long, k As Long, cluster As String, norm As String, cnt As Long
    gUseOffs = False
    gNorm = t
    If Not reCombining.Test(t) Then Exit Sub
    n = Len(t)
    ReDim gOffs(n * 2 + 1)
    gNorm = ""
    cnt = 0
    i = 1
    Do While i <= n
        j = i + 1
        Do While j <= n
            If Not reCombining.Test(Mid(t, j, 1)) Then Exit Do
            j = j + 1
        Loop
        cluster = Mid(t, i, j - i)
        If j - i > 1 Then
            norm = TrNormalizeCluster(cluster)
        Else
            norm = cluster
        End If
        For k = 1 To Len(norm)
            If cnt > UBound(gOffs) - 1 Then ReDim Preserve gOffs(cnt * 2 + 2)
            If k = 1 Then
                gOffs(cnt) = i - 1
            Else
                gOffs(cnt) = j - 1
            End If
            cnt = cnt + 1
        Next
        gNorm = gNorm & norm
        i = j
    Loop
    If cnt > UBound(gOffs) Then ReDim Preserve gOffs(cnt)
    gOffs(cnt) = n
    gUseOffs = True
End Sub


' ------------------------------------------------------------ javne funkcije

' Izmene za tekst: vraca broj izmena u TrA/TrB/TrOut (polozaji od 0, po redu).
' mode: "cyr" (latinica -> cirilica), "lat", "yuLat", "yuCyr"
Public Function TrSpans(ByVal text As String, ByVal mode As String) As Long
    InitCore
    TrCount = 0
    ReDim TrA(15)
    ReDim TrB(15)
    ReDim TrOut(15)
    TrSpans = 0
    If Len(text) = 0 Then Exit Function
    gOrig = text
    gMode = mode
    If mode = "cyr" Then
        NfcWithOffsets text
    Else
        gNorm = text
        gUseOffs = False
    End If
    gPos = 0
    BuildProtectors
    MapUnprotected gNorm, 0
    TrSpans = TrCount
End Function

' Primenjuje izmene poslednjeg poziva TrSpans na isti tekst
Public Function TrApplySpans(ByVal text As String) As String
    Dim i As Long, r As String, last As Long
    r = ""
    last = 0
    For i = 0 To TrCount - 1
        r = r & Mid(text, last + 1, TrA(i) - last) & TrOut(i)
        last = TrB(i)
    Next
    TrApplySpans = r & Mid(text, last + 1)
End Function

Public Function TrText(ByVal text As String, ByVal mode As String) As String
    If TrSpans(text, mode) = 0 Then
        TrText = text
    Else
        TrText = TrApplySpans(text)
    End If
End Function
