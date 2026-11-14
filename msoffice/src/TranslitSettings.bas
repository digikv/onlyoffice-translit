Attribute VB_Name = "TranslitSettings"
' Translit - podesavanja (zajednicka za Word, Excel i PowerPoint) i dugmad na kartici.
'
' Opcije se cuvaju u registru: HKCU\Software\VB and VBA Program Settings\Translit\Options
' Izuzeci: %APPDATA%\Translit\izuzeci.txt (UTF-8, jedna rec ili izraz po redu)
Option Explicit

Private Const APP_NAME As String = "Translit"

Public Function TrDefault(ByVal key As String) As Boolean
    TrDefault = Not (key = "djToDj" Or key = "scopeAll")
End Function

Public Function TrOption(ByVal key As String) As Boolean
    Dim v As String
    v = GetSetting(APP_NAME, "Options", key, "")
    If v = "" Then
        TrOption = TrDefault(key)
    Else
        TrOption = (v = "1")
    End If
End Function

Public Sub TrSetOption(ByVal key As String, ByVal value As Boolean)
    If value Then
        SaveSetting APP_NAME, "Options", key, "1"
    Else
        SaveSetting APP_NAME, "Options", key, "0"
    End If
End Sub

Public Function TrExceptionsFile() As String
    TrExceptionsFile = Environ("APPDATA") & "\Translit\izuzeci.txt"
End Function

Public Function TrReadExceptions() As String
    Dim fso As Object, st As Object, f As String
    TrReadExceptions = ""
    On Error GoTo failed
    f = TrExceptionsFile()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(f) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile f
    TrReadExceptions = st.ReadText(-1)
    st.Close
    Exit Function
failed:
    TrReadExceptions = ""
End Function

' Otvara spisak izuzetaka u Notepad-u (pravi ga ako ne postoji)
Public Sub TrEditExceptions()
    Dim fso As Object, st As Object, f As String, folder As String
    On Error GoTo failed
    f = TrExceptionsFile()
    folder = Environ("APPDATA") & "\Translit"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    If Not fso.FileExists(f) Then
        Set st = CreateObject("ADODB.Stream")
        st.Type = 2
        st.Charset = "utf-8"
        st.Open
        st.WriteText ""
        st.SaveToFile f, 2
        st.Close
    End If
    Shell "notepad.exe """ & f & """", vbNormalFocus
    Exit Sub
failed:
    MsgBox "Spisak izuzetaka nije mogu" & TrU("107") & "e otvoriti:" & vbCrLf & f & vbCrLf & vbCrLf & Err.Description, _
        vbExclamation, "Translit"
End Sub

' Prenosi opcije i izuzetke jezgru pre svakog preslovljavanja
Public Sub TrApplyOptions()
    TrSetOptions TrOption("skipForeign"), TrOption("skipRoman"), TrOption("skipUrls"), _
        TrOption("skipUnits"), TrOption("djToDj"), TrReadExceptions()
End Sub

Public Sub TrShowError(ByVal description As String)
    MsgBox "Gre" & TrU("161") & "ka pri preslovljavanju:" & vbCrLf & vbCrLf & description, vbExclamation, "Translit"
End Sub

Public Sub TrShowNothing()
    MsgBox "Nema teksta za preslovljavanje.", vbInformation, "Translit"
End Sub


' ------------------------------------------------------------ kartica Translit (ribbon)

Public Sub Translit_OnLoad(ribbon As IRibbonUI)
End Sub

Public Sub Translit_GetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = TrOption(control.Tag)
End Sub

Public Sub Translit_Toggle(control As IRibbonControl, pressed As Boolean)
    TrSetOption control.Tag, pressed
End Sub

Public Sub Translit_EditExceptions(control As IRibbonControl)
    TrEditExceptions
End Sub
