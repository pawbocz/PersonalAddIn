Attribute VB_Name = "modRG"
'==========================  modRG  ==========================
Option Explicit


'––– sta³e dla arkusza „Stawki” ––––––––––––––––––––––––––––––
Private Const SHEET_STAWKI As String = "Stawki"
Private Const COL_NAZWA    As Long = 1    'A
Private Const COL_KAT      As Long = 2    'B
Private Const COL_MIN      As Long = 3    'C   (minuty RG)

'-------------------------------------------------------------
' 1) Normalizacja tekstu
'-------------------------------------------------------------
Private Function CleanTxt(s As String) As String
    CleanTxt = LCase$(Trim$(Replace(Replace(s, vbTab, " "), Chr(160), " ")))
End Function

'-------------------------------------------------------------
' 2) Klucz przekroju (zamiana „×”, „*”, spacje, przecinki › kropki)
'-------------------------------------------------------------
Private Function NormPrzekrojKey(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, ChrW(215), "x")
    t = Replace(t, "*", "x")
    t = Replace(t, " ", "")
    t = Replace(t, ",", ".")
    NormPrzekrojKey = t
End Function

'-------------------------------------------------------------
' 3) Funkcja do wyci¹gania numeru koryta 50…600
'-------------------------------------------------------------
Private Function GetTrayWidth(txt As String) As String
    Dim reW As Object: Set reW = CreateObject("VBScript.RegExp")
    reW.Global = False: reW.IgnoreCase = True
    reW.Pattern = "(?:\b[kd]\s*(\d{2,3})\b)|(?:\b(\d{2,3})\s*mm\b)"

    If reW.Test(txt) Then
        Dim m As Object: Set m = reW.Execute(txt)(0)
        Dim numTxt As String
        If m.SubMatches(0) <> "" Then
            numTxt = m.SubMatches(0)     'K100 / D300 …
        Else
            numTxt = m.SubMatches(1)     '… 300 mm
        End If

        Select Case numTxt              'filtr dozwolonych
            Case "50", "100", "200", "300", "400", "500", "600"
                GetTrayWidth = numTxt
        End Select
    End If
End Function


Public Function WyodrebnijPrzekroj(opis As String) As String
    Dim re3 As Object: Set re3 = CreateObject("VBScript.RegExp")
    re3.Pattern = "^\s*\d+\s*[x×*]\s*(\d+\s*[x×*]\s*\d+(?:[\,\.]\d+)?)"
    re3.IgnoreCase = True
    If re3.Test(opis) Then
        'tu korzystamy z grupy przechwytuj¹cej
        WyodrebnijPrzekroj = NormPrzekrojKey(CStr(re3.Execute(opis)(0).SubMatches(0)))
        Exit Function
    End If

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d+\s*[x×*]\s*\d+(?:[\,\.]\d+)?)|(\bdn\d+\b)"
    re.IgnoreCase = True
    If re.Test(opis) Then
        Dim m As Object: Set m = re.Execute(opis)(0)
        Dim raw As String
        If m.SubMatches.Count > 0 And Len(m.SubMatches(0)) > 0 Then
            raw = CStr(m.SubMatches(0))
        Else
            raw = CStr(m.Value)
        End If
        WyodrebnijPrzekroj = NormPrzekrojKey(raw)
    Else
        WyodrebnijPrzekroj = ""
    End If
End Function


'================  S £ O W N I K I  RG  (exact + max)  =================
'
' BuildDicts(dictExact, dictMax)  ›  True/False
'   • dictExact  – klucze „kategoria|szczegó³” (np.  kable|5x25,  kor|300)
'   • dictMax    – maksymalna wartoœæ RG w danej kategorii (np.  kor › 60)
'=======================================================================
Private Function BuildDicts(ByRef dictExact As Object, _
                            ByRef dictMax As Object) As Boolean
    On Error GoTo Fail

    Set dictExact = CreateObject("Scripting.Dictionary")
    Set dictMax = CreateObject("Scripting.Dictionary")

    '–– 1. ustal skoroszyt ------------------------------------------------
    Dim wb As Workbook, vCaller As Variant
    On Error Resume Next
    vCaller = Application.Caller         'Range / pusty / Error
    On Error GoTo 0

    Select Case TypeName(vCaller)
        Case "Range":  Set wb = vCaller.Parent.Parent
        Case "String": Set wb = ActiveWorkbook
        Case Else:     Set wb = ActiveWorkbook
    End Select
    If wb Is Nothing Then Exit Function     'nie znaleziono – zwróci False

    '–– 2. arkusz „Stawki” -----------------------------------------------
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_STAWKI)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function     'brak – zwróci False

    '–– 3. zakres danych (tabela lub zwyk³y) ------------------------------
    Dim rng As Range
    If ws.ListObjects.Count > 0 Then
        Set rng = ws.ListObjects(1).DataBodyRange
    Else
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, COL_NAZWA).End(xlUp).Row
        If lastRow < 2 Then Exit Function
        Set rng = ws.Range(ws.Cells(2, COL_NAZWA), ws.Cells(lastRow, COL_MIN))
    End If

    '–– 4. budowanie s³owników -------------------------------------------
    Dim r As Range, cat As String, nazwa As String
    Dim minVal As Double, keyDot As String, keyComma As String

    For Each r In rng.Columns(COL_NAZWA).Cells
        nazwa = CleanTxt(r.Value)
        cat = CleanTxt(r.Offset(0, COL_KAT - COL_NAZWA).Value)
        If nazwa <> "" And cat <> "" Then
            minVal = CDbl(r.Offset(0, COL_MIN - COL_NAZWA).Value)

            '— dictMax: trzymaj najwy¿sz¹ wartoœæ w kategorii --------------
            If Not dictMax.Exists(cat) Then
                dictMax(cat) = minVal
            ElseIf minVal > dictMax(cat) Then
                dictMax(cat) = minVal
            End If

            ''— dictExact: wpisy szczegó³owe -------------------------------
            If InStr(cat, "kabl") > 0 Then
                '–––– KABLE ––––––––––––––––––––––––––––––––––––––––
                keyDot = NormPrzekrojKey(nazwa)          'np. 5x2.5
                If keyDot <> "" Then
                    dictExact(cat & "|" & keyDot) = minVal
                    keyComma = Replace(keyDot, ".", ",")  '5x2,5
                    dictExact(cat & "|" & keyComma) = minVal
                End If
            
            ElseIf InStr(cat, "kor") > 0 Then
                '–––– KORYTA – wychwyæ 50–600, nawet gdy po liczbie stoi przecinek, spacja, “mm”… ––––
                Dim reW As Object, m As Object
                Set reW = CreateObject("VBScript.RegExp")
                reW.Global = True: reW.IgnoreCase = True
                ' liczba 50–600, za któr¹ NIE stoi kolejna cyfra  (negative-look-ahead)
                reW.Pattern = "(50|100|200|300|400|500|600)(?!\d)"
                
                If reW.Test(nazwa) Then
                    For Each m In reW.Execute(nazwa)
                        dictExact(cat & "|" & m.SubMatches(0)) = minVal      'np. kor|100
                    Next m
                End If
            Else
                '–––– POZOSTA£E – pierwsze s³owo –––––––––––––––––––––––
                dictExact(cat & "|" & Split(nazwa, " ")(0)) = minVal
            End If
        End If
    Next r

    BuildDicts = True
    Exit Function
Fail:
    BuildDicts = False
End Function
'=======================================================================

'=============================================================
' 5) Funkcja arkuszowa – zwraca MINUTY RG (0, gdy brak)
'=============================================================
Public Function Roboczogodziny(kategoria As String, opis As String) As Double
    Static dictExact As Object, dictMax As Object

    '-- pierwszy raz w sesji: zbuduj s³owniki ------------------
    If dictExact Is Nothing Then
        If Not BuildDicts(dictExact, dictMax) Then
            Roboczogodziny = 0
            Exit Function
        End If
    End If

    Dim cat As String: cat = CleanTxt(kategoria)
    If cat = "" Then Roboczogodziny = 0: Exit Function

    '-----------------------------------------------------------
    ' 1) KABLE – szukamy dok³adnego przekroju (5x2,5 …)
    '-----------------------------------------------------------
    If InStr(cat, "kabl") > 0 Then
        Dim pr As String: pr = WyodrebnijPrzekroj(opis)
        If pr <> "" Then
            Dim k As String: k = cat & "|" & pr
            If dictExact.Exists(k) Then
                Roboczogodziny = dictExact(k)
                Exit Function
            End If
        End If

    '---------------------------------------------------------
    ' 2) KORYTA – szukamy szerokoœci (50-600)
    '-----------------------------------------------------------
    ElseIf InStr(cat, "kor") > 0 Then
        Dim trayW As String: trayW = GetTrayWidth(opis)   'np. "100"
        If trayW <> "" Then
            Dim kTray As String: kTray = cat & "|" & trayW
            If dictExact.Exists(kTray) Then
                Roboczogodziny = dictExact(kTray)
                Exit Function
            End If
        End If

    End If   ' <-- UWAGA: zamyka ELSEIF

    '-----------------------------------------------------------
    ' 3) Fallback – najwy¿sza wartoœæ w kategorii
    '-----------------------------------------------------------
    If dictMax.Exists(cat) Then
        Roboczogodziny = dictMax(cat)
    Else
        Roboczogodziny = 0
    End If
End Function



'=============================================================
' 6) Formularz + makro wstawiaj¹ce formu³y bez nadpisywania
'=============================================================
Public Sub WstawFormulyRG_Ask()
    Dim frm As New frmRGParams
    frm.Show
    If Not frm.FormOK Then Exit Sub

    WstawFormulyRG frm.OutCol, frm.CatCol, frm.DescCol, frm.FirstRow
End Sub

Public Sub WstawFormulyRG( _
    ByVal COL_OUT As Long, _
    ByVal COL_CAT As Long, _
    ByVal COL_DESC As Long, _
    ByVal FIRST_ROW As Long)

    Const BRAK_RG_COLOR As Long = vbRed    'RGB(255,0,0)

    Dim wb As Workbook: Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 2) = "LV" Then

            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow < FIRST_ROW Then GoTo NextWs

            Dim r As Long, adrCat As String, adrDesc As String, f As String
            For r = FIRST_ROW To lastRow

                adrCat = ws.Cells(r, COL_CAT).Address(False, False)
                adrDesc = ws.Cells(r, COL_DESC).Address(False, False)
                f = "=IFERROR(Roboczogodziny(" & adrCat & "," & adrDesc & "),0)"

                With ws.Cells(r, COL_OUT)
                    '1) Nie nadpisuj – pisz tylko, gdy brak formu³y i pusto/zero
                    If (Not .HasFormula) And (Len(.Value2) = 0 Or val(.Value2) = 0) Then
                        .Formula = f
                    End If

                    '2) Koloruj braki (kategoria jest, a wynik = 0)
                    Dim hasCat As Boolean, isZero As Boolean, isRed As Boolean
                    hasCat = (Len(Trim$(ws.Cells(r, COL_CAT).Value2)) > 0)
                    isZero = (val(.Value2) = 0)
                    isRed = (.Interior.Color = BRAK_RG_COLOR)

                    If hasCat And isZero Then
                        If Not isRed Then .Interior.Color = BRAK_RG_COLOR
                    Else
                        If isRed Then .Interior.Pattern = xlNone
                    End If
                End With
            Next r
        End If
NextWs:
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Formu³y RG dodane (tylko do pustych/zerowych). Braki oznaczone na czerwono.", _
           vbInformation
End Sub
'========================  /modRG  ==========================


Public Sub RG_RebuildCache()
    '--- 1. skasuj dotychczasowe cache ---
    Static dictExact As Object, dictMax As Object
    Set dictExact = Nothing
    Set dictMax = Nothing

    '--- 2. zbuduj ponownie (korzystamy z Twojej BuildDicts) ---
    Dim ok As Boolean
    ok = BuildDicts(dictExact, dictMax)

    If ok Then
        MsgBox "S³ownik roboczogodzin przebudowany." & vbCrLf & _
               "• dok³adnych kluczy: " & dictExact.Count & vbCrLf & _
               "• kategorii (MAX):  " & dictMax.Count, vbInformation
    Else
        MsgBox "Nie uda³o siê odœwie¿yæ s³ownika (brak danych w arkuszu 'Stawki'?).", _
               vbExclamation
    End If
End Sub

'---------------------------------------------------------------
'  SZYBKI PODGL¥D S£OWNIKA  –  TYLKO DO DEBUG
'---------------------------------------------------------------

Public Function RG_DictValue(keyTxt As String) As Variant
    Static dictExact As Object, dictMax As Object
    If dictExact Is Nothing Then Call BuildDicts(dictExact, dictMax)

    keyTxt = LCase$(Trim$(keyTxt))
    If dictExact.Exists(keyTxt) Then
        RG_DictValue = dictExact(keyTxt)
    Else
        RG_DictValue = CVErr(xlErrNA)
    End If
End Function


Public Sub RG_DumpCategory(catTxt As String)
    Static dictExact As Object, dictMax As Object
    If dictExact Is Nothing Then Call BuildDicts(dictExact, dictMax)

    catTxt = LCase$(Trim$(catTxt)) & "|"
    Dim k As Variant, cnt As Long
    Debug.Print "––– klucze w dictExact dla kategorii [" & catTxt & "] –––"
    For Each k In dictExact.Keys
        If Left$(k, Len(catTxt)) = catTxt Then
            Debug.Print k, dictExact(k)
            cnt = cnt + 1
        End If
    Next k
    Debug.Print "£¹cznie:", cnt, "elementów"
End Sub

