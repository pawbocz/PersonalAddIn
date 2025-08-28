Attribute VB_Name = "modRG"
'==========================  modRG  ==========================
Option Explicit

Private gDictExact As Object
Private gDictMax   As Object

Private Const SHEET_STAWKI As String = "Stawki"
Private Const COL_NAZWA    As Long = 1
Private Const COL_KAT      As Long = 2
Private Const COL_MIN      As Long = 3


Private Function CleanTxt(s As String) As String
    CleanTxt = LCase$(Trim$(Replace(Replace(s, vbTab, " "), Chr(160), " ")))
End Function


Private Function NormPrzekrojKey(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, ChrW(215), "x")
    t = Replace(t, "*", "x")
    t = Replace(t, " ", "")
    t = Replace(t, ",", ".")
    NormPrzekrojKey = t
End Function


Private Function GetTrayWidth(txt As String) As String
    Dim reW As Object: Set reW = CreateObject("VBScript.RegExp")
    reW.Global = False: reW.IgnoreCase = True
    reW.Pattern = "(?:\b[kd]\s*(\d{2,3})\b)|(?:\b(\d{2,3})\s*mm\b)"

    If reW.test(txt) Then
        Dim m As Object: Set m = reW.Execute(txt)(0)
        Dim numTxt As String
        If m.SubMatches(0) <> "" Then
            numTxt = m.SubMatches(0)
        Else
            numTxt = m.SubMatches(1)
        End If

        Select Case numTxt
            Case "50", "100", "200", "300", "400", "500", "600"
                GetTrayWidth = numTxt
        End Select
    End If
End Function


Public Function WyodrebnijPrzekroj(opis As String) As String
    Dim re3 As Object: Set re3 = CreateObject("VBScript.RegExp")
    re3.Pattern = "^\s*\d+\s*[x×*]\s*(\d+\s*[x×*]\s*\d+(?:[\,\.]\d+)?)"
    re3.IgnoreCase = True
    If re3.test(opis) Then
        
        WyodrebnijPrzekroj = NormPrzekrojKey(CStr(re3.Execute(opis)(0).SubMatches(0)))
        Exit Function
    End If

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d+\s*[x×*]\s*\d+(?:[\,\.]\d+)?)|(\bdn\d+\b)"
    re.IgnoreCase = True
    If re.test(opis) Then
        Dim m As Object: Set m = re.Execute(opis)(0)
        Dim raw As String
        If m.SubMatches.Count > 0 And Len(m.SubMatches(0)) > 0 Then
            raw = CStr(m.SubMatches(0))
        Else
            raw = CStr(m.value)
        End If
        WyodrebnijPrzekroj = NormPrzekrojKey(raw)
    Else
        WyodrebnijPrzekroj = ""
    End If
End Function



Private Function BuildDicts(ByRef dictExact As Object, _
                            ByRef dictMax As Object) As Boolean
    On Error GoTo Fail

    Set dictExact = CreateObject("Scripting.Dictionary")
    Set dictMax = CreateObject("Scripting.Dictionary")

    '–– 1. ustal skoroszyt ------------------------------------------------
    Dim wb As Workbook, vCaller As Variant
    On Error Resume Next
    vCaller = Application.Caller
    On Error GoTo 0

    Select Case TypeName(vCaller)
        Case "Range":  Set wb = vCaller.Parent.Parent
        Case "String": Set wb = ActiveWorkbook
        Case Else:     Set wb = ActiveWorkbook
    End Select
    If wb Is Nothing Then Exit Function

    '–– 2. arkusz „Stawki” -----------------------------------------------
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_STAWKI)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

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
        nazwa = CleanTxt(r.value)
        cat = CleanTxt(r.Offset(0, COL_KAT - COL_NAZWA).value)
        If nazwa <> "" And cat <> "" Then
            minVal = CDbl(r.Offset(0, COL_MIN - COL_NAZWA).value)

            '— dictMax: trzymaj najwy¿sz¹ wartoœæ w kategorii --------------
            If Not dictMax.Exists(cat) Then
                dictMax(cat) = minVal
            ElseIf minVal > dictMax(cat) Then
                dictMax(cat) = minVal
            End If

            ''— dictExact: wpisy szczegó³owe -------------------------------
            If InStr(cat, "kabl") > 0 Then
                '–––– KABLE ––––––––––––––––––––––––––––––––––––––––
                keyDot = NormPrzekrojKey(nazwa)
                If keyDot <> "" Then
                    dictExact(cat & "|" & keyDot) = minVal
                    keyComma = Replace(keyDot, ".", ",")  '
                    dictExact(cat & "|" & keyComma) = minVal
                End If
            
            ElseIf InStr(cat, "kor") > 0 Then
                '–––– KORYTA – wychwyæ 50–600, nawet gdy po liczbie stoi przecinek, spacja, “mm”… ––––
                Dim reW As Object, m As Object
                Set reW = CreateObject("VBScript.RegExp")
                reW.Global = True: reW.IgnoreCase = True
                
                reW.Pattern = "(50|100|200|300|400|500|600)(?!\d)"
                
                If reW.test(nazwa) Then
                    For Each m In reW.Execute(nazwa)
                        dictExact(cat & "|" & m.SubMatches(0)) = minVal
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

Public Function Roboczogodziny(kategoria As String, opis As String) As Double
    Application.Volatile True  'pozwala przeliczyæ przy Calculate/CalculateFull

    'Zbuduj cache przy pierwszym wywo³aniu lub po rêcznym resetcie
    If gDictExact Is Nothing Or gDictMax Is Nothing Then
        If Not BuildDicts(gDictExact, gDictMax) Then
            Roboczogodziny = 0
            Exit Function
        End If
    End If

    Dim cat As String: cat = CleanTxt(kategoria)
    If cat = "" Then Roboczogodziny = 0: Exit Function

    '1) kable – dok³adny przekrój
    If InStr(cat, "kabl") > 0 Then
        Dim pr As String: pr = WyodrebnijPrzekroj(opis)
        If pr <> "" Then
            Dim k As String: k = cat & "|" & pr
            If gDictExact.Exists(k) Then
                Roboczogodziny = gDictExact(k)
                Exit Function
            End If
        End If
    End If

    '2) koryta – dopasowanie po szerokoœci (np. kor|100)
    If cat = "kor" Or cat = "kor_pokr" Then
        Dim w As String: w = GetTrayWidth(opis)
        If w <> "" Then
            Dim kk As String: kk = cat & "|" & w
            If gDictExact.Exists(kk) Then
                Roboczogodziny = gDictExact(kk)
                Exit Function
            End If
        End If
    End If

    '3) fallback – maksimum dla kategorii
    If gDictMax.Exists(cat) Then
        Roboczogodziny = gDictMax(cat)
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

    WstawFormulyRG frm.OutCol, frm.CatCol, frm.DescCol, frm.firstRow
End Sub

Public Sub WstawFormulyRG( _
    ByVal COL_OUT As Long, _
    ByVal COL_CAT As Long, _
    ByVal COL_DESC As Long, _
    ByVal FIRST_ROW As Long)

    Const BRAK_RG_COLOR As Long = vbRed

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
                    
                    If (Not .HasFormula) And (Len(.Value2) = 0 Or val(.Value2) = 0) Then
                        .Formula = f
                    End If

                    
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


Public Sub RG_RebuildCache(Optional ByVal Recalculate As Boolean = True)
    '1) wyczyœæ cache modu³owy
    Set gDictExact = Nothing
    Set gDictMax = Nothing

    '2) zbuduj ponownie ze „Stawek”
    Dim ok As Boolean
    ok = BuildDicts(gDictExact, gDictMax)

    '3) ewentualnie przelicz wszystkie formu³y z UDF
    If Recalculate Then Application.CalculateFull

    If ok Then
        MsgBox "S³ownik roboczogodzin odœwie¿ony." & vbCrLf & _
               "• exact: " & gDictExact.Count & vbCrLf & _
               "• max/kategoria: " & gDictMax.Count, vbInformation
    Else
        MsgBox "Nie uda³o siê odœwie¿yæ s³ownika (sprawdŸ arkusz 'Stawki').", vbExclamation
    End If
End Sub



'  SZYBKI PODGL¥D S£OWNIKA  –  TYLKO DO DEBUG


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

