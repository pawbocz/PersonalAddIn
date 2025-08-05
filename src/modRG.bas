Attribute VB_Name = "modRG"
'==========================  modRG  ==========================
Option Explicit

'–––– sta³e dla arkusza „Stawki” –––––––––––––––––––––––––––––
Private Const SHEET_STAWKI As String = "Stawki"
Private Const COL_NAZWA    As Long = 1       'A  (np. 5x25, K600 …)
Private Const COL_KAT      As Long = 2       'B  (nazwa kategorii)
Private Const COL_MIN      As Long = 3       'C  (minuty RG)

'=============================================================
' 1) Normalizacja tekstu (trim, lower, NBSP›spacja)
'=============================================================
Private Function CleanTxt(s As String) As String
    CleanTxt = LCase$(Trim$(Replace(Replace(s, vbTab, " "), Chr(160), " ")))
End Function

'– normalizacja klucza przekroju
Private Function NormPrzekrojKey(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, ChrW(215), "x")   ' × › x
    t = Replace(t, "*", "x")         ' * › x
    t = Replace(t, " ", "")          ' usuñ spacje
    t = Replace(t, ",", ".")         ' , › .
    NormPrzekrojKey = t
End Function

Public Function WyodrebnijPrzekroj(opis As String) As String
    Dim re3 As Object: Set re3 = CreateObject("VBScript.RegExp")
    re3.Pattern = "^\s*\d+\s*[x×*]\s*(\d+\s*[x×*]\s*\d+(?:[\,\.]\d+)?)"
    re3.IgnoreCase = True
    If re3.test(opis) Then
        'tu korzystamy z grupy przechwytuj¹cej
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
            raw = CStr(m.Value)
        End If
        WyodrebnijPrzekroj = NormPrzekrojKey(raw)
    Else
        WyodrebnijPrzekroj = ""
    End If
End Function


'=============================================================
' 4) Budowa s³owników:
'     • dictExact("kat|klucz") = min (klucze: orygina³, 5x2.5, 5x2,5, 1. s³owo…)
'     • dictMax("kat") = najwy¿sza wartoœæ w kategorii
'     Zwraca True, jeœli OK.
'=============================================================
Private Function BuildDicts(ByRef dictExact As Object, ByRef dictMax As Object) As Boolean
    On Error GoTo oops
    Set dictExact = CreateObject("Scripting.Dictionary")
    Set dictMax = CreateObject("Scripting.Dictionary")

    '— wybór workbooka (dzia³a z funkcji i z makra) —
    Dim wb As Workbook
    Dim vCaller As Variant
    On Error Resume Next
    vCaller = Application.Caller            'przy makrze: Error 2023
    On Error GoTo 0
    Select Case TypeName(vCaller)
        Case "Range":  Set wb = vCaller.Parent.Parent
        Case "String": Set wb = ActiveWorkbook
        Case Else:     Set wb = ActiveWorkbook
    End Select
    If wb Is Nothing Then Exit Function

    '— arkusz „Stawki” —
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_STAWKI)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    '— tabela lub zakres —
    Dim rng As Range
    If ws.ListObjects.Count > 0 Then
        Set rng = ws.ListObjects(1).DataBodyRange
    Else
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, COL_NAZWA).End(xlUp).Row
        If lastRow < 2 Then Exit Function
        Set rng = ws.Range(ws.Cells(2, COL_NAZWA), ws.Cells(lastRow, COL_MIN))
    End If

    '— budowa s³owników —
    Dim r As Range, nazwa As String, kat As String, minv As Double
    Dim keyDot As String, keyComma As String, firstWord As String, k As String

    For Each r In rng.Columns(COL_NAZWA).Cells
        nazwa = CleanTxt(r.Value)
        kat = CleanTxt(r.Offset(0, COL_KAT - COL_NAZWA).Value)
        If nazwa <> "" And kat <> "" Then
            minv = CDbl(r.Offset(0, COL_MIN - COL_NAZWA).Value)

            ' maks w kategorii
            If dictMax.Exists(kat) Then
                If minv > dictMax(kat) Then dictMax(kat) = minv
            Else
                dictMax(kat) = minv
            End If

            ' exact: orygina³
            dictExact(kat & "|" & nazwa) = minv

            ' exact: warianty kablowe
            keyDot = NormPrzekrojKey(nazwa)                 'np. 5x2.5
            If keyDot <> "" Then
                dictExact(kat & "|" & keyDot) = minv
                keyComma = Replace(keyDot, ".", ",")        '5x2,5
                dictExact(kat & "|" & keyComma) = minv
            End If

            ' exact: 1. s³owo + warianty
            firstWord = Split(nazwa, " ")(0)
            If firstWord <> "" Then
                dictExact(kat & "|" & CleanTxt(firstWord)) = minv
                keyDot = NormPrzekrojKey(firstWord)
                If keyDot <> "" Then
                    dictExact(kat & "|" & keyDot) = minv
                    dictExact(kat & "|" & Replace(keyDot, ".", ",")) = minv
                End If
            End If
        End If
    Next r

    BuildDicts = True
    Exit Function
oops:
    BuildDicts = False
End Function

'=============================================================
' 5) Funkcja arkuszowa – zwraca MINUTY RG (0, gdy brak)
'     • kable: tylko dopasowanie „na przekrój”
'     • inne: maksimum w kategorii
'=============================================================
Public Function Roboczogodziny(kategoria As String, opis As String) As Double
    Static dictExact As Object, dictMax As Object

    If dictExact Is Nothing Then
        If Not BuildDicts(dictExact, dictMax) Then
            Roboczogodziny = 0
            Exit Function
        End If
    End If

    Dim cat As String: cat = CleanTxt(kategoria)
    If cat = "" Then Roboczogodziny = 0: Exit Function

    '— KABLE: tylko dok³adne trafienie (bez fallbacku do maksimum) —
    If InStr(cat, "kabl") > 0 Then
        Dim pr As String: pr = WyodrebnijPrzekroj(opis)
        If pr <> "" Then
            Dim k As String: k = cat & "|" & pr
            If dictExact.Exists(k) Then
                Roboczogodziny = dictExact(k)
                Exit Function
            End If
        End If
        Roboczogodziny = 0
        Exit Function
    End If

    '— INNE KATEGORIE: zwróæ maksimum w danej kategorii —
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


