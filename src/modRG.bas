Attribute VB_Name = "modRG"
'==========================  modRG  ==========================
Option Explicit

'–––– sta³e dla arkusza „Stawki” –––––––––––––––––––––––––––––
Private Const SHEET_STAWKI As String = "Stawki"
Private Const COL_NAZWA    As Long = 1       'A  (np. 5x25, K600 …)
Private Const COL_KAT      As Long = 2       'B  (nazwa kategorii)
Private Const COL_MIN      As Long = 3       'C  (minuty RG)

'=============================================================
' 1. Funkcja pomocnicza  –  szybka „normalizacja” tekstu
'=============================================================
Private Function CleanTxt(s As String) As String
    CleanTxt = LCase$(Trim$(s))
End Function

'=============================================================
' 2. Wyodrêbnienie przekroju kabla  (obs³uguje 4×5×10 › 5×10)
'=============================================================
Public Function WyodrebnijPrzekroj(opis As String) As String
    Dim re3 As Object: Set re3 = CreateObject("VBScript.RegExp")
    re3.Pattern = "^\s*\d+\s*x\s*(\d+\s*x\s*\d+(\,\d+)?)"
    re3.IgnoreCase = True
    If re3.test(opis) Then
        WyodrebnijPrzekroj = LCase$(Replace(re3.Execute(opis)(0).SubMatches(0), " ", ""))
        Exit Function
    End If

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d+\s*x\s*\d+(\,\d+)?)|(\bdn\d+\b)"
    re.IgnoreCase = True
    If re.test(opis) Then
        WyodrebnijPrzekroj = LCase$( _
            Replace(Replace(re.Execute(opis)(0), " ", ""), ",", "."))
    Else
        WyodrebnijPrzekroj = ""
    End If
End Function

'=============================================================
' 3. Budowa s³owników:
'    • dictRG      – wszystkie wpisy "kat|nazwa"  (szukamy tu TYLKO kabli)
'    • dictMaxCat  – maksymalna wartoœæ w ka¿dej kategorii
'=============================================================
Private Function BuildDicts(ByRef dExact As Object, ByRef dMax As Object) As Boolean
    Set dExact = CreateObject("Scripting.Dictionary")
    Set dMax = CreateObject("Scripting.Dictionary")

    '–– wybór skoroszytu z arkuszem „Stawki” ––
    Dim wb As Workbook
    If Application.Caller Is Nothing Then
        Set wb = ActiveWorkbook
    Else
        Set wb = Application.Caller.Parent.Parent
    End If
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_STAWKI)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    '–– zakres danych ––
    Dim rng As Range
    If ws.ListObjects.Count > 0 Then
        Set rng = ws.ListObjects(1).DataBodyRange
    Else
        Dim lr As Long: lr = ws.Cells(ws.Rows.Count, COL_NAZWA).End(xlUp).Row
        If lr < 2 Then Exit Function
        Set rng = ws.Range(ws.Cells(2, COL_NAZWA), ws.Cells(lr, COL_MIN))
    End If

    '–– pêtla po wierszach ––
    Dim r As Range, nm As String, cat As String, v As Double
    For Each r In rng.Columns(COL_NAZWA).Cells
        nm = CleanTxt(r.Value)
        cat = CleanTxt(r.Offset(0, COL_KAT - COL_NAZWA).Value)
        If nm = "" Or cat = "" Then GoTo NextR

        v = CDbl(r.Offset(0, COL_MIN - COL_NAZWA).Value)
        dExact(cat & "|" & nm) = v                      'dok³adna nazwa

        If Not dMax.Exists(cat) Or v > dMax(cat) Then dMax(cat) = v
NextR:
    Next r

    BuildDicts = (dExact.Count > 0)
End Function

'=============================================================
' 4. Funkcja arkuszowa – zwraca MINUTY RG (0, gdy brak)
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

    '–– 1) kable – próba dok³adnego przekroju ––
    If InStr(cat, "kabl") > 0 Then
        Dim pr As String: pr = WyodrebnijPrzekroj(opis)
        If pr <> "" Then
            Dim k As String: k = cat & "|" & pr
            If dictExact.Exists(k) Then
                Roboczogodziny = dictExact(k)
                Exit Function
            End If
        End If
    End If

    '–– 2) pozosta³e – zwracamy najwy¿sz¹ wartoœæ w kategorii ––
    If dictMax.Exists(cat) Then
        Roboczogodziny = dictMax(cat)
    Else
        Roboczogodziny = 0
    End If
End Function

'=============================================================
' 5. Makro wstawiaj¹ce formu³y + koloruj¹ce braki
'=============================================================
Public Sub WstawFormulyRG()

    Const COL_OUT   As Long = 19            'S  (Roboczogodziny)
    Const COL_CAT   As String = "AD"        'kolumna kategorii
    Const COL_DESC  As String = "C"         'kolumna opisu
    Const FIRST_ROW As Long = 8
    Const BRAK_RG_COLOR As Long = &HFF      'RGB(255,0,0)

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If Left$(ws.Name, 2) = "LV" Then                     'obs³ugujemy tylko LV

            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow < FIRST_ROW Then GoTo NextWs          'brak danych

            Dim r As Long, adrCat As String, adrDesc As String, f As String
            For r = FIRST_ROW To lastRow

                'adresy bez $ (wzglêdne) – u³atwia kopiowanie/FillDown
                adrCat = ws.Cells(r, COL_CAT).Address(False, False)
                adrDesc = ws.Cells(r, COL_DESC).Address(False, False)
                f = "=IFERROR(Roboczogodziny(" & adrCat & "," & adrDesc & "),0)"

                With ws.Cells(r, COL_OUT)
                    '-------------------------------
                    '1) Wpisz formu³ê TYLKO gdy:
                    '   • brak formu³y ORAZ
                    '   • komórka pusta albo 0
                    '-------------------------------
                    If (Not .HasFormula) And _
                       (Len(.Value) = 0 Or .Value = 0) Then
                        .Formula = f
                    End If

                    '-------------------------------
                    '2) Koloruj braki RG
                    '-------------------------------
                    If .HasFormula Then
                        If ws.Cells(r, COL_CAT).Value <> "" And .Value = 0 Then
                            .Interior.Color = BRAK_RG_COLOR      'brak trafienia
                        ElseIf .Interior.Color = BRAK_RG_COLOR Then
                            .Interior.Pattern = xlNone           ' RG ju¿ jest › zdejmij
                        End If
                    End If
                End With
            Next r
        End If
NextWs:
    Next ws

    MsgBox "Formu³y RG dodane (tylko do pustych komórek). " & _
           "Braki oznaczone czerwono.", vbInformation
End Sub
'=============================================================
