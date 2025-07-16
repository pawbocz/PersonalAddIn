Attribute VB_Name = "modRG"
'=======================  modRG  ==========================
Option Explicit

Private dictRG As Object
'----------- arkusz i kolumny s³ownika RG -----------------
Private Const SHEET_STAWKI     As String = "Stawki"
Private Const COL_STAWKI_NAZWA As Long = 1    'A
Private Const COL_STAWKI_CAT   As Long = 2    'B
Private Const COL_STAWKI_MIN   As Long = 3    'C

'----------- kolumny / wiersze w LV -----------------------
Private Const OUT_COL   As Long = 19  'S (roboczogodziny)
Private Const CAT_COL   As String = "AD"
Private Const DESC_COL  As String = "C"
Private Const FIRST_ROW As Long = 8   'pierwszy wiersz danych

'----------------------------------------------------------
' 1. WY£USKAJ przekrój z opisu
'     – obs³uguje tak¿e warianty typu 4x5x10  ?  5x10
'----------------------------------------------------------
Public Function WyodrebnijPrzekroj(opis As String) As String
    
    Dim re As Object, mc As Object
    
    '––– 1) najpierw szukamy wzoru  n×m×k  -------------------
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.Pattern = "\b(\d+)\s*x\s*(\d+)\s*x\s*(\d+(?:,\d+)?)\b"   'np. 4x5x10
    If re.test(opis) Then
        Set mc = re.Execute(opis)(0)
        WyodrebnijPrzekroj = LCase$(Replace( _
                              Replace(mc.SubMatches(1) & "x" & mc.SubMatches(2), _
                                      " ", ""), ",", "."))     '› 5x10
        Exit Function
    End If
    
    '––– 2) standard: n×m lub dnXX  ---------------------------
    re.Pattern = "(\d+\s*x\s*\d+(?:,\d+)?)|(\bdn\d+\b)"
    If re.test(opis) Then
        WyodrebnijPrzekroj = LCase$( _
            Replace(Replace(re.Execute(opis)(0), " ", ""), ",", "."))
    Else
        WyodrebnijPrzekroj = ""
    End If
End Function

'----------------------------------------------------------
'  Roboczogodziny – zwraca MINUTY (0, gdy brak trafienia)
'----------------------------------------------------------
Public Function Roboczogodziny( _
        kategoria As String, opis As String) As Variant

    '––––– 0. s³ownik w pamiêci (cache) –––––
    If dictRG Is Nothing Then
        Set dictRG = CreateObject("Scripting.Dictionary")
        
        'sk¹d wzi¹æ arkusz „Stawki”
        Dim wb As Workbook
        If TypeName(Application.Caller) = "Range" Then
            Set wb = Application.Caller.Parent.Parent        'formu³a w arkuszu
        Else
            Set wb = ActiveWorkbook                          'wywo³anie z VBA
        End If
        
        'tabela Stawki (zak³adamy 1-sza tabela na arkuszu)
        Dim lo As ListObject
        On Error Resume Next
        Set lo = wb.Worksheets(SHEET_STAWKI).ListObjects(1)
        On Error GoTo 0
        If lo Is Nothing Then
            Roboczogodziny = CVErr(xlErrRef): Exit Function
        End If
        
        'wczytaj wszystkie wiersze do s³ownika
        Dim rw As ListRow, nazwa As String, kat As String, val As Variant
        For Each rw In lo.ListRows
            nazwa = LCase$(Trim$(rw.Range(1, COL_STAWKI_NAZWA).Value))   'A
            kat = LCase$(Trim$(rw.Range(1, COL_STAWKI_CAT).Value))       'B
            val = rw.Range(1, COL_STAWKI_MIN).Value                      'C
            If Len(nazwa) * Len(kat) > 0 Then
                dictRG(kat & "|" & nazwa) = val                          'np kable|5x25
            End If
        Next rw
    End If
    '––––– koniec jednorazowego ³adowania s³ownika –––––

    '1. wyci¹gnij przekrój z opisu
    Dim przekroj As String: przekroj = WyodrebnijPrzekroj(opis)
    If przekroj = "" Then Roboczogodziny = 0: Exit Function
    
    '2. zbuduj klucz i zwróæ minuty
    Dim klucz As String
    klucz = LCase$(Trim$(kategoria)) & "|" & LCase$(przekroj)
    
    If dictRG.Exists(klucz) Then
        Roboczogodziny = dictRG(klucz)     'minuty
    Else
        Roboczogodziny = 0
    End If
End Function


'----------------------------------------------------------
' 3.  Wstaw formu³y do kolumny S we wszystkich arkuszach LV
'----------------------------------------------------------
Public Sub WstawFormulyRG()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 2) = "LV" Then
            Dim lastR As Long
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastR < FIRST_ROW Then GoTo NextWs
            
            With ws.Cells(FIRST_ROW, OUT_COL)
                .Formula = "=Roboczogodziny($" & CAT_COL & FIRST_ROW & ",$" & _
                                          DESC_COL & FIRST_ROW & ")"
                .AutoFill ws.Range(ws.Cells(FIRST_ROW, OUT_COL), _
                                   ws.Cells(lastR, OUT_COL))
            End With
        End If
NextWs:
    Next ws
    
    MsgBox "Roboczogodziny wpisane do kolumny S w pliku: " & wb.Name, _
           vbInformation
End Sub
'==========================================================


Sub IleMamWpisów()
    If dictRG Is Nothing Then
        MsgBox "dictRG = Nothing  (nie by³ budowany)", vbExclamation
    Else
        MsgBox "Wpisów w dictRG: " & dictRG.Count, vbInformation
    End If
End Sub

Sub Debug10Keys()
    Dim k As Variant, i As Long
    For Each k In dictRG.Keys
        Debug.Print k, dictRG(k)     'klucz, minuty
        i = i + 1: If i = 10 Then Exit For
    Next k
End Sub

