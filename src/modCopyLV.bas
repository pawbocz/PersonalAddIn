Attribute VB_Name = "modCopyLV"
'========= modCopyLV ===========================================
Option Explicit

'--- kolumna ID jest nadal sztywno w kolumnie A -----------------
Const COL_HID_ID As Long = 1          'A � ukryte ID

'--- domy�lne kolumny / pierwszy wiersz (gdy user anuluje) -------
Const DEF_LP_COL     As Long = 2      'B
Const DEF_OPIS_COL   As Long = 3      'C
Const DEF_PRZEM_COL  As Long = 4      'D
Const DEF_JEDN_COL   As Long = 6      'F
Const DEF_START_ROW  As Long = 8

'--- GLOBALNE ---------------------------------------------------
Public gSourceWB   As Workbook
Public gTargetWB   As Workbook
Public gTemplateLV As Worksheet      'pierwszy arkusz zaczynaj�cy si� od "LV"

'============================  M A I N  =========================
Sub MainCopy()

    Dim wbSrc As Workbook, wbTgt As Workbook
    Dim pathTgt As Variant
    
    If ActiveWorkbook Is Nothing Then Exit Sub
    Set wbSrc = ActiveWorkbook:  Set gSourceWB = wbSrc
    
    pathTgt = Application.GetOpenFilename("Pliki Excel (*.xls*;*.xlsm), *.xls*;*.xlsm")
    If pathTgt = False Then Exit Sub
    
    Set wbTgt = Workbooks.Open(pathTgt)
    Set gTargetWB = wbTgt
    
    Set gTemplateLV = GetTemplateLV(wbTgt)
    If gTemplateLV Is Nothing Then
        MsgBox "W pliku docelowym brak arkusza, kt�rego nazwa zaczyna si� od 'LV'.", _
               vbCritical
        wbTgt.Close False
        Exit Sub
    End If
    
    '�������� 1. dopasowanie arkuszy  ����������������������������
    Load frmSheetMap
    frmSheetMap.Show
    
    If Not frmSheetMap.FormOK Then    'anulowano
        wbTgt.Close False
        Exit Sub
    End If
    
    Dim userHdrRow As Long: userHdrRow = frmSheetMap.hdrRow
    
    '�������� 2. dodatkowe ustawienia kolumn LV ������������������
    Dim mapLp As Long, mapOpis As Long, mapJedn As Long, mapPrzedm As Long
    Dim mapStart As Long
    
    mapLp = frmSheetMap.colLp
    mapOpis = frmSheetMap.colOpis
    mapJedn = frmSheetMap.colJedn
    mapPrzedm = frmSheetMap.colPrzedm
    mapStart = frmSheetMap.startRow
    
    'gdy pole anulowane / puste � wracamy do domy�lnych
    If mapLp = 0 Then mapLp = DEF_LP_COL
    If mapOpis = 0 Then mapOpis = DEF_OPIS_COL
    If mapJedn = 0 Then mapJedn = DEF_JEDN_COL
    If mapPrzedm = 0 Then mapPrzedm = DEF_PRZEM_COL
    If mapStart = 0 Then mapStart = DEF_START_ROW
    
    '���� zapisz list� par do arkusza �Ustawienia� ���������������
    SavePairsToSettings frmSheetMap.pairs, wbTgt
    
    '���� 3. PRE-kopiowanie brakuj�cych arkuszy LV (z szablonu) ��
    Dim pr As Variant
    For Each pr In frmSheetMap.pairs
        If UCase$(pr(1)) <> "SUMA" Then
            If Not SheetExists(wbTgt, pr(1)) Then
                gTemplateLV.Copy After:=wbTgt.Sheets(wbTgt.Sheets.Count)
                wbTgt.Sheets(wbTgt.Sheets.Count).Name = pr(1)
            End If
        End If
    Next pr
    
    '���� 4. w�a�ciwe kopiowanie danych ��������������������������
    Application.ScreenUpdating = False
    
    Dim wsSrc As Worksheet
    For Each pr In frmSheetMap.pairs           'Pair = Array(src, tgt)
        Set wsSrc = wbSrc.Sheets(pr(0))
        CopyOnePair wsSrc, wbTgt, CStr(pr(1)), userHdrRow, _
            mapLp, mapOpis, mapJedn, mapPrzedm, mapStart
    Next pr
    
    Application.ScreenUpdating = True
    
    wbTgt.Activate
    MsgBox "Kopiowanie zako�czone pomy�lnie.", vbInformation
End Sub
'================================================================

'================================================================
'  C O P Y  O N E   P A I R
'  Kopiuje dane z arkusza �r�d�owego (wsSrc) do arkusza LV w
'  skoroszycie docelowym; kolumny i pierwszy wiersz przekazywane
'  jako parametry.
'================================================================
Private Sub CopyOnePair( _
        wsSrc As Worksheet, _
        wbTgt As Workbook, _
        tgtName As String, _
        userHdrRow As Long, _
        ByVal COL_LP_TGT As Long, _
        ByVal COL_OPIS_TGT As Long, _
        ByVal COL_JEDN_TGT As Long, _
        ByVal COL_PRZEM_TGT As Long, _
        ByVal START_ROW_TGT As Long)

    '0) ignoruj arkusz �SUMA�
    If UCase$(tgtName) = "SUMA" Then Exit Sub

    '----------------------------------------------------------------
    '1.  Arkusz docelowy (kopiuj wzorzec LV, je�li jeszcze nie ma)
    '----------------------------------------------------------------
    Dim wsTgt As Worksheet
    On Error Resume Next
    Set wsTgt = wbTgt.Sheets(tgtName)
    On Error GoTo 0
    
    If wsTgt Is Nothing Then
        gTemplateLV.Copy After:=wbTgt.Sheets(wbTgt.Sheets.Count)
        Set wsTgt = wbTgt.Sheets(wbTgt.Sheets.Count)
        On Error Resume Next: wsTgt.Name = tgtName: On Error GoTo 0
    End If
    
    '� upewnij si�, �e w LV istnieje ukryta kolumna A = ID
    EnsureHiddenIDColumn wsTgt

    '----------------------------------------------------------------
    '2.  Ustal wiersz nag��wk�w i pe�ny zakres tabeli w �r�dle
    '----------------------------------------------------------------
    Dim headerRow As Long, tbl As Range
    
    If userHdrRow > 0 Then
        headerRow = userHdrRow
        Dim lastCol As Long, lastRow As Long
        lastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column
        lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        Set tbl = wsSrc.Range(wsSrc.Cells(headerRow, 1), _
                              wsSrc.Cells(lastRow, lastCol))
    Else
        Set tbl = Selection.CurrentRegion
        headerRow = tbl.Row
    End If

    '----------------------------------------------------------------
    '3.  Znajd� bazowe kolumny w �r�dle
    '----------------------------------------------------------------
    Dim idColSrc As Long, opisColSrc As Long, jednColSrc As Long, przedmColSrc As Long
    
    idColSrc = ZnajdzKolumneWRegion(tbl, "ID")
    opisColSrc = ZnajdzKolumneWRegion(tbl, "Opis")
    jednColSrc = ZnajdzKolumneWRegion(tbl, "Jedn.przedm.")
    przedmColSrc = ZnajdzKolumneWRegion(tbl, "Przedmiar")
    
    If idColSrc * opisColSrc * jednColSrc * przedmColSrc = 0 Then
        MsgBox "Brakuje nag��wk�w (ID / Opis / Jedn.przedm. / Przedmiar) w '" & _
               wsSrc.Name & "'.", vbCritical
        Exit Sub
    End If

    '----------------------------------------------------------------
    '4.  Skopiuj wiersze 1-7 z szablonu LV (tytu�y, formaty, �)
    '----------------------------------------------------------------
    If UCase$(Left$(wsTgt.Name, 2)) = "LV" Then
        gTemplateLV.Rows("1:8").Copy
        With wsTgt.Rows("1:8")
            .PasteSpecial xlPasteAllUsingSourceTheme
            .PasteSpecial xlPasteColumnWidths
        End With
        Application.CutCopyMode = False
    End If

    '5a) kolumny A-E (ID, Lp, Opis, Jedn, Przedm.)  � czy�cimy od wiersza 8
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, COL_HID_ID), _
                wsTgt.Cells(wsTgt.Rows.Count, COL_PRZEM_TGT)).ClearContents
    
    '5b) PRAWY blok F-AU � czy�cimy dopiero od wiersza 9,
    '    aby nie skasowa� formu� w wierszu szablonu (wiersz 8)
    Const LAST_COL As Long = 47                 'AU
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT + 1, COL_PRZEM_TGT + 1), _
                wsTgt.Cells(wsTgt.Rows.Count, LAST_COL)).ClearContents
    '----------------------------------------------------------------
    '6.  Pierwszy i ostatni wiersz danych w �r�dle (wg. kol. ID)
    '----------------------------------------------------------------
    Dim firstDataRow As Long
    firstDataRow = headerRow + 1
    Do While firstDataRow <= tbl.Rows(tbl.Rows.Count).Row _
           And Not IsNumeric(wsSrc.Cells(firstDataRow, idColSrc).Value)
        firstDataRow = firstDataRow + 1
    Loop

    Dim lastRowSrc As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, idColSrc).End(xlUp).Row
    If lastRowSrc < firstDataRow Then Exit Sub

    '----------------------------------------------------------------
    '7.  Kopiuj wiersz po wierszu
    '----------------------------------------------------------------
    Dim w As Long: w = START_ROW_TGT
    Dim i As Long
    For i = firstDataRow To lastRowSrc
        If LenB(wsSrc.Cells(i, idColSrc).Value) <> 0 Then
            wsTgt.Cells(w, COL_HID_ID).Value = wsSrc.Cells(i, idColSrc).Value
            wsTgt.Cells(w, COL_LP_TGT).Value = wsSrc.Cells(i, idColSrc + 1).Value
            wsTgt.Cells(w, COL_OPIS_TGT).Value = wsSrc.Cells(i, opisColSrc).Value
            wsTgt.Cells(w, COL_JEDN_TGT).Value = wsSrc.Cells(i, jednColSrc).Value
            wsTgt.Cells(w, COL_PRZEM_TGT).Value = wsSrc.Cells(i, przedmColSrc).Value
            w = w + 1
        End If
    Next i

    '----------------------------------------------------------------
    '8.  Ramki na nowo wstawionych danych
    '----------------------------------------------------------------
    UstawRamkiAll wsTgt.Range(wsTgt.Cells(START_ROW_TGT, COL_LP_TGT), _
                              wsTgt.Cells(w - 1, COL_PRZEM_TGT))

    '----------------------------------------------------------------
    '9.  Formu�y / sumy � tylko w arkuszach LV
    '----------------------------------------------------------------
    If UCase$(Left$(wsTgt.Name, 2)) = "LV" Then AfterPasteLV wsTgt
End Sub
'================================================================




'-------------- PO PASTE � formu�y, sumy, ramki -----------------
Private Sub AfterPasteLV(ByVal ws As Worksheet)
    RozszerzFormulyLV ws
End Sub
'----------------------------------------------------------------

Private Function GetTemplateLV(wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetTemplateLV = wb.Worksheets("LV_SZABLON")   'preferowany
    On Error GoTo 0
    If GetTemplateLV Is Nothing Then                  'fallback
        Dim sh As Worksheet
        For Each sh In wb.Worksheets
            If LCase(sh.Name) Like "lv*" Then
                Set GetTemplateLV = sh
                Exit Function
            End If
        Next sh
    End If
End Function

Private Function SheetExists(wb As Workbook, ByVal nm As String) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Sheets(nm) Is Nothing
    On Error GoTo 0
End Function

'===== w modCopyLV =============================================
Private Sub SavePairsToSettings( _
        ByVal pairs As Collection, _
        ByVal wb As Workbook)

    Const SH_NAME As String = "Ustawienia"
    
    Dim sh As Worksheet
    '--- spr�buj z�apa� istniej�cy arkusz ----------------------
    On Error Resume Next
    Set sh = wb.Worksheets(SH_NAME)
    On Error GoTo 0
    
    '--- je�li nie ma � utw�rz i ukryj -------------------------
    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(Before:=wb.Sheets(1))
        sh.Name = SH_NAME
        sh.Visible = xlSheetVeryHidden       'ukryty, ale nie "VeryHidden"?
        '� xlSheetVeryHidden = ca�kiem niewidoczny w Excelu
        '� xlSheetHidden     = mo�na ods�oni� przez Format ? Unhide
    End If
    
    '--- wyczy�� i wpisz nag��wki ------------------------------
    sh.Cells.Clear
    sh.Range("A1:B1").Value = Array("SourceSheet", "TargetLV")
    sh.Range("A1:B1").Font.Bold = True
    
    '--- zapisuj pary ------------------------------------------
    Dim r As Long: r = 2
    Dim pr As Variant
    For Each pr In pairs
        sh.Cells(r, 1).Value = pr(0)          'arkusz �r�d�owy
        sh.Cells(r, 2).Value = pr(1)          'arkusz LV
        r = r + 1
    Next pr
    
    '--- (opc.) komunikat do Immediate -------------------------
    Debug.Print "SavePairsToSettings:", pairs.Count, "par zapisano do", wb.Name
End Sub


'------------------------------------------------------------------
'  Zapewnia, �e w arkuszu docelowym jest ukryta kolumna A = "ID"
'  � wstawia, je�li jej brak
'  � ustawia szeroko�� 0 i blokuje przed edycj�
'------------------------------------------------------------------
Private Sub EnsureHiddenIDColumn(ws As Worksheet)

    Const HDR_ROW As Long = 4        'wiersz nag��wk�w

    If Trim$(LCase$(ws.Cells(HDR_ROW, 1).Value)) <> "id" Then
        ws.Columns(1).Insert Shift:=xlToRight          'wstaw A
        ws.Cells(HDR_ROW, 1).Value = "ID"              'nag��wek

        With ws.Columns(1)
            .ColumnWidth = 0                           'ukryj
            .Locked = True                             'zablokuj
        End With
    End If
End Sub

