Attribute VB_Name = "modCopyLV"
'========  modCopyLV  ==========================================
Option Explicit

'--- kolumna ID jest zawsze ukryta w kolumnie A -----------------
Const COL_HID_ID As Long = 1          'A

'--- domy�lne mapowanie kolumn LV (po dodaniu ID) ---------------
Const DEF_LP_COL     As Long = 2      'B
Const DEF_OPIS_COL   As Long = 3      'C
Const DEF_PRZEM_COL  As Long = 4      'D      ��� PRZEDMIAR
Const DEF_JEDN_COL   As Long = 6      'F      ��� JEDN.przedm.
Const DEF_START_ROW  As Long = 8      'pierwszy wiersz danych w LV

'--- GLOBALNE ---------------------------------------------------
Public gSourceWB   As Workbook        'plik inwestorski
Public gTargetWB   As Workbook        'plik LV
Public gTemplateLV As Worksheet       'pierwszy arkusz �LV�� lub LV_SZABLON
'===============================================================


'=== HELPER: wgrywa top 1..8 z szablonu bez schowka, plus szer. i wys. ===
Private Sub ApplyTemplateTop8(ByVal wsTgt As Worksheet)
    'zakres: A1:AU8 (47 kolumn)
    gTemplateLV.Range("A1:AU8").Copy Destination:=wsTgt.Range("A1")

    'szeroko�ci kolumn 1..47 (A:AU)
    Dim c As Long
    For c = 1 To 47
        wsTgt.Columns(c).ColumnWidth = gTemplateLV.Columns(c).ColumnWidth
    Next c

    'wysoko�ci wierszy 1..8
    Dim r As Long
    For r = 1 To 8
        wsTgt.Rows(r).RowHeight = gTemplateLV.Rows(r).RowHeight
    Next r
End Sub




'============================  M A I N  =========================
Sub MainCopy()

    '��� 0. wyb�r pliku docelowego LV ���������������������������
    Set gSourceWB = ActiveWorkbook
    If gSourceWB Is Nothing Then Exit Sub
    
    Dim pathTgt As Variant
    pathTgt = Application.GetOpenFilename( _
              "Pliki Excel (*.xls*;*.xlsm;*.xltx;*.xltm), *.xls*;*.xlsm;*.xltx;*.xltm")
    If pathTgt = False Then Exit Sub
    
    Set gTargetWB = Workbooks.Open(pathTgt)
    Set gTemplateLV = GetTemplateLV(gTargetWB)
    If gTemplateLV Is Nothing Then
        MsgBox "W pliku docelowym brak arkusza, kt�rego nazwa zaczyna si� od 'LV'.", _
               vbCritical
        gTargetWB.Close False
        Exit Sub
    End If
    
    '��� 1. formularz mapowania i (opc.) kolumn LV ��������������
    Load frmSheetMap
    frmSheetMap.Show
    
    If Not frmSheetMap.FormOK Then            'anulowano formularz
        gTargetWB.Close False
        Exit Sub
    End If
    
    Dim userHdrRow As Long: userHdrRow = frmSheetMap.hdrRow
    
    '------ mapowanie kolumn docelowych LV -----------------------
    Dim mapLp As Long, mapOpis As Long, mapJedn As Long
    Dim mapPrzedm As Long, mapStart As Long
    
    If frmSheetMap.UseCustomCols Then
        mapLp = val(frmSheetMap.txtLp.Text)
        mapOpis = val(frmSheetMap.txtOpis.Text)
        mapJedn = val(frmSheetMap.txtJedn.Text)
        mapPrzedm = val(frmSheetMap.txtPrzedm.Text)
        mapStart = val(frmSheetMap.txtStart.Text)
    Else
        mapLp = DEF_LP_COL
        mapOpis = DEF_OPIS_COL
        mapJedn = DEF_JEDN_COL
        mapPrzedm = DEF_PRZEM_COL
        mapStart = DEF_START_ROW
    End If
    
    'gdy kt�rego� pola brak / zero � wska� domy�lne
    If mapLp = 0 Then mapLp = DEF_LP_COL
    If mapOpis = 0 Then mapOpis = DEF_OPIS_COL
    If mapJedn = 0 Then mapJedn = DEF_JEDN_COL
    If mapPrzedm = 0 Then mapPrzedm = DEF_PRZEM_COL
    If mapStart = 0 Then mapStart = DEF_START_ROW
    
    '��� 2. zapis par arkuszy do �Ustawienia� �������������������
    SavePairsToSettings frmSheetMap.pairs, gTargetWB
    
    '��� 3. PRE-kopiowanie brakuj�cych LV z szablonu ������������
    Dim pr As Variant
    For Each pr In frmSheetMap.pairs
        If UCase$(pr(1)) <> "SUMA" Then
            If Not SheetExists(gTargetWB, pr(1)) Then
                gTemplateLV.Copy After:=gTargetWB.Sheets(gTargetWB.Sheets.Count)
                gTargetWB.Sheets(gTargetWB.Sheets.Count).Name = pr(1)
            End If
        End If
    Next pr
    
    '��� 4. w�a�ciwe kopiowanie danych ��������������������������
    Application.ScreenUpdating = False
    
    Dim wsSrc As Worksheet
    For Each pr In frmSheetMap.pairs            'pr = Array(srcName, tgtName)
        Set wsSrc = gSourceWB.Sheets(pr(0))
        CopyOnePair wsSrc, gTargetWB, CStr(pr(1)), userHdrRow, _
                     mapLp, mapOpis, mapJedn, mapPrzedm, mapStart
    Next pr
    
    Application.ScreenUpdating = True
    
    gTargetWB.Activate
    MsgBox "Kopiowanie zako�czone pomy�lnie.", vbInformation
End Sub
'===============================================================

'================================================================
'  C O P Y  O N E   P A I R
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


    '��������������������������������������������������������������
    '1.  Arkusz docelowy (kopiuj wzorzec LV, je�li nie istnieje)
    '��������������������������������������������������������������
    Dim wsTgt As Worksheet, newSheet As Boolean
    On Error Resume Next
    Set wsTgt = wbTgt.Worksheets(tgtName)
    On Error GoTo 0

    If wsTgt Is Nothing Then               '�� tworzymy z szablonu ��
        gTemplateLV.Copy After:=wbTgt.Sheets(wbTgt.Sheets.Count)
        Set wsTgt = wbTgt.Sheets(wbTgt.Sheets.Count)
        On Error Resume Next: wsTgt.Name = tgtName: On Error GoTo 0
        newSheet = True                    '��flaga: to �wie�y arkusz
    End If

    'Je�li to NIE jest �wie�o skopiowany arkusz, upewnij si�, �e ma
    'ukryt� kolumn� A=ID.  Szablon ma j� ju� w sobie, wi�c pomijamy.
    


    '��������������������������������������������������������������
    '2.  Ustal wiersz nag��wk�w i pe�ny zakres tabeli w �r�dle
    '��������������������������������������������������������������
    Dim headerRow As Long, tbl As Range
    If userHdrRow > 0 Then
        headerRow = userHdrRow
        Dim lastCol As Long, lastRow As Long
        lastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column
        lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        Set tbl = wsSrc.Range(wsSrc.Cells(headerRow, 1), wsSrc.Cells(lastRow, lastCol))
    Else
        Set tbl = Selection.CurrentRegion
        headerRow = tbl.Row
    End If


    '��������������������������������������������������������������
    '3.  Znajd� bazowe kolumny w �r�dle
    '��������������������������������������������������������������
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


    '4) bezwarunkowo wgraj top z szablonu (nie polegaj na schowku)
    If UCase$(Left$(wsTgt.Name, 2)) = "LV" Then
        'opcjonalna os�ona: nie nadpisuj samego szablonu
        If wsTgt.CodeName <> gTemplateLV.CodeName Then
            ApplyTemplateTop8 wsTgt
        End If
    End If
    
    '*** TERAZ zadbaj o kolumn� ID (po wgraniu topu) ***
    EnsureHiddenIDColumn wsTgt
    


    '��������������������������������������������������������������
    '5.  Czyszczenie starych danych
    '��������������������������������������������������������������
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT + 1, COL_HID_ID), _
            wsTgt.Cells(wsTgt.Rows.Count, COL_PRZEM_TGT)).ClearContents

    Const LAST_COL As Long = 47           'AU
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT + 1, COL_PRZEM_TGT + 1), _
                wsTgt.Cells(wsTgt.Rows.Count, LAST_COL)).ClearContents


    '��������������������������������������������������������������
    '6.  Pierwszy-ostatni�wiersz danych
    '��������������������������������������������������������������
    Dim firstDataRow As Long: firstDataRow = headerRow + 1
    Do While firstDataRow <= tbl.Rows(tbl.Rows.Count).Row _
           And Not IsNumeric(wsSrc.Cells(firstDataRow, idColSrc).Value)
        firstDataRow = firstDataRow + 1
    Loop

    Dim lastRowSrc As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, idColSrc).End(xlUp).Row
    If lastRowSrc < firstDataRow Then Exit Sub


    '��������������������������������������������������������������
    '7.  Przenoszenie danych
    '��������������������������������������������������������������
    Dim i As Long, w As Long: w = START_ROW_TGT
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


    '��������������������������������������������������������������
    '8.  Ramki (na nowo skopiowanych danych)
    '��������������������������������������������������������������
    UstawRamkiAll wsTgt.Range(wsTgt.Cells(START_ROW_TGT, COL_LP_TGT), _
                              wsTgt.Cells(w - 1, COL_PRZEM_TGT))


    '��������������������������������������������������������������
    '9.  LV-specyficzne formu�y / sumy
    '��������������������������������������������������������������
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

