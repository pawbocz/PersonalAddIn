Attribute VB_Name = "DataCopyingScript"
Sub KopiujDaneDoPlikuLV()
    Const START_ROW_TGT As Long = 8        ' <-- zawsze zapis od 8

    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wbTgt As Workbook, wsTgt As Worksheet
    Dim tbl As Range, headerRow As Long
    Dim idColSrc As Long, opisColSrc As Long, jednColSrc As Long, przedmColSrc As Long
    Dim idColTgt As Long, opisColTgt As Long, jednColTgt As Long, przedmColTgt As Long
    Dim lastRowSrc As Long, writeRow As Long, i As Long
    Dim sciezka As Variant
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Zaznacz dowoln� kom�rk� w tabeli �r�d�owej i uruchom makro ponownie.", vbExclamation
        Exit Sub
    End If
    
    Set wbSrc = ActiveWorkbook
    Set gSourceWB = wbSrc
    Set wsSrc = ActiveSheet
    Set tbl = Selection.CurrentRegion
    headerRow = tbl.Row
    
    idColSrc = ZnajdzKolumneWRegion(tbl, "ID")
    opisColSrc = ZnajdzKolumneWRegion(tbl, "Opis")
    jednColSrc = ZnajdzKolumneWRegion(tbl, "Jedn.przedm.")
    przedmColSrc = ZnajdzKolumneWRegion(tbl, "Przedmiar")
    
    If idColSrc * opisColSrc * jednColSrc * przedmColSrc = 0 Then
        MsgBox "Brakuje kt�rego� nag��wka (ID, Opis, Jedn.przedm., Przedmiar).", vbCritical
        Exit Sub
    End If
    
    sciezka = Application.GetOpenFilename("Pliki Excel (*.xls*), *.xls*")
    If sciezka = False Then Exit Sub
    
    Set wbTgt = Workbooks.Open(sciezka)
    On Error Resume Next
    Set wsTgt = wbTgt.Sheets("LV")
    On Error GoTo 0
    If wsTgt Is Nothing Then
        MsgBox "Brak arkusza 'LV' w pliku docelowym.", vbCritical
        wbTgt.Close SaveChanges:=False
        Exit Sub
    End If
    
    DataCopy.Show
    If Not DataCopy.FormOK Then
        wbTgt.Close SaveChanges:=False
        Exit Sub
    End If
    
    idColTgt = Columns(DataCopy.idCol).Column
    opisColTgt = Columns(DataCopy.OpisCol).Column
    jednColTgt = Columns(DataCopy.JednCol).Column
    przedmColTgt = Columns(DataCopy.PrzedmCol).Column
    
    '-- wyczy�� stare dane od wiersza 8 w d� we wskazanych kolumnach
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, idColTgt), wsTgt.Cells(wsTgt.Rows.Count, idColTgt)).ClearContents
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, opisColTgt), wsTgt.Cells(wsTgt.Rows.Count, opisColTgt)).ClearContents
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, jednColTgt), wsTgt.Cells(wsTgt.Rows.Count, jednColTgt)).ClearContents
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, przedmColTgt), wsTgt.Cells(wsTgt.Rows.Count, przedmColTgt)).ClearContents
    
    lastRowSrc = tbl.Rows.Count + headerRow - 1
    writeRow = START_ROW_TGT
    
    For i = headerRow + 1 To lastRowSrc
        If wsSrc.Cells(i, idColSrc).Value <> "" Then
            wsTgt.Cells(writeRow, idColTgt).Value = wsSrc.Cells(i, idColSrc).Value
            wsTgt.Cells(writeRow, opisColTgt).Value = wsSrc.Cells(i, opisColSrc).Value
            wsTgt.Cells(writeRow, jednColTgt).Value = wsSrc.Cells(i, jednColSrc).Value
            wsTgt.Cells(writeRow, przedmColTgt).Value = wsSrc.Cells(i, przedmColSrc).Value
            writeRow = writeRow + 1
        End If
    Next i

'============ NOWA CZʌ� � RAMKI ALL-BORDERS ========================

Dim firstCol As Long, lastCol As Long
Dim pasteRange As Range

' minimalna i maksymalna kolumna skopiowanego zakresu
firstCol = WorksheetFunction.Min(idColTgt, opisColTgt, jednColTgt, przedmColTgt)
lastCol = WorksheetFunction.Max(idColTgt, opisColTgt, jednColTgt, przedmColTgt)

' zakres od wiersza 8 do ostatniego wklejonego (writeRow-1)
Set pasteRange = wsTgt.Range(wsTgt.Cells(START_ROW_TGT, firstCol), _
                              wsTgt.Cells(writeRow - 1, lastCol))

Call UstawRamkiAll(pasteRange)      ' cienkie ci�g�e obramowanie


    wbTgt.Activate
    wsTgt.Activate
    RozszerzFormulyLV wsTgt

    
    
    MsgBox "Dane nadpisane od wiersza " & START_ROW_TGT & ".", vbInformation
    ' wbTgt.Close SaveChanges:=True   ' odkomentuj, je�li chcesz od razu zapisa� i zamkn��
End Sub

'--- pomocnicza ----------------------------------------------------
Function ZnajdzKolumneWRegion(tbl As Range, naglowek As String) As Long
    Dim c As Range
    For Each c In tbl.Rows(1).Cells
        If LCase(Trim(c.Value)) = LCase(naglowek) Then
            ZnajdzKolumneWRegion = c.Column
            Exit Function
        End If
    Next c
    ZnajdzKolumneWRegion = 0
End Function


Sub UstawRamkiAll(rng As Range)

    With rng.Borders
        '--- wszystkie linie wewn�trz -----------------------------
        .LineStyle = xlContinuous
        .Weight = xlThin
        
        '--- kraw�dzie zewn�trzne � wymu� ponownie ----------------
        .Item(xlEdgeLeft).LineStyle = xlContinuous
        .Item(xlEdgeTop).LineStyle = xlContinuous
        .Item(xlEdgeRight).LineStyle = xlContinuous    '<<<< PRAWA
        .Item(xlEdgeBottom).LineStyle = xlContinuous
        
        .Item(xlEdgeLeft).Weight = xlThin
        .Item(xlEdgeTop).Weight = xlThin
        .Item(xlEdgeRight).Weight = xlThin
        .Item(xlEdgeBottom).Weight = xlThin
    End With
End Sub
'====================================================================




'================================================================
Sub PrepareSourceData()

    '---- 1. PARAMETRY Z FORMULARZA -----------------------------
    Dim frm As New frmPrepSettings
    frm.Show
    If Not frm.FormOK Then Exit Sub        'u�ytkownik Cancel
    
    Dim hdrRow       As Long: hdrRow = frm.hdrRow                'wiersz nag��wk�w
    Dim firstDataRow As Long: firstDataRow = frm.FirstData       'pierwszy wiersz danych
    Dim colLp        As Long: colLp = frm.colLp
    Dim colOpis      As Long: colOpis = frm.colOpis
    Dim colJedn      As Long: colJedn = frm.colJedn
    Dim colPrzedm    As Long: colPrzedm = frm.colPrzedm
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    Application.ScreenUpdating = False
    
    '---- 2. WSTAW KOLUMN� A = ID (je�li brak) ------------------
    If LCase$(Trim$(ws.Cells(hdrRow, 1).Value)) <> "id" Then
        ws.Columns(1).Insert Shift:=xlToRight
        ws.Cells(hdrRow, 1).Value = "ID"
        'po wstawieniu kol. A przesuwamy liczniki w prawo
        colLp = colLp + 1: colOpis = colOpis + 1
        colJedn = colJedn + 1: colPrzedm = colPrzedm + 1
    End If
    
    '---- 3. ROZBIJ SCALENIA w nag��wkach -----------------------
    If ws.Cells(hdrRow, colLp).MergeCells Then _
        ws.Cells(hdrRow, colLp).MergeArea.UnMerge
    If ws.Cells(hdrRow, colOpis).MergeCells Then _
        ws.Cells(hdrRow, colOpis).MergeArea.UnMerge
    If ws.Cells(hdrRow, colJedn).MergeCells Then _
        ws.Cells(hdrRow, colJedn).MergeArea.UnMerge
    If ws.Cells(hdrRow, colPrzedm).MergeCells Then _
        ws.Cells(hdrRow, colPrzedm).MergeArea.UnMerge
    
    '---- 4. NADPISZ NAZWY NAG��WK�W ----------------------------
    With ws
        .Cells(hdrRow, colLp).Value = "Lp."
        .Cells(hdrRow, colOpis).Value = "Opis"
        .Cells(hdrRow, colJedn).Value = "Jedn.przedm."
        .Cells(hdrRow, colPrzedm).Value = "Przedmiar"
    End With
    
    '---- 5. USTAL OSTATNI WIERSZ DANYCH ------------------------
    Dim lastRow As Long, r As Long, coreEmpty As Boolean
    r = firstDataRow
    Do
        coreEmpty = _
            LenB(ws.Cells(r, colLp).Value) = 0 And _
            LenB(ws.Cells(r, colOpis).Value) = 0 And _
            LenB(ws.Cells(r, colJedn).Value) = 0 And _
            LenB(ws.Cells(r, colPrzedm).Value) = 0
        If coreEmpty Then Exit Do
        r = r + 1
    Loop
    lastRow = r - 1                       'ostatni rzeczywisty wiersz
    
    If lastRow < firstDataRow Then
        MsgBox "Nie znaleziono danych pod nag��wkami.", vbInformation
        Exit Sub
    End If
    
    '---- 6. NUMERACJA ID (na sta�e, bez formu�y) ---------------
    With ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastRow, 1))
        .Formula = "=ROW()-" & firstDataRow - 1
        .Value = .Value
    End With
    
    '---- 7. WYR�WNANIE + RAMKI ---------------------------------
    With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lastRow, colPrzedm))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .Rows(hdrRow).Font.Bold = True
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Arkusz ustandaryzowany � ID ponumerowane do wiersza " & lastRow & ".", vbInformation
End Sub
'================================================================


'===========  P R Z E N O S Z E N I E   K O L U M N Y  =============
Private Sub MoveCol(ws As Worksheet, srcIndex As Long, tgtIndex As Long)
    If srcIndex = tgtIndex Then Exit Sub
    ws.Columns(srcIndex).Cut
    ws.Columns(tgtIndex).Insert Shift:=xlToRight
End Sub


