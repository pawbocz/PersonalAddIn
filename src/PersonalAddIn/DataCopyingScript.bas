Attribute VB_Name = "DataCopyingScript"
Sub KopiujDaneDoPlikuLV()
    Const START_ROW_TGT As Long = 8

    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wbTgt As Workbook, wsTgt As Worksheet
    Dim tbl As Range, headerRow As Long
    Dim idColSrc As Long, opisColSrc As Long, jednColSrc As Long, przedmColSrc As Long
    Dim idColTgt As Long, opisColTgt As Long, jednColTgt As Long, przedmColTgt As Long
    Dim lastRowSrc As Long, writeRow As Long, i As Long
    Dim sciezka As Variant
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Zaznacz dowoln¹ komórkê w tabeli Ÿród³owej i uruchom makro ponownie.", vbExclamation
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
        MsgBox "Brakuje któregoœ nag³ówka (ID, Opis, Jedn.przedm., Przedmiar).", vbCritical
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
    

    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, idColTgt), wsTgt.Cells(wsTgt.Rows.Count, idColTgt)).ClearContents
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, opisColTgt), wsTgt.Cells(wsTgt.Rows.Count, opisColTgt)).ClearContents
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, jednColTgt), wsTgt.Cells(wsTgt.Rows.Count, jednColTgt)).ClearContents
    wsTgt.Range(wsTgt.Cells(START_ROW_TGT, przedmColTgt), wsTgt.Cells(wsTgt.Rows.Count, przedmColTgt)).ClearContents
    
    lastRowSrc = tbl.Rows.Count + headerRow - 1
    writeRow = START_ROW_TGT
    
    For i = headerRow + 1 To lastRowSrc
        If wsSrc.Cells(i, idColSrc).value <> "" Then
            wsTgt.Cells(writeRow, idColTgt).value = wsSrc.Cells(i, idColSrc).value
            wsTgt.Cells(writeRow, opisColTgt).value = wsSrc.Cells(i, opisColSrc).value
            wsTgt.Cells(writeRow, jednColTgt).value = wsSrc.Cells(i, jednColSrc).value
            wsTgt.Cells(writeRow, przedmColTgt).value = wsSrc.Cells(i, przedmColSrc).value
            writeRow = writeRow + 1
        End If
    Next i



Dim firstCol As Long, lastCol As Long
Dim pasteRange As Range


firstCol = WorksheetFunction.min(idColTgt, opisColTgt, jednColTgt, przedmColTgt)
lastCol = WorksheetFunction.Max(idColTgt, opisColTgt, jednColTgt, przedmColTgt)

Set pasteRange = wsTgt.Range(wsTgt.Cells(START_ROW_TGT, firstCol), _
                              wsTgt.Cells(writeRow - 1, lastCol))

Call UstawRamkiAll(pasteRange)


    wbTgt.Activate
    wsTgt.Activate
    RozszerzFormulyLV wsTgt

    
    
    MsgBox "Dane nadpisane od wiersza " & START_ROW_TGT & ".", vbInformation

End Sub


Function ZnajdzKolumneWRegion(tbl As Range, naglowek As String) As Long
    Dim c As Range
    For Each c In tbl.Rows(1).Cells
        If LCase(Trim(c.value)) = LCase(naglowek) Then
            ZnajdzKolumneWRegion = c.Column
            Exit Function
        End If
    Next c
    ZnajdzKolumneWRegion = 0
End Function


Sub UstawRamkiAll(rng As Range)

    With rng.Borders
 
        .LineStyle = xlContinuous
        .Weight = xlThin
        
     
        .Item(xlEdgeLeft).LineStyle = xlContinuous
        .Item(xlEdgeTop).LineStyle = xlContinuous
        .Item(xlEdgeRight).LineStyle = xlContinuous
        .Item(xlEdgeBottom).LineStyle = xlContinuous
        
        .Item(xlEdgeLeft).Weight = xlThin
        .Item(xlEdgeTop).Weight = xlThin
        .Item(xlEdgeRight).Weight = xlThin
        .Item(xlEdgeBottom).Weight = xlThin
    End With
End Sub

Sub PrepareSourceData()

    '---- 1. PARAMETRY Z FORMULARZA -----------------------------
    Dim frm As New frmPrepSettings
    frm.Show
    If Not frm.FormOK Then Exit Sub
    
    Dim hdrRow       As Long: hdrRow = frm.hdrRow
    Dim firstDataRow As Long: firstDataRow = frm.FirstData
    Dim colLp        As Long: colLp = frm.colLp
    Dim colOpis      As Long: colOpis = frm.colOpis
    Dim colJedn      As Long: colJedn = frm.colJedn
    Dim colPrzedm    As Long: colPrzedm = frm.colPrzedm
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    Application.ScreenUpdating = False
    
    '---- 2. WSTAW KOLUMNÊ A = ID (jeœli brak) ------------------
    If LCase$(Trim$(ws.Cells(hdrRow, 1).value)) <> "id" Then
        ws.Columns(1).Insert Shift:=xlToRight
        ws.Cells(hdrRow, 1).value = "ID"
    
        colLp = colLp + 1: colOpis = colOpis + 1
        colJedn = colJedn + 1: colPrzedm = colPrzedm + 1
    End If
    
    '---- 3. ROZBIJ SCALENIA w nag³ówkach -----------------------
    If ws.Cells(hdrRow, colLp).MergeCells Then _
        ws.Cells(hdrRow, colLp).MergeArea.UnMerge
    If ws.Cells(hdrRow, colOpis).MergeCells Then _
        ws.Cells(hdrRow, colOpis).MergeArea.UnMerge
    If ws.Cells(hdrRow, colJedn).MergeCells Then _
        ws.Cells(hdrRow, colJedn).MergeArea.UnMerge
    If ws.Cells(hdrRow, colPrzedm).MergeCells Then _
        ws.Cells(hdrRow, colPrzedm).MergeArea.UnMerge
    
    '---- 4. NADPISZ NAZWY NAG£ÓWKÓW ----------------------------
    With ws
        .Cells(hdrRow, colLp).value = "Lp."
        .Cells(hdrRow, colOpis).value = "Opis"
        .Cells(hdrRow, colJedn).value = "Jedn.przedm."
        .Cells(hdrRow, colPrzedm).value = "Przedmiar"
    End With
    
    '---- 5. USTAL OSTATNI WIERSZ DANYCH ------------------------
    Dim lastRow As Long, r As Long, coreEmpty As Boolean
    r = firstDataRow
    Do
        coreEmpty = _
            LenB(ws.Cells(r, colLp).value) = 0 And _
            LenB(ws.Cells(r, colOpis).value) = 0 And _
            LenB(ws.Cells(r, colJedn).value) = 0 And _
            LenB(ws.Cells(r, colPrzedm).value) = 0
        If coreEmpty Then Exit Do
        r = r + 1
    Loop
    lastRow = r - 1
    
    If lastRow < firstDataRow Then
        MsgBox "Nie znaleziono danych pod nag³ówkami.", vbInformation
        Exit Sub
    End If
    
    '---- 6. NUMERACJA ID (na sta³e, bez formu³y) ---------------
    With ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastRow, 1))
        .Formula = "=ROW()-" & firstDataRow - 1
        .value = .value
    End With
    
    '---- 7. WYRÓWNANIE + RAMKI ---------------------------------
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
    MsgBox "Arkusz ustandaryzowany – ID ponumerowane do wiersza " & lastRow & ".", vbInformation
End Sub


Private Sub MoveCol(ws As Worksheet, srcIndex As Long, tgtIndex As Long)
    If srcIndex = tgtIndex Then Exit Sub
    ws.Columns(srcIndex).Cut
    ws.Columns(tgtIndex).Insert Shift:=xlToRight
End Sub


