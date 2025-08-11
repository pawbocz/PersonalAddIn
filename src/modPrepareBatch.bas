Attribute VB_Name = "modPrepareBatch"
Sub PrepareSourceData_All()

    '---- 1. pobierz ustawienia z formularza --------------------
    Dim frm As New frmPrepSettings
    frm.Show
    If Not frm.FormOK Then Exit Sub
    
    Dim hdrRow       As Long: hdrRow = frm.hdrRow
    Dim firstDataRow As Long: firstDataRow = frm.FirstData
    Dim colLp        As Long: colLp = frm.colLp
    Dim colOpis      As Long: colOpis = frm.colOpis
    Dim colJedn      As Long: colJedn = frm.colJedn
    Dim colPrzedm    As Long: colPrzedm = frm.colPrzedm
    
    '---- 2. które arkusze? -------------------------------------
    Dim sel As Collection: Set sel = New Collection
    Dim ans As VbMsgBoxResult
    
    ans = MsgBox("Przetworzyæ wszystkie arkusze w skoroszycie?" & vbCrLf & _
                 "(Nie = poka¿ listê i wybierz)", _
                 vbYesNoCancel + vbQuestion, "PrepareSourceData – batch")
    If ans = vbCancel Then Exit Sub
    
    Dim sh As Worksheet
    If ans = vbYes Then
        For Each sh In ActiveWorkbook.Worksheets
            sel.Add sh
        Next sh
    Else
        Dim frmS As New frmSelectSheets
        frmS.Init ActiveWorkbook.Worksheets
        frmS.Show
        If Not frmS.FormOK Then Exit Sub
        For Each sh In frmS.SelectedSheets
            sel.Add sh
        Next sh
    End If
    
    '---- 3. przetwarzaj wybrane arkusze ------------------------
    Application.ScreenUpdating = False
    Dim okCnt As Long, errCnt As Long
    
    For Each sh In sel
        On Error Resume Next
        PrepOneSheet sh, hdrRow, firstDataRow, _
                     colLp, colOpis, colJedn, colPrzedm
        If Err.Number = 0 Then
            okCnt = okCnt + 1
        Else
            errCnt = errCnt + 1
            Debug.Print "PrepareSourceData_All › b³¹d w", sh.Name, Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next sh
    Application.ScreenUpdating = True
    
    MsgBox "Gotowe!" & vbCrLf & _
           "Ustandaryzowano: " & okCnt & vbCrLf & _
           "Pominiêto (b³¹d): " & errCnt, vbInformation
End Sub

Private Sub PrepOneSheet(ws As Worksheet, _
                         ByVal hdrRow As Long, ByVal firstDataRow As Long, _
                         ByVal colLp As Long, ByVal colOpis As Long, _
                         ByVal colJedn As Long, ByVal colPrzedm As Long)


    '--- 1. kolumna A = ID (jeœli brak) -------------------------
    If LCase$(Trim$(ws.Cells(hdrRow, 1).Value)) <> "id" Then
        ws.Columns(1).Insert Shift:=xlToRight
        ws.Cells(hdrRow, 1).Value = "ID"
        colLp = colLp + 1: colOpis = colOpis + 1
        colJedn = colJedn + 1: colPrzedm = colPrzedm + 1
    End If
    
    '--- 2. unMerge + nag³ówki ----------------------------------
    If ws.Cells(hdrRow, colLp).MergeCells Then _
        ws.Cells(hdrRow, colLp).MergeArea.UnMerge
    If ws.Cells(hdrRow, colOpis).MergeCells Then _
        ws.Cells(hdrRow, colOpis).MergeArea.UnMerge
    If ws.Cells(hdrRow, colJedn).MergeCells Then _
        ws.Cells(hdrRow, colJedn).MergeArea.UnMerge
    If ws.Cells(hdrRow, colPrzedm).MergeCells Then _
        ws.Cells(hdrRow, colPrzedm).MergeArea.UnMerge
    
    With ws
        .Cells(hdrRow, colLp).Value = "Lp."
        .Cells(hdrRow, colOpis).Value = "Opis"
        .Cells(hdrRow, colJedn).Value = "Jedn.przedm."
        .Cells(hdrRow, colPrzedm).Value = "Przedmiar"
    End With
    
    '--- 3. ostatni wiersz danych (pierwsza pusta linia) --------
    Dim r As Long: r = firstDataRow
    Do While LenB(ws.Cells(r, colLp).Value) <> 0 Or _
             LenB(ws.Cells(r, colOpis).Value) <> 0 Or _
             LenB(ws.Cells(r, colJedn).Value) <> 0 Or _
             LenB(ws.Cells(r, colPrzedm).Value) <> 0
        r = r + 1
    Loop
    Dim lastRow As Long: lastRow = r - 1
    If lastRow < firstDataRow Then Exit Sub
    
    '--- 4. numeracja ID ---------------------------------------
    With ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastRow, 1))
        .Formula = "=ROW()-" & firstDataRow - 1
        .Value = .Value
    End With
    
    '--- 5. wyrównanie + ramki ---------------------------------
    With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lastRow, colPrzedm))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .Rows(hdrRow).Font.Bold = True
    End With
End Sub


