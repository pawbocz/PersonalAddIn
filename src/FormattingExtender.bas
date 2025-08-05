Attribute VB_Name = "FormattingExtender"

Sub RozszerzFormulyLV(ByVal wsLV As Worksheet)

    If Left$(wsLV.Name, 2) <> "LV" Then Exit Sub        'obs³uga tylko LV

    Const START_ROW As Long = 8
    Const FIRST_COL As Long = 7          'G   (F+1)
    Const LAST_COL  As Long = 48         'AV  (AU+1)
    Const ID_COL    As Long = 1          'A   (ukryte ID)

    '--- ustal ostatni wiersz danych wg kol. ID -----------------
    Dim lastRow As Long
    lastRow = wsLV.Cells(wsLV.Rows.Count, ID_COL).End(xlUp).Row
    If lastRow < START_ROW Then Exit Sub

    '--- 1) kopiuj formu³y / formaty / walidacje ----------------
    Dim srcRng As Range, tgtRng As Range
    Set srcRng = wsLV.Range(wsLV.Cells(START_ROW, FIRST_COL), _
                             wsLV.Cells(START_ROW, LAST_COL))
    Set tgtRng = wsLV.Range(wsLV.Cells(START_ROW, FIRST_COL), _
                             wsLV.Cells(lastRow, LAST_COL))

    srcRng.Copy
    tgtRng.PasteSpecial xlPasteFormulas
    tgtRng.PasteSpecial xlPasteFormats
    tgtRng.PasteSpecial xlPasteValidation
    Application.CutCopyMode = False

    '--- 2) ramki segmentowe ------------------------------------
    NakladanieSegmentowychRamek wsLV, START_ROW, lastRow

    '--- 3) wiersz SUM ------------------------------------------
    Dim sumRow As Long: sumRow = lastRow + 2
    Dim colsToSum As Variant
    colsToSum = Array(8, 11, 35, 36, 37, 38, 39, 40, 42, 43, _
                      44, 45, 46, 47, 48)   'ka¿da +1 w stosunku do starej

    Dim i As Long
    For i = LBound(colsToSum) To UBound(colsToSum)
        With wsLV.Cells(sumRow, colsToSum(i))
            .FormulaR1C1 = "=SUM(R" & START_ROW & "C" & colsToSum(i) & _
                            ":R" & lastRow & "C" & colsToSum(i) & ")"
            .Font.Bold = True
        End With
    Next i
    wsLV.Cells(sumRow, 7).Value = "Razem:"    'kol. G
    wsLV.Cells(sumRow, 7).Font.Bold = True
    wsLV.Cells(sumRow, 10).Value = "Razem:"   'kol. J
    wsLV.Cells(sumRow, 10).Font.Bold = True

    NakladanieSegmentowychRamek wsLV, sumRow, sumRow

    '------------------------------------------------------------
    ' 4)  Sekcja  P O D S U M O W A N I E   (AI:AN, 4 wiersze)
    '------------------------------------------------------------
    Dim hdrRow As Long, catRow As Long, unitRow As Long, valRow As Long
    hdrRow = sumRow + 2
    catRow = hdrRow + 1
    unitRow = hdrRow + 2
    valRow = hdrRow + 3

    '4a. Nag³ówek
    With wsLV.Range(wsLV.Cells(hdrRow, 35), wsLV.Cells(hdrRow, 40)) 'AI:AN
        .Merge
        .Value = "PODSUMOWANIE"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = vbWhite
        .Font.Size = 9
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
    End With

    '4b. Etykiety
    Dim cats As Variant
    cats = Array("WARTOŒÆ", "Robocizna", "Materia³", "US£UGA", _
                 "Materia³ w Euro", "Wartoœæ EKE")
    For i = 0 To UBound(cats)
        With wsLV.Cells(catRow, 35 + i)
            .Value = cats(i)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
        End With
    Next i

    '4c. Jednostki
    Dim units As Variant
    units = Array("PLN", "PLN", "PLN", "PLN", "EUR", "PLN")
    For i = 0 To UBound(units)
        With wsLV.Cells(unitRow, 35 + i)
            .Value = units(i)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
        End With
    Next i

    '4d. Formu³y – odwo³ania do SUM
    With wsLV
        .Cells(valRow, 35).Formula = "=" & .Cells(sumRow, 11).Address 'kol. K
        .Cells(valRow, 36).Formula = "=" & .Cells(sumRow, 36).Address 'AJ
        .Cells(valRow, 37).Formula = "=" & .Cells(sumRow, 46).Address 'AT
        .Cells(valRow, 38).Formula = "=" & .Cells(sumRow, 40).Address 'AN
        .Cells(valRow, 39).Formula = "=" & .Cells(sumRow, 47).Address 'AU
        .Cells(valRow, 40).Formula = "=" & .Cells(sumRow, 48).Address 'AV
    End With

    '4e. Ramki – niebieskie
    Dim sumTbl As Range
    Set sumTbl = wsLV.Range(wsLV.Cells(hdrRow, 35), wsLV.Cells(valRow, 40))
    With sumTbl.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 102, 204)
        .Weight = xlThin
        .Item(xlEdgeLeft).Weight = xlMedium
        .Item(xlEdgeTop).Weight = xlMedium
        .Item(xlEdgeRight).Weight = xlMedium
        .Item(xlEdgeBottom).Weight = xlMedium
    End With
End Sub
'==============================================================


'======== NAK£ADANIE RAMEK NA ZDEFINIOWANE SEGMENTY =============
Sub NakladanieSegmentowychRamek(ws As Worksheet, FirstRow As Long, lastRow As Long)

    Dim segments As Variant, s As Variant
    segments = Array(Array(7, 8), _
        Array(10, 11), _
        Array(35, 40), _
        Array(42, 48))
        
    For Each s In segments
        With ws.Range(ws.Cells(FirstRow, s(0)), ws.Cells(lastRow, s(1))).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next s
End Sub
'==============================================================


