Attribute VB_Name = "modLV_MegaExtender"
'=================== modLVMegaFormat ===================
Option Explicit

Public Sub RozszerzFormulyLVMega(ByVal wsLV As Worksheet)
    If Left$(wsLV.Name, 2) <> "LV" Then Exit Sub

    ' --- sta³e uk³adu ---
    Const PROTO_ROW  As Long = 8      ' wiersz wzorcowy (jak w starym projekcie)
    Const DATA_FIRST As Long = 9      ' pierwszy wiersz danych
    Const FIRST_COL  As Long = 7      ' G
    Const LAST_COL   As Long = 47     ' AU
    Const ID_COL     As Long = 2      ' np. B – kolumna pewna dla koñca danych

    ' --- zmienne ---
    Dim lastDataRow As Long
    Dim srcProtoFmt As Range, dstAll As Range
    Dim rowProto As Range, srcForm As Range
    Dim c As Range, rngCol As Range, blanks As Range
    Dim sumRow As Long
    Dim sumMap As Variant
    Dim k As Long, tgtCol As Long, srcCol As Long
    Dim hdrRow As Long, catRow As Long, unitRow As Long, valRow As Long
    Dim cats As Variant, units As Variant
    Dim sumTbl As Range
    Dim i As Long

    ' --- ostatni wiersz danych ---
    lastDataRow = wsLV.Cells(wsLV.Rows.Count, ID_COL).End(xlUp).Row
    If lastDataRow < DATA_FIRST Then Exit Sub

    ' === (2) FORMATY + WALIDACJE z wiersza 9  ===
    Set srcProtoFmt = wsLV.Range(wsLV.Cells(DATA_FIRST, FIRST_COL), wsLV.Cells(DATA_FIRST, LAST_COL))
    Set dstAll = wsLV.Range(wsLV.Cells(DATA_FIRST, FIRST_COL), wsLV.Cells(lastDataRow, LAST_COL))

    srcProtoFmt.Copy
    dstAll.PasteSpecial xlPasteFormats
    dstAll.PasteSpecial xlPasteValidation
    Application.CutCopyMode = False

    ' === (3) FORMU£Y – bierzemy wzorzec Z WIERSZA 8, uzupe³niamy tylko PUSTE komórki ===
    Set rowProto = wsLV.Range(wsLV.Cells(PROTO_ROW, FIRST_COL), wsLV.Cells(PROTO_ROW, LAST_COL))

    On Error Resume Next
    Set srcForm = rowProto.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not srcForm Is Nothing Then
        For Each c In srcForm.Cells
            Set rngCol = wsLV.Range(wsLV.Cells(DATA_FIRST, c.Column), wsLV.Cells(lastDataRow, c.Column))
            On Error Resume Next
            Set blanks = rngCol.SpecialCells(xlCellTypeBlanks)
            On Error GoTo 0
            If Not blanks Is Nothing Then
                blanks.FormulaR1C1 = c.FormulaR1C1
            End If
            Set blanks = Nothing
        Next c
    End If

    ' === (4) cienkie ramki na segmentach danych ===
    NakladanieSegmentowychRamekMega wsLV, DATA_FIRST, lastDataRow

    ' === (5) WIERSZ SUM ===
    sumRow = lastDataRow + 2

    ' docelowa_kolumna › Ÿród³owa_kolumna
    ' G pokazuje sumê z H; J pokazuje sumê z J; AH..AM mapowanie specjalne; AO..AU sumuj¹ siebie
    sumMap = Array(Array(7, 7), Array(10, 10), Array(34, 34), Array(35, 35), Array(36, 45), Array(37, 39), Array(38, 46), Array(39, 47), Array(41, 41), Array(42, 42), Array(43, 43), Array(44, 44), Array(45, 45), Array(46, 46), Array(47, 47))

    For k = LBound(sumMap) To UBound(sumMap)
        tgtCol = sumMap(k)(0)
        srcCol = sumMap(k)(1)
        With wsLV.Cells(sumRow, tgtCol)
            .FormulaR1C1 = "=SUM(R" & DATA_FIRST & "C" & srcCol & ":R" & lastDataRow & "C" & srcCol & ")"
            .Font.Bold = True
            If tgtCol = 38 Then
                .NumberFormat = "#,##0.00 [$€-x-euro1]"
            Else
                .NumberFormat = "#,##0.00 $"
            End If
        End With
    Next k


    ' etykiety "Razem:" (dla G i J)
    wsLV.Cells(sumRow, 6).Value = "Razem:"   ' F – opis dla sumy w G (7)
    wsLV.Cells(sumRow, 6).Font.Bold = True
    wsLV.Cells(sumRow, 9).Value = "Razem:"   ' I – opis dla sumy w J (10)
    wsLV.Cells(sumRow, 9).Font.Bold = True

    ' ramki dla wiersza SUM
    NakladanieSegmentowychRamekMega wsLV, sumRow, sumRow

    ' === (6) sekcja PODSUMOWANIE w AH:AM (34–39), 4 wiersze ===
    hdrRow = sumRow + 2
    catRow = hdrRow + 1
    unitRow = hdrRow + 2
    valRow = hdrRow + 3

    ' 6a. nag³ówek
    With wsLV.Range(wsLV.Cells(hdrRow, 34), wsLV.Cells(hdrRow, 39))
        .Merge
        .Value = "PODSUMOWANIE"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = vbWhite
        .Font.Size = 9
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
    End With

    ' 6b. etykiety
    cats = Array("WARTOŒÆ", "Robocizna", "Materia³", "US£UGA", "Materia³ w Euro", "Wartoœæ EKE")
    For i = 0 To UBound(cats)
        With wsLV.Cells(catRow, 34 + i)
            .Value = cats(i)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
        End With
    Next i

    ' 6c. jednostki
    units = Array("PLN", "PLN", "PLN", "PLN", "EUR", "PLN")
    For i = 0 To UBound(units)
        With wsLV.Cells(unitRow, 34 + i)
            .Value = units(i)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
        End With
    Next i

    ' 6d. wartoœci – odwo³ania do wiersza SUM
    With wsLV
        .Cells(valRow, 34).Formula = "=" & .Cells(sumRow, 10).Address
        .Cells(valRow, 35).Formula = "=" & .Cells(sumRow, 35).Address
        .Cells(valRow, 36).Formula = "=" & .Cells(sumRow, 45).Address
        .Cells(valRow, 37).Formula = "=" & .Cells(sumRow, 39).Address
        .Cells(valRow, 38).Formula = "=" & .Cells(sumRow, 46).Address
        .Cells(valRow, 39).Formula = "=" & .Cells(sumRow, 47).Address
    End With
    ' 6d+. format waluty dla wiersza PODSUMOWANIE
    With wsLV
        .Cells(sumRow, 46).NumberFormat = "#,##0.00 [$€-x-euro1]"
        .Cells(valRow, 38).NumberFormat = "#,##0.00 [$€-x-euro1]" ' AL = EUR
        .Cells(valRow, 34).NumberFormat = "#,##0.00 $"            ' AH = PLN
        .Cells(valRow, 35).NumberFormat = "#,##0.00 $"
        .Cells(valRow, 36).NumberFormat = "#,##0.00 $"
        .Cells(valRow, 37).NumberFormat = "#,##0.00 $"
        .Cells(valRow, 39).NumberFormat = "#,##0.00 $"
    End With


    ' 6e. ramki wokó³ sekcji PODSUMOWANIE
    Set sumTbl = wsLV.Range(wsLV.Cells(hdrRow, 34), wsLV.Cells(valRow, 39))
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

Private Sub NakladanieSegmentowychRamekMega(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    ' Przesuniêto G:H -> F:G oraz I:J -> H:I; reszta bez zmian (AH:AM i AO:AU)
    Dim segments As Variant, s As Variant
    segments = Array(Array(6, 7), Array(9, 10), Array(34, 39), Array(41, 47))
    For Each s In segments
        With ws.Range(ws.Cells(firstRow, s(0)), ws.Cells(lastRow, s(1))).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next s
End Sub
'================= /modLVMegaFormat ===================

