Attribute VB_Name = "modLV_Mega"
'==========================  modMegaLV  ==========================
Option Explicit

'--- RÓD£O (arkusze b-2..b-40) ---
Private Const SRC_COL_PROD   As Long = 2   'B  producent
Private Const SRC_COL_DOST   As Long = 3   'C  dostawca
Private Const SRC_COL_OPIS   As Long = 8   'H  produkt/opis
Private Const SRC_COL_ILOSC  As Long = 9   'I  iloœæ
Private Const SRC_COL_JM     As Long = 10  'J  jedn.
Private Const SRC_COL_CENA   As Long = 11  'K  cena zakupu
Private Const SRC_COL_WAL    As Long = 12  'L  waluta
Private Const SRC_COL_RABAT  As Long = 13  'M  rabat %
Private Const SRC_COL_CENAR  As Long = 14  'N  cena jedn. z rabatem
Private Const SRC_COL_RG_H   As Long = 17  'P  iloœæ r-g (w godzinach)
Private Const SRC_FIRST_DATA As Long = 18  '17 = nag³ówki, dane od 18

'--- CEL = LV **bez ID** ---
Private Const TGT_COL_OPIS      As Long = 2   'B
Private Const TGT_COL_ILOSC     As Long = 4   'D
Private Const TGT_COL_JM        As Long = 5   'E
Private Const TGT_COL_CENA_PLN  As Long = 11  'K
Private Const TGT_COL_CENA_EUR  As Long = 12  'L
Private Const TGT_COL_RABAT     As Long = 13  'M
Private Const TGT_COL_RG_MIN    As Long = 18  'R
Private Const TGT_COL_DOSTPROD  As Long = 30  'AD
Private Const TGT_FIRST_ROW     As Long = 8

Private Const MARK_CONVERTED As Boolean = True
Private Const CONV_FONT_COLOR As Long = vbRed   'np. RGB(192, 0, 0)


Private Const LV_TEMPLATE_NAME As String = "LV_SZABLON"

'============================= PUBLIC =============================
'Buduje mega-LV do pierwszego arkusza „LV” w pliku docelowym.
Public Sub BuildMegaLV_ToTarget()
    Dim wbSrc As Workbook, wbTgt As Workbook
    Dim wsRazem As Worksheet, wsOut As Worksheet, wsTpl As Worksheet
    Dim pathTgt As Variant

    Set wbSrc = ActiveWorkbook
    If wbSrc Is Nothing Then Exit Sub

    On Error Resume Next
    Set wsRazem = wbSrc.Worksheets("Razem")
    On Error GoTo 0
    If wsRazem Is Nothing Then
        MsgBox "Brak arkusza 'Razem' w pliku Ÿród³owym.", vbExclamation
        Exit Sub
    End If

    '–– 1) wybór pliku LV docelowego ––
    pathTgt = Application.GetOpenFilename("Pliki Excel (*.xls*;*.xlsm;*.xltm),*.xls*;*.xlsm;*.xltm")
    If pathTgt = False Then Exit Sub
    Set wbTgt = Workbooks.Open(pathTgt)

    '–– 2) pierwszy arkusz „LV*” ? LV_SZABLON (fallback: LV_SZABLON) ––
    Set wsOut = GetFirstLVSheet(wbTgt)
    If wsOut Is Nothing Then
        MsgBox "W pliku docelowym nie ma ¿adnego arkusza 'LV'.", vbCritical
        Exit Sub
    End If

    '–– 2a) nag³ówki 1:8 z szablonu (jeœli dostêpny i to nie ten sam arkusz) ––
    Set wsTpl = GetTemplateLV(wbTgt)
    If Not wsTpl Is Nothing Then
        If Not wsOut Is wsTpl Then
            wsTpl.Rows("1:8").Copy
            With wsOut.Rows("1:8")
                .PasteSpecial xlPasteAllUsingSourceTheme
                .PasteSpecial xlPasteColumnWidths
            End With
            Application.CutCopyMode = False
        End If
    End If

    '–– 3) wyczyœæ dane (od 9 w dó³) ––
    Dim rngClear As Range
   
    Set rngClear = Union( _
        wsOut.Range(wsOut.Cells(9, 2), wsOut.Cells(wsOut.Rows.Count, 5)), _
        wsOut.Range(wsOut.Cells(9, 8), wsOut.Cells(wsOut.Rows.Count, 15)), _
        wsOut.Range(wsOut.Cells(9, 17), wsOut.Cells(wsOut.Rows.Count, 47)) _
    )
    rngClear.ClearContents





    Dim wOut As Long: wOut = TGT_FIRST_ROW
    Dim copied As Long

    'STARTUJEMY od 10, ¿eby pomin¹æ b-1 (zgodnie z Twoj¹ uwag¹)
    Dim r As Long: r = 10
    Do While Len(Trim$(wsRazem.Cells(r, 1).Value2)) > 0
        Dim tabName As String, include As Long
        tabName = CStr(wsRazem.Cells(r, 1).Value2)          'np. b-2 lub 2
        include = CLng(Val(wsRazem.Cells(r, 2).Value2))      '0/1 (kol. B)

        If include = 1 Then
            Dim wsSys As Worksheet
            Set wsSys = ResolveSystemSheet(wbSrc, tabName)
            If Not wsSys Is Nothing Then
                wOut = AppendSectionHeader(wsOut, wsSys, wOut)
                wOut = AppendSystemRows(wsOut, wsSys, wOut)
            Else
                'Debug.Print "Nie znaleziono arkusza dla pozycji:", tabName
            End If
        End If
        r = r + 1
    Loop
    
    copied = Application.Max(0, wOut - TGT_FIRST_ROW)
    
    If copied > 0 Then
        ' uruchom rozszerzanie na arkuszu wyjœciowym,
        ' bazuj¹c na tym samym wierszu startowym:
        Call RozszerzFormulyLVMega(wsOut)
    End If
    
    
    wbTgt.Activate: wsOut.Activate
    MsgBox "Wstawiono " & copied & " wierszy do arkusza '" & wsOut.Name & "'.", vbInformation

End Sub

'============================= HELPERS ============================

'Resolwer arkusza systemu: próbuje nazwy jak w „Razem” oraz warianty „b-<nr>”, „<nr>”.
Private Function ResolveSystemSheet(wb As Workbook, ByVal tabName As String) As Worksheet
    Dim nm As String: nm = Trim$(tabName)
    Dim re As Object, m As Object

    '1) próba „jak jest”
    On Error Resume Next
    Set ResolveSystemSheet = wb.Worksheets(nm)
    On Error GoTo 0
    If Not ResolveSystemSheet Is Nothing Then Exit Function

    '2) jeœli jest w formacie b-<num> lub <num> – spróbuj wariantów
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d+)$"
    re.IgnoreCase = True
    If re.test(nm) Then
        Set m = re.Execute(nm)(0)
        Dim numTxt As String: numTxt = m.SubMatches(0)

        '2a) „b-<num>”
        On Error Resume Next
        Set ResolveSystemSheet = wb.Worksheets("b-" & numTxt)
        On Error GoTo 0
        If Not ResolveSystemSheet Is Nothing Then Exit Function

        '2b) sam „<num>”
        On Error Resume Next
        Set ResolveSystemSheet = wb.Worksheets(CStr(numTxt))
        On Error GoTo 0
        If Not ResolveSystemSheet Is Nothing Then Exit Function
    End If
End Function

'Pierwszy arkusz zaczynaj¹cy siê od „LV”, z pominiêciem „LV_SZABLON”.
Private Function GetFirstLVSheet(wb As Workbook) As Worksheet
    Dim sh As Worksheet, tpl As Worksheet
    On Error Resume Next
    Set tpl = wb.Worksheets(LV_TEMPLATE_NAME)
    On Error GoTo 0

    For Each sh In wb.Worksheets
        If LCase$(Left$(sh.Name, 2)) = "lv" Then
            If tpl Is Nothing Or Not sh Is tpl Then
                Set GetFirstLVSheet = sh
                Exit Function
            End If
        End If
    Next sh
    If Not tpl Is Nothing Then Set GetFirstLVSheet = tpl
End Function

Private Function GetTemplateLV(wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetTemplateLV = wb.Worksheets(LV_TEMPLATE_NAME)
    On Error GoTo 0
    If GetTemplateLV Is Nothing Then
        Dim sh As Worksheet
        For Each sh In wb.Worksheets
            If LCase$(Left$(sh.Name, 2)) = "lv" Then
                Set GetTemplateLV = sh: Exit Function
            End If
        Next sh
    End If
End Function

Private Function AppendSectionHeader(wsOut As Worksheet, wsSrc As Worksheet, ByVal wOut As Long) As Long
    Dim t As String: t = Trim$(CStr(wsSrc.Range("H16").Value))
    If t <> "" Then
        With wsOut.Rows(wOut)
            .Interior.Color = RGB(77, 148, 255)
            .Font.Bold = True
        End With
        wsOut.Cells(wOut, TGT_COL_OPIS).Value = t
        wOut = wOut + 1
    End If
    AppendSectionHeader = wOut
End Function

Private Function AppendSystemRows(wsOut As Worksheet, wsSrc As Worksheet, ByVal wOut As Long) As Long
    Dim r As Long: r = SRC_FIRST_DATA
    Do While Len(Trim$(wsSrc.Cells(r, SRC_COL_OPIS).Value2)) > 0
        Dim opis As String, jm As String
        Dim ilosc As Double, cena As Double, cenaR As Double
        Dim wal As String, rab As Double, rgH As Double
        Dim dst As String, prod As String, dp As String

        opis = CStr(wsSrc.Cells(r, SRC_COL_OPIS).Value2)
        ilosc = ValD(wsSrc.Cells(r, SRC_COL_ILOSC).Value2)
        jm = CStr(wsSrc.Cells(r, SRC_COL_JM).Value2)
        cena = ValD(wsSrc.Cells(r, SRC_COL_CENA).Value2)
        cenaR = ValD(wsSrc.Cells(r, SRC_COL_CENAR).Value2)
        wal = UCase$(Trim$(CStr(wsSrc.Cells(r, SRC_COL_WAL).Value2)))
        rab = ValPercent(wsSrc.Cells(r, SRC_COL_RABAT).Value2)
        rgH = ValD(wsSrc.Cells(r, SRC_COL_RG_H).Value2)
        prod = CStr(wsSrc.Cells(r, SRC_COL_PROD).Value2)
        dst = CStr(wsSrc.Cells(r, SRC_COL_DOST).Value2)
        dp = Trim$(dst & IIf(dst <> "" And prod <> "", " / ", "") & prod)

        wsOut.Cells(wOut, TGT_COL_OPIS).Value = opis
        wsOut.Cells(wOut, TGT_COL_ILOSC).Value = ilosc
        wsOut.Cells(wOut, TGT_COL_JM).Value = jm
        wsOut.Cells(wOut, TGT_COL_DOSTPROD).Value = dp

        ' --- ustalenie ceny i Ÿród³a ---
        Dim basePrice As Double
        Dim fromCenaR As Boolean

        If cena > 0 Then
            basePrice = cena
            fromCenaR = False
            If rab <> 0 Then
                With wsOut.Cells(wOut, TGT_COL_RABAT)
                    .Value = rab            ' np. -0.04
                    .NumberFormat = "0.00%" ' poka¿e -4,00%
                End With
            Else
                wsOut.Cells(wOut, TGT_COL_RABAT).ClearContents
            End If
        ElseIf cenaR > 0 Then
            basePrice = cenaR
            fromCenaR = True
            wsOut.Cells(wOut, TGT_COL_RABAT).ClearContents
        Else
            basePrice = 0
            fromCenaR = False
            wsOut.Cells(wOut, TGT_COL_RABAT).ClearContents
        End If
        

        ' --- ZAWSZE: reset stylu/komentarzy w cenach ---
        With wsOut.Cells(wOut, TGT_COL_CENA_PLN)
            .Font.ColorIndex = xlColorIndexAutomatic
            On Error Resume Next: .ClearComments: On Error GoTo 0
        End With
        With wsOut.Cells(wOut, TGT_COL_CENA_EUR)
            .Font.ColorIndex = xlColorIndexAutomatic
            On Error Resume Next: .ClearComments: On Error GoTo 0
        End With

        ' --- wycena / przewalutowanie ---
        If basePrice > 0 Then
            If fromCenaR Then
                ' N zawsze traktujemy jako PLN – bez przeliczenia i bez oznaczania
                wsOut.Cells(wOut, TGT_COL_CENA_PLN).Value = basePrice
                wsOut.Cells(wOut, TGT_COL_CENA_EUR).ClearContents

            Else
                Select Case wal
                    Case "PLN", ""                       ' puste = traktuj jak PLN
                        wsOut.Cells(wOut, TGT_COL_CENA_PLN).Value = basePrice
                        wsOut.Cells(wOut, TGT_COL_CENA_EUR).ClearContents

                    Case "EUR"
                        ' EUR wpisujemy jako EUR, bez przeliczeñ
                        wsOut.Cells(wOut, TGT_COL_CENA_EUR).Value = basePrice
                        wsOut.Cells(wOut, TGT_COL_CENA_PLN).ClearContents

                    Case "GBP", "USD", "CAD"
                        ' realne przewalutowanie -> oznaczamy
                        Dim rate As Double: rate = GetRate(wsSrc, wal)
                        wsOut.Cells(wOut, TGT_COL_CENA_PLN).Value = basePrice * rate
                        wsOut.Cells(wOut, TGT_COL_CENA_EUR).ClearContents

                        If MARK_CONVERTED Then
                            With wsOut.Cells(wOut, TGT_COL_CENA_PLN)
                                .Font.Color = CONV_FONT_COLOR
                                On Error Resume Next
                                .AddComment "Przeliczono z " & wal & " po kursie " & Format(rate, "0.######")
                                On Error GoTo 0
                            End With
                        End If

                    Case Else
                        ' inne kody walut – traktuj jak przeliczenie na PLN
                        Dim rate2 As Double: rate2 = GetRate(wsSrc, wal)
                        If rate2 > 0 Then
                            wsOut.Cells(wOut, TGT_COL_CENA_PLN).Value = basePrice * rate2
                            wsOut.Cells(wOut, TGT_COL_CENA_EUR).ClearContents
                            If MARK_CONVERTED Then
                                With wsOut.Cells(wOut, TGT_COL_CENA_PLN)
                                    .Font.Color = CONV_FONT_COLOR
                                    On Error Resume Next
                                    .AddComment "Przeliczono z " & wal & " po kursie " & Format(rate2, "0.######")
                                    On Error GoTo 0
                                End With
                            End If
                        Else
                            ' brak kursu – wpisz do PLN bez oznaczenia, ¿eby nie zgubiæ wartoœci
                            wsOut.Cells(wOut, TGT_COL_CENA_PLN).Value = basePrice
                            wsOut.Cells(wOut, TGT_COL_CENA_EUR).ClearContents
                        End If
                End Select
            End If
        Else
            ' brak ceny – wyczyœæ pola cenowe
            wsOut.Cells(wOut, TGT_COL_CENA_PLN).ClearContents
            wsOut.Cells(wOut, TGT_COL_CENA_EUR).ClearContents
        End If

        ' --- RG w minutach * stawka z K$3 ---
        If rgH > 0 Then
            Dim rgLit As String
            rgLit = Replace(CStr(rgH), Application.International(xlDecimalSeparator), ".") ' przecinek -> kropka
            wsOut.Cells(wOut, TGT_COL_RG_MIN).Formula = "=(" & rgLit & "/60)*" & wsOut.Range("K3").Address(RowAbsolute:=True, ColumnAbsolute:=True)
        End If



        wOut = wOut + 1
        r = r + 1
    Loop

    AppendSystemRows = wOut
End Function


Private Function GetRate(wsSrc As Worksheet, ByVal code As String) As Double
    Dim r As Long
    For r = 2 To 6
        If UCase$(Trim$(CStr(wsSrc.Cells(r, 1).Value2))) = code Then
            GetRate = ValD(wsSrc.Cells(r, 2).Value2)
            Exit Function
        End If
    Next r
    GetRate = 0#
End Function

Private Function ValD(v As Variant) As Double
    Dim s As String: s = CStr(v)
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")
    s = Replace(s, ",", ".")
    ValD = Val(s)
End Function

Private Function ValPercent(v As Variant) As Double
    Dim s As String: s = CStr(v)
    ' wyczyœæ spacje (w tym NBSP) i nawiasy ksiêgowe
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")
    Dim isNeg As Boolean
    If InStr(s, "(") > 0 And InStr(s, ")") > 0 Then isNeg = True
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")

    ' usuñ znak % i ustaw kropkê jako separator
    s = Replace(s, "%", "")
    s = Replace(s, ",", ".")

    If s = "" Or s = "." Or s = "-" Then
        ValPercent = 0#
        Exit Function
    End If

    Dim x As Double: x = Val(s)
    If isNeg Then x = -Abs(x)

    ' DZIEL PRZEZ 100 tak¿e dla wartoœci ujemnych
    If Abs(x) > 1 Then x = x / 100#

    ValPercent = x
End Function


'----------------------------------------------
'  Rozszerzanie formu³/formatów w arkuszu LV
'  – PROTO: wiersz 8, dane od wiersza 9
'  – kopiuje formu³y tylko z kolumn, gdzie prototyp MA formu³ê
'  – PODSUMOWANIE: AH:AM (34..39)
'----------------------------------------------
Private Function LV_LastDataRow(ByVal ws As Worksheet, _
                                ByVal START_ROW As Long, _
                                Optional ByVal FallbackCol As Long = 2) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr >= START_ROW Then LV_LastDataRow = lr: Exit Function

    lr = ws.Cells(ws.Rows.Count, FallbackCol).End(xlUp).Row   'kol. B (Opis)
    If lr >= START_ROW Then LV_LastDataRow = lr: Exit Function

    On Error Resume Next
    lr = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                       SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    If lr < START_ROW Then lr = START_ROW
    LV_LastDataRow = lr
End Function
