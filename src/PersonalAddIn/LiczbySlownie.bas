Attribute VB_Name = "LiczbySlownie"
'==================== modLiczbyNaSlowa_PL ====================
Option Explicit

' Publiczna funkcja: liczba -> s³owa (PL).
' - value: liczba lub tekst z liczb¹
' - asMoney:=True -> „z³ote i XX/100” (dla kwot)
Public Function NumberToWordsPL(ByVal value As Variant, Optional ByVal asMoney As Boolean = False) As String
    Dim s As String: s = Trim$(CStr(value))
    If s = "" Then NumberToWordsPL = "zero": Exit Function

    Dim neg As Boolean
    If Left$(s, 1) = "-" Then neg = True: s = Mid$(s, 2)

    Dim zl As Currency, gr As Long
    Dim dotPos As Long: dotPos = InStr(s, Application.International(xlDecimalSeparator))
    If dotPos > 0 Then
        zl = CDec(Left$(s, dotPos - 1))
        gr = CLng(Format(CDec("0" & Mid$(s, dotPos)), "0.00") * 100)
    Else
        zl = CDec(s)
        gr = 0
    End If

    Dim body As String
    body = ChunkToWordsPL(zl)

    If asMoney Then
        body = body & " " & Pluralize zl, "z³oty", "z³ote", "z³otych"
        If gr > 0 Then
            body = body & " i " & Format(gr, "00") & "/100"
        Else
            body = body & " i 00/100"
        End If
    Else
        If gr > 0 Then body = body & " przecinek " & ChunkToWordsPL(gr)
    End If

    If neg Then body = "minus " & body
    NumberToWordsPL = Trim$(body)
End Function

' Pomocnicza: zamienia 0..999 999 999 999 na s³owa
Private Function ChunkToWordsPL(ByVal n As Double) As String
    If n = 0 Then ChunkToWordsPL = "zero": Exit Function
    If n < 0 Then n = -n

    Dim units(): units = Split("zero jeden dwa trzy cztery piêæ szeœæ siedem osiem dziewiêæ")
    Dim teens(): teens = Split("dziesiêæ jedenaœcie dwanaœcie trzynaœcie czternaœcie piêtnaœcie szesnaœcie siedemnaœcie osiemnaœcie dziewiêtnaœcie")
    Dim tens():  tens = Split("zero dziesiêæ dwadzieœcia trzydzieœci czterdzieœci piêædziesi¹t szeœædziesi¹t siedemdziesi¹t osiemdziesi¹t dziewiêædziesi¹t")
    Dim hund():  hund = Split("zero sto dwieœcie trzysta czterysta piêæset szeœæset siedemset osiemset dziewiêæset")

    ' formy mianownik/mianownik mnoga/dope³niacz mnoga
    Dim scales(0 To 4, 0 To 2) As String
    scales(0, 0) = "":        scales(0, 1) = "":         scales(0, 2) = ""
    scales(1, 0) = "tysi¹c":  scales(1, 1) = "tysi¹ce":  scales(1, 2) = "tysiêcy"
    scales(2, 0) = "milion":  scales(2, 1) = "miliony":  scales(2, 2) = "milionów"
    scales(3, 0) = "miliard": scales(3, 1) = "miliardy": scales(3, 2) = "miliardów"
    scales(4, 0) = "bilion":  scales(4, 1) = "biliony":  scales(4, 2) = "bilionów"

    Dim parts() As String: ReDim parts(0 To 4)
    Dim i As Long, chunk As Long, words As String, res As String
    For i = 0 To 4
        chunk = n Mod 1000
        If chunk > 0 Then
            words = ThreeDigitsPL(chunk, units, teens, tens, hund)
            res = words & IIf(i > 0, " " & ScaleForm(i, chunk, scales), "") & IIf(res <> "", " " & res, "")
        End If
        n = Int(n / 1000)
        If n = 0 Then Exit For
    Next i

    ChunkToWordsPL = Trim$(res)
End Function

' 0..999 do s³ów (zale¿ne listy przekazane jako tablice)
Private Function ThreeDigitsPL(ByVal x As Long, units, teens, tens, hund) As String
    Dim h As Long, t As Long, u As Long
    h = x \ 100
    t = (x Mod 100) \ 10
    u = x Mod 10

    Dim s As String
    If h > 0 Then s = hund(h)

    If t = 1 Then
        If s <> "" Then s = s & " "
        s = s & teens(u)
    Else
        If t > 0 Then
            If s <> "" Then s = s & " "
            s = s & tens(t)
        End If
        If u > 0 Then
            If s <> "" Then s = s & " "
            s = s & units(u)
        End If
    End If

    ThreeDigitsPL = s
End Function

' Wybór formy: 1 -> mianownik; 2..4 -> mianownik mnoga; reszta -> dope³niacz mnoga, z wyj¹tkiem 12–14.
Private Function ScaleForm(ByVal idx As Long, ByVal chunk As Long, ByRef scales) As String
    Dim d As Long: d = chunk Mod 10
    Dim dd As Long: dd = chunk Mod 100
    Dim formIdx As Long

    If d = 1 And dd <> 11 Then
        formIdx = 0
    ElseIf d >= 2 And d <= 4 And (dd < 12 Or dd > 14) Then
        formIdx = 1
    Else
        formIdx = 2
    End If
    ScaleForm = scales(idx, formIdx)
End Function

' Odmiana s³owa (np. "z³oty") wg liczby n
Private Function Pluralize(ByVal n As Double, ByVal sing As String, ByVal plural As String, ByVal genPlural As String) As String
    Dim d As Long: d = CLng(n) Mod 10
    Dim dd As Long: dd = CLng(n) Mod 100
    If CLng(n) = 1 Then
        Pluralize = sing
    ElseIf d >= 2 And d <= 4 And (dd < 12 Or dd > 14) Then
        Pluralize = plural
    Else
        Pluralize = genPlural
    End If
End Function
'===============================================================

'================== modExportDoWord ==================
Option Explicit

' Eksport: zaznaczone komórki -> Word (tabela 2 kolumny: Liczba | S³ownie)
' - asMoney:=True -> zapis "z³ote i XX/100"
Public Sub ExportSelectionNumbersToWord(Optional ByVal asMoney As Boolean = False)
    Dim rng As Range, cell As Range
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "Zaznacz komórki z liczbami w Excelu.", vbExclamation
        Exit Sub
    End If

    Dim arrOut() As Variant
    Dim r As Long, cnt As Long: cnt = 0
    For Each cell In rng.Cells
        If Len(Trim$(cell.value)) > 0 And IsNumeric(cell.value) Then cnt = cnt + 1
    Next cell
    If cnt = 0 Then
        MsgBox "Brak liczb w zaznaczeniu.", vbInformation
        Exit Sub
    End If

    ReDim arrOut(1 To cnt, 1 To 2)
    r = 0
    For Each cell In rng.Cells
        If Len(Trim$(cell.value)) > 0 And IsNumeric(cell.value) Then
            r = r + 1
            arrOut(r, 1) = cell.value
            arrOut(r, 2) = NumberToWordsPL(cell.value, asMoney)
        End If
    Next cell

    ' === Word (late binding) ===
    Dim wdApp As Object, wdDoc As Object, wdTable As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add

    wdDoc.Content.Paragraphs.Alignment = 1 'wdAlignParagraphLeft
    wdDoc.Content.Text = "Liczby na s³owa" & vbCrLf
    wdDoc.Paragraphs.Add

    Set wdTable = wdDoc.Tables.Add(Range:=wdDoc.Paragraphs.Last.Range, NumRows:=cnt + 1, NumColumns:=2)
    With wdTable
        .cell(1, 1).Range.Text = "Liczba"
        .cell(1, 2).Range.Text = "Zapis s³owny"
        .Rows(1).Range.Bold = True
        .PreferredWidthType = 1 'wdPreferredWidthPercent
        .PreferredWidth = 100
        .Columns(1).PreferredWidthType = 1
        .Columns(1).PreferredWidth = 25
        .Columns(2).PreferredWidthType = 1
        .Columns(2).PreferredWidth = 75
        Dim i As Long
        For i = 1 To cnt
            .cell(i + 1, 1).Range.Text = CStr(arrOut(i, 1))
            .cell(i + 1, 2).Range.Text = CStr(arrOut(i, 2))
        Next i
        .Rows.Borders.Enable = True
    End With

    Set wdTable = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
'=====================================================

