Attribute VB_Name = "modSync"
Sub SyncLVtoSource()

    
Const HDR_ROW_LV  As Long = 8   '?  ten wiersz
Const COL_ID      As Long = 1   'ukryta kol. A
Const COL_MARK    As Long = 2   'widoczna kol. B (Nr./Lp.)

Const COL_LV_JEDN  As Long = 5  'E  (Jedn.przedm.)
Const COL_LV_PRZEM As Long = 7  'G  (Przedmiar)

Const OFF_SRC_JEDN  As Long = 4 'A+4 = kol. E w Ÿródle
Const OFF_SRC_PRZEM As Long = 5 'A+5 = kol. F w Ÿródle

    Dim wbLV As Workbook: Set wbLV = ActiveWorkbook

    '------------ wybór pliku Ÿród³owego --------------------------------
    Dim pathSrc As Variant
    MsgBox "Wska¿ oryginalny plik Ÿród³owy.", vbInformation
    pathSrc = Application.GetOpenFilename("Pliki Excel (*.xls*), *.xls*")
    If pathSrc = False Then Exit Sub
    Dim wbSrc As Workbook: Set wbSrc = Workbooks.Open(pathSrc)

    '------------ arkusz Ustawienia w LV --------------------------------
    Dim shSet As Worksheet
    On Error Resume Next
    Set shSet = wbLV.Worksheets("Ustawienia")
    On Error GoTo 0
    If shSet Is Nothing Then
        MsgBox "Brak arkusza 'Ustawienia' w pliku LV.", vbExclamation
        Exit Sub
    End If

    Dim lastSet As Long: lastSet = shSet.Cells(shSet.Rows.Count, 1).End(xlUp).Row
    If lastSet < 2 Then
        MsgBox "'Ustawienia' nie zawiera par arkuszy.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim okCnt As Long, missCnt As Long, dupCnt As Long, rowSet As Long

    For rowSet = 2 To lastSet
        Dim srcName As String: srcName = shSet.Cells(rowSet, 1).Value
        Dim lvName  As String: lvName = shSet.Cells(rowSet, 2).Value

        Dim wsSrc As Worksheet, wsLV As Worksheet
        On Error Resume Next
        Set wsSrc = wbSrc.Worksheets(srcName)
        Set wsLV = wbLV.Worksheets(lvName)
        On Error GoTo 0
        If wsSrc Is Nothing Or wsLV Is Nothing Then GoTo ContinuePair

        '--- s³ownik ID dla tego arkusza LV -----------------------------
        Dim seenIDs As Object: Set seenIDs = CreateObject("Scripting.Dictionary")

        Dim lastLV As Long
        lastLV = wsLV.Cells(wsLV.Rows.Count, COL_ID).End(xlUp).Row

        Dim r As Long
        For r = HDR_ROW_LV To lastLV
            Debug.Assert Not wsLV Is Nothing
            Dim idKey As String: idKey = Trim$(CStr(wsLV.Cells(r, COL_ID).Value))

            '====== 1) pusty ID =========================================
            If idKey = "" Then
                wsLV.Cells(r, COL_MARK).Interior.Color = RGB(255, 204, 204)
                missCnt = missCnt + 1
                GoTo NextRow
            End If

            '====== 2) duplikat ID w tym samym LV =======================
            If seenIDs.Exists(idKey) Then
                wsLV.Cells(r, COL_MARK).Interior.Color = RGB(255, 0, 0)
                dupCnt = dupCnt + 1
                GoTo NextRow
            Else
                seenIDs.Add idKey, True
            End If

            '====== 3) szukaj w arkuszu Ÿród³owym =======================
            Dim hit As Range
            Set hit = wsSrc.Columns(COL_ID).Find(What:=idKey, LookAt:=xlWhole)

            If Not hit Is Nothing Then
                hit.Offset(0, OFF_SRC_JEDN).Value = wsLV.Cells(r, COL_LV_JEDN).Value    'E › E
                hit.Offset(0, OFF_SRC_PRZEM).Value = wsLV.Cells(r, COL_LV_PRZEM).Value  'G › F

                wsLV.Cells(r, COL_MARK).Interior.Pattern = xlNone
                okCnt = okCnt + 1
            Else
                wsLV.Cells(r, COL_MARK).Interior.Color = RGB(255, 204, 204)
                missCnt = missCnt + 1
            End If
NextRow:
        Next r
ContinuePair:
    Next rowSet

    Application.ScreenUpdating = True
    MsgBox "Synchronizacja zakoñczona." & vbCrLf & _
           "Zaktualizowano: " & okCnt & vbCrLf & _
           "Brak dopasowania ID: " & missCnt & vbCrLf & _
           "Duplikaty ID: " & dupCnt, vbInformation
End Sub


