Attribute VB_Name = "modSync"
'=====================  modSync  =================================
Option Explicit


Sub SyncLVtoSource()

    Const COL_ID As Long = 1          'ukryta kolumna A z ID
    Const HDR_ROW_LV As Long = 8      'pierwszy wiersz tabeli w LV

    Dim wbLV As Workbook:  Set wbLV = ActiveWorkbook
    If wbLV Is Nothing Then Exit Sub

    '–––– 1. wska¿ oryginalny plik Ÿród³owy –––––––––––––––––––––
    MsgBox "Wska¿ oryginalny plik Ÿród³owy.", vbInformation
    Dim pathSrc As Variant
    pathSrc = Application.GetOpenFilename("Pliki Excel (*.xls*;*.xlsm),*.xls*;*.xlsm")
    If pathSrc = False Then Exit Sub

    Dim wbSrc As Workbook:  Set wbSrc = Workbooks.Open(pathSrc)

    '–––– 2. arkusz ”Ustawienia” w pliku LV ––––––––––––––––––––
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

    '–––– 3. formularz – mapowanie kolumn –––––––––––––––––––––––
    Load frmSyncCols
    frmSyncCols.Show
    If Not frmSyncCols.FormOK Then Exit Sub

    Dim colLV_Cena As Long, colLV_Wart As Long
    Dim colSRC_Cena As Long, colSRC_Wart As Long

    colLV_Cena = frmSyncCols.LV_Cena
    colLV_Wart = frmSyncCols.LV_Wart
    colSRC_Cena = frmSyncCols.SRC_Cena
    colSRC_Wart = frmSyncCols.SRC_Wart     '––> wszystkie >0 bo sprawdzane w formularzu

    '–––– 4. synchronizacja dla ka¿dej pary –––––––––––––––––––––
    Application.ScreenUpdating = False

    Dim okCnt As Long, missCnt As Long, dupCnt As Long
    Dim rowSet As Long

    For rowSet = 2 To lastSet
        Dim srcName As String: srcName = shSet.Cells(rowSet, 1).Value
        Dim lvName  As String: lvName = shSet.Cells(rowSet, 2).Value

        Dim wsSrc As Worksheet, wsLV As Worksheet
        On Error Resume Next
        Set wsSrc = wbSrc.Worksheets(srcName)
        Set wsLV = wbLV.Worksheets(lvName)
        On Error GoTo 0
        If wsSrc Is Nothing Or wsLV Is Nothing Then GoTo ContinuePair

        '–– s³ownik ID ? wiersz w Source –––––––––––––––––––––––
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        Dim rSrc As Long, lastSrc As Long
        lastSrc = wsSrc.Cells(wsSrc.Rows.Count, COL_ID).End(xlUp).Row
        For rSrc = 1 To lastSrc
            If IsNumeric(wsSrc.Cells(rSrc, COL_ID).Value) Then _
                dict(CLng(wsSrc.Cells(rSrc, COL_ID).Value)) = rSrc
        Next rSrc

        '–– przegl¹daj LV od wiersza 8 –––––––––––––––––––––––––
        Dim lastLV As Long: lastLV = wsLV.Cells(wsLV.Rows.Count, COL_ID).End(xlUp).Row
        Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

        Dim r As Long
        For r = HDR_ROW_LV To lastLV
            Dim idKey As String: idKey = Trim$(CStr(wsLV.Cells(r, COL_ID).Value))

            '1) pusty ID  › ró¿owe t³o
            If idKey = "" Then
                wsLV.Cells(r, 2).Interior.Color = RGB(255, 204, 204)
                missCnt = missCnt + 1
                GoTo NextRow
            End If

            '2) duplikat ID  › czerwone t³o
            If seen.Exists(idKey) Then
                wsLV.Cells(r, 2).Interior.Color = RGB(255, 0, 0)
                dupCnt = dupCnt + 1
                GoTo NextRow
            Else
                seen.Add idKey, True
            End If

            '3) szukaj w Source
            If dict.Exists(CLng(idKey)) Then
                Dim trgRow As Long: trgRow = dict(CLng(idKey))
                
                wsSrc.Cells(trgRow, colSRC_Cena).Value = wsLV.Cells(r, colLV_Cena).Value
                wsSrc.Cells(trgRow, colSRC_Wart).Value = wsLV.Cells(r, colLV_Wart).Value
                
                wsLV.Cells(r, 2).Interior.Pattern = xlNone
                okCnt = okCnt + 1
            Else
                wsLV.Cells(r, 2).Interior.Color = RGB(255, 204, 204)
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
'================================================================


