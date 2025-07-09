Attribute VB_Name = "sklejanieArkuszy"

Option Explicit


Public Sub MergeExcelFilesInFolder()

    Dim folderPath As String
    folderPath = ActiveWorkbook.Path          '<<< g³ówna zmiana

    If Len(folderPath) = 0 Then
        MsgBox "Aktywny skoroszyt nie jest zapisany." & vbCrLf & _
               "Zapisz go w folderze z plikami do scalenia i uruchom makro ponownie.", _
               vbExclamation, "MergeExcelFilesInFolder"
        Exit Sub
    End If
       
    MergeCore folderPath
End Sub

'JAKIŒ PRZYK£ADOWY KOMENTARZ
Private Sub MergeCore(folderPath As String)

    Const OUT_NAME As String = "z³¹czony_plik.xlsx"
    
    Dim wbOut As Workbook
    Set wbOut = Workbooks.Add(xlWBATWorksheet)          '1 pusty arkusz
    Dim tempSheet As Worksheet: Set tempSheet = wbOut.Sheets(1)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim fname As String
    fname = Dir(folderPath & "\*.xls*")                 'xls, xlsx, xlsm
    
    Do While Len(fname) > 0
        If fname <> OUT_NAME And Left$(fname, 2) <> "~$" Then     'pomiñ wynik i plik tmp
            Debug.Print "Kopiujê: "; fname
            
            Dim wbIn As Workbook
            Set wbIn = Workbooks.Open(folderPath & "\" & fname, ReadOnly:=True)
            
            wbIn.Sheets(1).Copy Before:=wbOut.Sheets(1)            'skopiuj na pocz¹tek
            Dim newSh As Worksheet: Set newSh = wbOut.Sheets(1)
            
            '--- unikalna nazwa arkusza ---------------------------
            Dim base$, candidate$, i&
            base = Left$(Split(fname, " ")(0), 31)                 'pierwszy wyraz, max 31
            candidate = base: i = 1
            Do While SheetNameExists(wbOut, candidate)
                candidate = Left$(base, 31 - Len("_" & i)) & "_" & i
                i = i + 1
            Loop
            newSh.Name = candidate
            
            wbIn.Close SaveChanges:=False
        End If
        fname = Dir
    Loop
    
    '--- usuñ pusty startowy, jeœli zosta³y inne arkusze ----------
    If wbOut.Sheets.Count > 1 Then tempSheet.Delete
    
    '--- zapisz wynik ---------------------------------------------
    Dim outFull As String
    outFull = folderPath & "\" & OUT_NAME
    
    On Error Resume Next
    Kill outFull                                'usuñ istniej¹cy plik
    On Error GoTo 0
    
    wbOut.SaveAs Filename:=outFull, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "?? Po³¹czono arkusze!" & vbCrLf & _
           "Plik wynikowy: " & outFull, vbInformation
End Sub

'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦ POMOCNICZA: sprawdŸ unikalnoœæ nazwy ¦¦¦¦¦¦¦¦¦¦¦¦¦
Private Function SheetNameExists(wb As Workbook, nm As String) As Boolean
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If LCase$(sh.Name) = LCase$(nm) Then
            SheetNameExists = True
            Exit Function
        End If
    Next sh
End Function
'==================================================================


