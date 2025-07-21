VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinkiFormularz 
   Caption         =   "Linki do zostawienia"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   OleObjectBlob   =   "LinkiFormularz.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LinkiFormularz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btnDodaj_Click()
    Dim wpis As String
    wpis = Trim(txtLinkFragment.Text)
    
    If wpis <> "" Then
        lstZachowane.AddItem wpis
        txtLinkFragment.Text = ""
        txtLinkFragment.SetFocus
    Else
        MsgBox "Wpisz fragment linku, kt�ry chcesz zachowa�.", vbExclamation
    End If
End Sub

Private Sub btnStart_Click()
    Dim keepList() As String
    Dim i As Integer
    Dim linki As Variant
    Dim zostaw As Boolean
    Dim fragment As Variant
    
    If lstZachowane.ListCount = 0 Then
        MsgBox "Nie dodano �adnych fragment�w link�w do zachowania.", vbExclamation
        Exit Sub
    End If
    
    ' Stw�rz tablic� z ListBoxa
    ReDim keepList(lstZachowane.ListCount - 1)
    For i = 0 To lstZachowane.ListCount - 1
        keepList(i) = LCase(lstZachowane.List(i))
    Next i

    ' Pobierz linki
    linki = ActiveWorkbook.LinkSources(xlLinkTypeExcelLinks)
    
    If Not IsEmpty(linki) Then
        For i = LBound(linki) To UBound(linki)
            zostaw = False
            
            For Each fragment In keepList
                If InStr(LCase(linki(i)), fragment) > 0 Then
                    zostaw = True
                    Exit For
                End If
            Next fragment
            
            If Not zostaw Then
                ActiveWorkbook.BreakLink Name:=linki(i), Type:=xlLinkTypeExcelLinks
                Debug.Print "Usuni�to link: " & linki(i)
            Else
                Debug.Print "Zachowano link: " & linki(i)
            End If
        Next i
        MsgBox "Proces zako�czony. Sprawd� okno Immediate (Ctrl+G), by zobaczy� szczeg�y.", vbInformation
    Else
        MsgBox "Nie znaleziono �adnych zewn�trznych link�w.", vbInformation
    End If
    
    Unload Me
End Sub

Private Sub btnAnuluj_Click()
    Unload Me
End Sub

