VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectSheets 
   Caption         =   "Prep data masówka"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelectSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=== ZMIENNE PRYWATNE ==========================================
Private pFormOK        As Boolean          'status OK/Cancel
Private pSelected      As Collection       'zwracane arkusze

'=== W£AŒCIWOŒCI PUBLICZNE =====================================
Public Property Get FormOK() As Boolean
    FormOK = pFormOK
End Property

Public Property Get SelectedSheets() As Collection
    Set SelectedSheets = pSelected
End Property

'=== METODA INIT  (wo³ana z modu³u) =============================
'  Przeka¿ kolekcjê arkuszy (np. ActiveWorkbook.Worksheets)
Public Sub Init(wsCol As Sheets)
    
    Dim sh As Worksheet
    Me.lstSheets.Clear
    
    For Each sh In wsCol
        Me.lstSheets.AddItem sh.Name
    Next sh
    
    If Me.lstSheets.ListCount > 0 Then _
        Me.lstSheets.Selected(0) = True      'zaznacz pierwszy
    
End Sub

'=== PRZYCISK  OK  =============================================
Private Sub cmdOK_Click()

    Dim i As Long
    
    Set pSelected = New Collection
    For i = 0 To Me.lstSheets.ListCount - 1
        If Me.lstSheets.Selected(i) Then
            pSelected.Add ActiveWorkbook.Worksheets(Me.lstSheets.List(i))
        End If
    Next i
    
    If pSelected.Count = 0 Then
        MsgBox "Nie wybrano ¿adnego arkusza.", vbExclamation
        Exit Sub
    End If
    
    pFormOK = True
    Me.Hide
End Sub

'=== PRZYCISK  Anuluj  =========================================
Private Sub cmdCancel_Click()
    pFormOK = False
    Me.Hide
End Sub

'=== ZA£ADOWANIE FORMA (opc.) ==================================
Private Sub UserForm_Initialize()
    pFormOK = False
End Sub


