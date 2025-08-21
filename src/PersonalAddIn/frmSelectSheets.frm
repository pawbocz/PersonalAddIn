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


Private pFormOK        As Boolean
Private pSelected      As Collection


Public Property Get FormOK() As Boolean
    FormOK = pFormOK
End Property

Public Property Get SelectedSheets() As Collection
    Set SelectedSheets = pSelected
End Property




Public Sub Init(wsCol As Sheets)
    
    Dim sh As Worksheet
    Me.lstSheets.Clear
    
    For Each sh In wsCol
        Me.lstSheets.AddItem sh.Name
    Next sh
    
    If Me.lstSheets.ListCount > 0 Then _
        Me.lstSheets.Selected(0) = True
    
End Sub


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


Private Sub cmdCancel_Click()
    pFormOK = False
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    On Error Resume Next
    If Application.PixelsPerInch <> 96 Then _
        Me.Zoom = 100 * Application.PixelsPerInch / 96
    On Error GoTo 0
    
    pFormOK = False
End Sub
