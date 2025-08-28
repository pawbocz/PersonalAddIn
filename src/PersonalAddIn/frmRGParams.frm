VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRGParams 
   Caption         =   "UserForm1"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7995
   OleObjectBlob   =   "frmRGParams.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRGParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormOK As Boolean
Public OutCol As Long
Public CatCol As Long
Public DescCol As Long
Public firstRow As Long

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next

    If Application.PixelsPerInch <> 96 Then Me.Zoom = 100 * Application.PixelsPerInch / 96
    On Error GoTo 0


    FormOK = False
End Sub

Private Sub cmdOK_Click()
    Dim oCol As Long, cCol As Long, dCol As Long, fRow As Long

    oCol = ColTextToNumber(Trim$(Me.txtOutCol.Text))
    cCol = ColTextToNumber(Trim$(Me.txtCatCol.Text))
    dCol = ColTextToNumber(Trim$(Me.txtDescCol.Text))

    If Not IsNumeric(Me.txtFirstRow.Text) Then
        MsgBox "Podaj liczbowy pierwszy wiersz (np. 8).", vbExclamation: Exit Sub
    End If
    fRow = CLng(Me.txtFirstRow.Text)

    If oCol <= 0 Or cCol <= 0 Or dCol <= 0 Or fRow <= 0 Then
        MsgBox "Uzupe³nij poprawnie wszystkie pola.", vbExclamation
        Exit Sub
    End If

    OutCol = oCol: CatCol = cCol: DescCol = dCol: firstRow = fRow
    FormOK = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    FormOK = False
    Me.Hide
End Sub


Private Function ColTextToNumber(ByVal s As String) As Long
    Dim i As Long, ch As String, n As Long
    If Len(s) = 0 Then Exit Function

    If IsNumeric(s) Then
        ColTextToNumber = CLng(s)
        Exit Function
    End If

    s = UCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "A" Or ch > "Z" Then
            ColTextToNumber = 0
            Exit Function
        End If
        n = n * 26 + (Asc(ch) - 64)
    Next i
    ColTextToNumber = n
End Function

