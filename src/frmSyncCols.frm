VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSyncCols 
   Caption         =   "UserForm1"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7815
   OleObjectBlob   =   "frmSyncCols.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSyncCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public FormOK As Boolean
Private mLVCena  As Long, mLVWart  As Long
Private mSRCCena As Long, mSRCWart As Long

Public Property Get LV_Cena() As Long:   LV_Cena = mLVCena:   End Property
Public Property Get LV_Wart() As Long:   LV_Wart = mLVWart:   End Property
Public Property Get SRC_Cena() As Long:  SRC_Cena = mSRCCena: End Property
Public Property Get SRC_Wart() As Long:  SRC_Wart = mSRCWart: End Property



Private Sub UserForm_Initialize()
    On Error Resume Next
    If Application.PixelsPerInch <> 96 Then _
        Me.Zoom = 100 * Application.PixelsPerInch / 96
    On Error GoTo 0
 
End Sub


Private Function ColIndex(txt As String) As Long
    Dim s As String: s = UCase$(Trim$(txt))
    If s = "" Then ColIndex = 0: Exit Function

    If IsNumeric(s) Then
        ColIndex = CLng(s)
        Exit Function
    End If
    
  
    Dim i As Long, n As Long
    For i = 1 To Len(s)
        Dim ch As Integer: ch = Asc(Mid$(s, i, 1))
        If ch < 65 Or ch > 90 Then
            ColIndex = 0: Exit Function
        End If
        n = n * 26 + (ch - 64)
    Next i
    ColIndex = n
End Function


Private Sub cmdOK_Click()
    Dim ok As Boolean
    mLVCena = ColIndex(Me.txtLV_Cena.Text)
    mLVWart = ColIndex(Me.txtLV_Wart.Text)
    mSRCCena = ColIndex(Me.txtSRC_Cena.Text)
    mSRCWart = ColIndex(Me.txtSRC_Wart.Text)
    
    ok = (mLVCena > 0 And mLVWart > 0 And _
          mSRCCena > 0 And mSRCWart > 0)
    
    If ok Then
        FormOK = True
        Me.Hide
    Else
        MsgBox "Podaj poprawne kolumny (litery lub numery).", vbExclamation
    End If
End Sub

Private Sub cmdCancel_Click()
    FormOK = False
    Me.Hide
End Sub



