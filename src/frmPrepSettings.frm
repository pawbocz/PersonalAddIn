VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrepSettings 
   Caption         =   "UserForm1"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8580.001
   OleObjectBlob   =   "frmPrepSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrepSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FormOK As Boolean

Public hdrRow As Long, colLp As Long, colOpis As Long
Public colJedn As Long, colPrzedm As Long, FirstData As Long


Private Sub cmdOK_Click()
    On Error GoTo Bad
    hdrRow = CLng(txtHdrRow.Value)
    FirstData = CLng(IIf(txtDataRow.Value = "", hdrRow + 1, txtDataRow.Value))
    colLp = ColToNum(txtColLp.Value)
    colOpis = ColToNum(txtColOpis.Value)
    colJedn = ColToNum(txtColJedn.Value)
    colPrzedm = ColToNum(txtColPrzedm.Value)
    FormOK = True
    Me.Hide
    Exit Sub
Bad:
    MsgBox "Sprawdü podane liczby / litery kolumn.", vbExclamation
End Sub


Private Sub cmdCancel_Click()
    FormOK = False
    Me.Hide
End Sub

'ó Aõ1, Bõ2, Ö lub liczba õ liczba ---------------------------
Private Function ColToNum(v As Variant) As Long
    If IsNumeric(v) Then
        ColToNum = CLng(v)
    Else
        ColToNum = Columns(UCase$(Trim$(v))).Column
    End If
End Function

