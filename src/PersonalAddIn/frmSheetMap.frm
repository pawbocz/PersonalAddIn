VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetMap 
   Caption         =   "UserForm jakiœ tam 1"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9300.001
   OleObjectBlob   =   "frmSheetMap.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSheetMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Public FormOK As Boolean
Public pairs  As Collection


Private mColLp     As Long
Private mColOpis   As Long
Private mColJedn   As Long
Private mColPrzedm As Long
Private mStartRow  As Long

Public Property Get hdrRow() As Long
    hdrRow = val(Me.txtHdrRow.value)
End Property

Public Property Get UseCustomCols() As Boolean
    UseCustomCols = Me.chkCustom.value
End Property

Public Property Get colLp() As Long:     colLp = mColLp:         End Property
Public Property Get colOpis() As Long:   colOpis = mColOpis:     End Property
Public Property Get colJedn() As Long:   colJedn = mColJedn:     End Property
Public Property Get colPrzedm() As Long: colPrzedm = mColPrzedm: End Property
Public Property Get startRow() As Long:  startRow = mStartRow:   End Property


Private Sub Label6_Click()

End Sub


Private Sub UserForm_Initialize()

    
    On Error Resume Next
    If Application.PixelsPerInch <> 96 Then _
        Me.Zoom = 100 * Application.PixelsPerInch / 96
    On Error GoTo 0
 
    Dim wbSrc As Workbook: Set wbSrc = gSourceWB
    Dim wbTgt As Workbook: Set wbTgt = gTargetWB

    Dim sh As Worksheet
    For Each sh In wbSrc.Worksheets
        lstSrc.AddItem sh.Name
    Next sh
    For Each sh In wbTgt.Worksheets
        lstTgt.AddItem sh.Name
    Next sh

    
    txtLp.value = "2"
    txtOpis.value = "3"
    txtJedn.value = "4"
    txtPrzedm.value = "5"
    txtStart.value = "8"

    Set pairs = New Collection
    FormOK = False
End Sub



Private Sub chkCustom_Click()
    Dim enab As Boolean: enab = Me.chkCustom.value
    
    Me.txtLp.Enabled = enab
    Me.txtOpis.Enabled = enab
    Me.txtJedn.Enabled = enab
    Me.txtPrzedm.Enabled = enab
    Me.txtStart.Enabled = enab
End Sub

Private Sub btnAdd_Click()
    If lstSrc.ListIndex = -1 Or lstTgt.ListIndex = -1 Then Exit Sub

    Dim srcName$, tgtName$
    srcName = lstSrc.List(lstSrc.ListIndex)
    tgtName = lstTgt.List(lstTgt.ListIndex)

    lstPairs.AddItem srcName & "  ›  " & tgtName
    pairs.Add Array(srcName, tgtName)
End Sub

Private Sub btnRemove_Click()
    If lstPairs.ListIndex = -1 Then Exit Sub
    pairs.Remove lstPairs.ListIndex + 1
    lstPairs.RemoveItem lstPairs.ListIndex
End Sub

Private Sub btnAddNewTgt_Click()
    Dim nm$: nm = Trim(txtNewTgt.Text)
    If nm = "" Then Exit Sub

    Dim sh As Worksheet
    For Each sh In gTargetWB.Worksheets
        If LCase(sh.Name) = LCase(nm) Then
            MsgBox "Arkusz '" & nm & "' ju¿ istnieje.", vbExclamation
            Exit Sub
        End If
    Next sh

    gTargetWB.Worksheets.Add(After:=gTargetWB.Sheets(gTargetWB.Sheets.Count)).Name = nm
    lstTgt.AddItem nm
    txtNewTgt.Text = ""
End Sub


Private Sub btnStart_Click()

    If pairs.Count = 0 Then
        MsgBox "Nie zdefiniowano ¿adnych par arkuszy.", vbExclamation
        Exit Sub
    End If


    On Error GoTo BadInput
    mColLp = CLng(txtLp.value)
    mColOpis = CLng(txtOpis.value)
    mColJedn = CLng(txtJedn.value)
    mColPrzedm = CLng(txtPrzedm.value)
    mStartRow = CLng(txtStart.value)

    If Application.min(mColLp, mColOpis, mColJedn, mColPrzedm, mStartRow) < 1 _
       Or Application.Max(mColLp, mColOpis, mColJedn, mColPrzedm) > 16384 Then
        Err.Raise 1
    End If

    FormOK = True
    Me.Hide
    Exit Sub

BadInput:
    MsgBox "Nieprawid³owe wartoœci kolumn / wiersza." & vbCrLf & _
           "Podaj liczby z zakresu 1-16384.", vbExclamation
End Sub

Private Sub btnCancel_Click()
    FormOK = False
    Me.Hide
End Sub

Private Sub MultiPage1_Change()
    'nic
End Sub



