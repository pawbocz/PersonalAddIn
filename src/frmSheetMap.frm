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
'===========================  frmSheetMap  =========================
Option Explicit

'––––––––––––––––  Z M I E N N E   M O D U £ U  –––––––––––––––––––
'   (musz¹ byæ NAD pierwsz¹ procedur¹!)

'wyniki, których oczekuje MainCopy
Public FormOK As Boolean
Public pairs  As Collection          'each item = Array(srcName, tgtName)

'ustawienia kolumn LV (zak³adka 2)
Private mColLp     As Long
Private mColOpis   As Long
Private mColJedn   As Long
Private mColPrzedm As Long
Private mStartRow  As Long
'------------------------------------------------------------------

'––––––––––––––––  P R O P E R T Y   G E T  –––––––––––––––––––––––
'zak³adka 1 – wiersz nag³ówków
Public Property Get hdrRow() As Long
    hdrRow = val(Me.txtHdrRow.Value)          '0, gdy pole puste
End Property

Public Property Get UseCustomCols() As Boolean
    UseCustomCols = Me.chkCustom.Value
End Property

'zak³adka 2 – kolumny LV + pierwszy wiersz danych
Public Property Get colLp() As Long:     colLp = mColLp:         End Property
Public Property Get colOpis() As Long:   colOpis = mColOpis:     End Property
Public Property Get colJedn() As Long:   colJedn = mColJedn:     End Property
Public Property Get colPrzedm() As Long: colPrzedm = mColPrzedm: End Property
Public Property Get startRow() As Long:  startRow = mStartRow:   End Property
'------------------------------------------------------------------

Private Sub Label6_Click()

End Sub

'==================================================================
'  I N I T
'==================================================================
Private Sub UserForm_Initialize()

    '––––  AUTOSKALA DPI  ––––––––––––––––––––––––––––––––––––––
    On Error Resume Next                      'Excel ?2016 nie ma PixelsPerInch
    If Application.PixelsPerInch <> 96 Then _
        Me.Zoom = 100 * Application.PixelsPerInch / 96
    On Error GoTo 0
    '––––  KONIEC AUTOSKALI  –––––––––––––––––––––––––––––––––––

    Dim wbSrc As Workbook: Set wbSrc = gSourceWB
    Dim wbTgt As Workbook: Set wbTgt = gTargetWB

    Dim sh As Worksheet
    For Each sh In wbSrc.Worksheets
        lstSrc.AddItem sh.Name
    Next sh
    For Each sh In wbTgt.Worksheets
        lstTgt.AddItem sh.Name
    Next sh

    '--- domyœlne kolumny LV (standard) --------------------------
    txtLp.Value = "2"         'B
    txtOpis.Value = "3"       'C
    txtJedn.Value = "4"       'D
    txtPrzedm.Value = "5"     'E
    txtStart.Value = "8"      'pierwszy wiersz danych

    Set pairs = New Collection
    FormOK = False
End Sub

'==================================================================
'  P A R O W A N I E   A R K U S Z Y
'==================================================================

Private Sub chkCustom_Click()
    Dim enab As Boolean: enab = Me.chkCustom.Value
    
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

'==================================================================
'  O K   /   A N U L U J
'==================================================================
Private Sub btnStart_Click()

    If pairs.Count = 0 Then
        MsgBox "Nie zdefiniowano ¿adnych par arkuszy.", vbExclamation
        Exit Sub
    End If

    '–– odczyt i walidacja kolumn LV ––––––––––––––––––––––––––––
    On Error GoTo BadInput
    mColLp = CLng(txtLp.Value)
    mColOpis = CLng(txtOpis.Value)
    mColJedn = CLng(txtJedn.Value)
    mColPrzedm = CLng(txtPrzedm.Value)
    mStartRow = CLng(txtStart.Value)

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
'==================================================================
'  P O M O C N I C Z E   –   brak akcji
'==================================================================
Private Sub MultiPage1_Change()
    'nic – zostaw puste
End Sub
'==================================================================


