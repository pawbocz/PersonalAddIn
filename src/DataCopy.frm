VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataCopy 
   Caption         =   "Szczegó³y dzia³ania"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "DataCopy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' DataCopy --------------------------
Public idCol As String
Public OpisCol As String
Public JednCol As String
Public PrzedmCol As String
Public FormOK As Boolean

Private Sub btnOK_Click()
    idCol = UCase(Trim(txtID.Text))
    OpisCol = UCase(Trim(txtOpis.Text))
    JednCol = UCase(Trim(txtJedn.Text))
    PrzedmCol = UCase(Trim(txtPrzedm.Text))
    
    If idCol = "" Or OpisCol = "" Or JednCol = "" Or PrzedmCol = "" Then
        MsgBox "Wszystkie pola musz¹ byæ wype³nione.", vbExclamation
        Exit Sub
    End If
    
    FormOK = True
    Me.Hide
End Sub

Private Sub btnAnuluj_Click()
    FormOK = False
    Me.Hide
End Sub

