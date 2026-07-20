VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MeldungenForm 
   Caption         =   "Fehlende Datenübernahme"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4650
   OleObjectBlob   =   "MeldungenForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "MeldungenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
    Dim z, i, j, m As String
    With Me.MeldungsListe
        z = .ListCount - 1
        For i = 0 To z
            For j = 0 To 2
                If j > 0 Then m = m + vbTab
                m = m & .List(i, j)
            Next j
            m = m & vbCrLf
        Next i
    End With
    ClipBoard_SetData m
End Sub

Private Sub OK_Click()
    Me.Hide
End Sub


