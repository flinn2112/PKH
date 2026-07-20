VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Auswahl 
   Caption         =   "Datenauswahl"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   OleObjectBlob   =   "Auswahl.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Auswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Base 0
Option Explicit

Dim labor_daten As Collection

Private Sub CommandButton1_Click()
Dim a As Long
Dim X As Page, i As Long

   a = 1
   For i = a To 10
     MultiPage1.Pages.Item(i).Visible = False
   Next i
   
End Sub

Private Sub CommandButton2_Click()
Dim i As Long
    For i = 0 To 4
        MultiPage1.Pages.Item(i).Caption = "Hugo" & i
    Next i
End Sub


Private Sub CommandButton3_Click()
    Dim i As Long
    For i = 0 To LB_Labor.ListCount - 1
        labor_daten(LB_Labor.List(i, 3)).Add IIf(LB_Labor.Selected(i), 1, 0), "gewaehlt"
    Next i
    Me.Hide
End Sub

Private Sub CommandButton4_Click()
Dim i As Long
  For i = 0 To LB_Labor.ListCount - 1
     LB_Labor.Selected(i) = True
  Next i
End Sub

Private Sub CommandButton5_Click()
Dim i As Long
  For i = 0 To LB_Labor.ListCount - 1
     LB_Labor.Selected(i) = False
  Next i
End Sub

Private Sub UserForm_Initialize()
    ' tue nix
End Sub

Friend Property Set daten(d As Collection)
    Set labor_daten = d
End Property

Private Sub UserForm_Activate()
    Dim i As Long
    Dim y As Long
    Dim colLIKey As String
    Dim colLKey As String
    Dim zeilen As Long
    Dim zeile

    '*** Hide all Tabs not used
    CommandButton1_Click
    
    If labor_daten Is Nothing Then
       Exit Sub
    End If
    
    zeilen = 0
    For Each zeile In labor_daten
        LB_Labor.AddItem ISHMED_Datum(zeile("N2LADATUM").Item("var") & "") & " / " & ISHMED_Uhrzeit(zeile("N2LATIME").Item("var") & "")
        LB_Labor.Column(1, zeilen) = zeile("N2KATTEXT").Item("var")
        LB_Labor.Column(2, zeilen) = zeile("N2VALUE").Item("var")
        LB_Labor.Column(3, zeilen) = zeile("N2UNIT").Item("var")
        ' Mit diesem Schluessel findest Du die zeile
        ' nachher in der Collection labor_daten wieder
        LB_Labor.Column(3, zeilen) = zeile("row_key").Item("var")
        zeilen = zeilen + 1
    Next
End Sub
