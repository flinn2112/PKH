VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddressForm 
   Caption         =   "Adressen bearbeiten"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   OleObjectBlob   =   "AddressForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "AddressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit
Rem *******************************
Rem OK-Button
Rem *******************************
Private Sub CommandButton1_Click()
    'BUG  SaveProperties
    Unload AddressForm
End Sub
Rem *******************************
Rem Abbrechen-Button
Rem *******************************
Private Sub CommandButton2_Click()
    Unload AddressForm
End Sub
Rem *******************************
Rem Übernehmen-Button
Rem *******************************
Private Sub CommandButton3_Click()
    'BUG  SaveProperties
End Sub


