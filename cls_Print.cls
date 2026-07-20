VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klassenmodul: cls_Print
 Public WithEvents oApp As Application
Attribute oApp.VB_VarHelpID = -1

 Private Sub oApp_DocumentBeforePrint(ByVal doc As Document, Cancel As Boolean)
 ' MsgBox "Hinweis vor dem Drucken" & vbCrLf & "Mit 'Cancel=True' kann der Druckbefehl abgebrochen werden", vbInformation
 ' Cancel = True ' zum Ausdrucken Zeile entfernen
 End Sub

