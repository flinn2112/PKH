Attribute VB_Name = "SANA_INIT"
Private Function PropertyExists(ByVal PropertyName As String) As Boolean

    Dim p As DocumentProperty

    On Error Resume Next
    Set p = ActiveDocument.CustomDocumentProperties(PropertyName)

    PropertyExists = (Err.Number = 0)

    Err.Clear
    On Error GoTo 0

End Function


Private Sub EnsureDocumentProperty(ByVal PropertyName As String)

    If Not PropertyExists(PropertyName) Then

        ActiveDocument.CustomDocumentProperties.Add _
            name:=PropertyName, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=""

    End If

End Sub

Public Sub InitCustomDocumentProperties()

    Dim i As Integer

    EnsureDocumentProperty "IsSaved"

    For i = 1 To 7

        EnsureDocumentProperty "name" & i
        EnsureDocumentProperty "anrede" & i
        EnsureDocumentProperty "ansprache" & i
        EnsureDocumentProperty "Firma" & i
        EnsureDocumentProperty "strass" & i
        EnsureDocumentProperty "plz" & i
        EnsureDocumentProperty "ort" & i

    Next i

End Sub

