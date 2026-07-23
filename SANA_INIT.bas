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

Public Sub SetPersistentDate(ByVal strVariableName As String, _
                             ByVal strBookmarkName As String)

    Dim strDate As String
    Dim rng As Word.Range

    ' Try to read our own persistent value
    On Error Resume Next
    strDate = ActiveDocument.Variables(strVariableName).Value
    On Error GoTo 0

    ' First creation: store today's date
    If Trim$(strDate) = "" Then

        strDate = Format$(Date, "dd.MM.yyyy")

        ' Replace the variable if it already exists
        On Error Resume Next
        ActiveDocument.Variables(strVariableName).Delete
        On Error GoTo 0

        ActiveDocument.Variables.Add _
            name:=strVariableName, _
            Value:=strDate
    End If

    ' Fill the bookmark
    If ActiveDocument.Bookmarks.exists(strBookmarkName) Then

        Set rng = ActiveDocument.Bookmarks(strBookmarkName).Range
        rng.text = strDate

        ' Word deletes a bookmark when text is assigned
        ActiveDocument.Bookmarks.Add _
            name:=strBookmarkName, _
            Range:=rng

    End If

End Sub

