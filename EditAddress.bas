Attribute VB_Name = "EditAddress"
Option Explicit
Rem *******************************************************************
Rem Auslesen der Felder aus den CustomDocumentProperties
Rem *******************************************************************

Sub EditAddress()
    
    Dim ad As Object
    Dim af As Object
    
    Set ad = ActiveDocument
    Set af = AddressForm
    
    Load AddressForm
    
    With af
        .TextBox1.Value = ad.CustomDocumentProperties("anrede1").Value
        .TextBox2.Value = ad.CustomDocumentProperties("name1").Value
        .TextBox3.Value = ad.CustomDocumentProperties("firma1").Value
        .TextBox4.Value = ad.CustomDocumentProperties("strass1").Value
        .TextBox5.Value = ad.CustomDocumentProperties("plz1").Value
        .TextBox6.Value = ad.CustomDocumentProperties("ort1").Value
        .TextBox7.Value = ad.CustomDocumentProperties("land1").Value
        .TextBox8.Value = ad.CustomDocumentProperties("ansprache1").Value
    
        .TextBox9.Value = ad.CustomDocumentProperties("anrede2").Value
        .TextBox10.Value = ad.CustomDocumentProperties("name2").Value
        .TextBox11.Value = ad.CustomDocumentProperties("firma2").Value
        .TextBox12.Value = ad.CustomDocumentProperties("strass2").Value
        .TextBox13.Value = ad.CustomDocumentProperties("plz2").Value
        .TextBox14.Value = ad.CustomDocumentProperties("ort2").Value
        .TextBox15.Value = ad.CustomDocumentProperties("land2").Value
        .TextBox16.Value = ad.CustomDocumentProperties("ansprache2").Value
    
        .TextBox17.Value = ad.CustomDocumentProperties("anrede3").Value
        .TextBox18.Value = ad.CustomDocumentProperties("name3").Value
        .TextBox19.Value = ad.CustomDocumentProperties("firma3").Value
        .TextBox20.Value = ad.CustomDocumentProperties("strass3").Value
        .TextBox21.Value = ad.CustomDocumentProperties("plz3").Value
        .TextBox22.Value = ad.CustomDocumentProperties("ort3").Value
        .TextBox23.Value = ad.CustomDocumentProperties("land3").Value
        .TextBox24.Value = ad.CustomDocumentProperties("ansprache3").Value
    
        .TextBox25.Value = ad.CustomDocumentProperties("anrede4").Value
        .TextBox26.Value = ad.CustomDocumentProperties("name4").Value
        .TextBox27.Value = ad.CustomDocumentProperties("firma4").Value
        .TextBox28.Value = ad.CustomDocumentProperties("strass4").Value
        .TextBox29.Value = ad.CustomDocumentProperties("plz4").Value
        .TextBox30.Value = ad.CustomDocumentProperties("ort4").Value
        Rem: .TextBox31.Value = ad.CustomDocumentProperties("land4").Value
        .TextBox32.Value = ad.CustomDocumentProperties("ansprache4").Value
    
        .TextBox33.Value = ad.CustomDocumentProperties("anrede5").Value
        .TextBox34.Value = ad.CustomDocumentProperties("name5").Value
        .TextBox35.Value = ad.CustomDocumentProperties("firma5").Value
        .TextBox36.Value = ad.CustomDocumentProperties("strass5").Value
        .TextBox37.Value = ad.CustomDocumentProperties("plz5").Value
        .TextBox38.Value = ad.CustomDocumentProperties("ort5").Value
        Rem: .TextBox39.Value = ad.CustomDocumentProperties("land5").Value
        .TextBox40.Value = ad.CustomDocumentProperties("ansprache5").Value
    
        .TextBox41.Value = ad.CustomDocumentProperties("anrede6").Value
        .TextBox42.Value = ad.CustomDocumentProperties("name6").Value
        .TextBox43.Value = ad.CustomDocumentProperties("firma6").Value
        .TextBox44.Value = ad.CustomDocumentProperties("strass6").Value
        .TextBox45.Value = ad.CustomDocumentProperties("plz6").Value
        .TextBox46.Value = ad.CustomDocumentProperties("ort6").Value
        Rem: .TextBox47.Value = ad.CustomDocumentProperties("land6").Value
        .TextBox48.Value = ad.CustomDocumentProperties("ansprache6").Value
    
        .TextBox49.Value = ad.CustomDocumentProperties("anrede7").Value
        .TextBox50.Value = ad.CustomDocumentProperties("name7").Value
        .TextBox51.Value = ad.CustomDocumentProperties("firma7").Value
        .TextBox52.Value = ad.CustomDocumentProperties("strass7").Value
        .TextBox53.Value = ad.CustomDocumentProperties("plz7").Value
        .TextBox54.Value = ad.CustomDocumentProperties("ort7").Value
        Rem: .TextBox55.Value = ad.CustomDocumentProperties("land7").Value
        .TextBox56.Value = ad.CustomDocumentProperties("ansprache7").Value
    End With
    
    AddressForm.Show

End Sub

Rem *******************************************************************
Rem Rückgabe der Adressen in die Felder der CustomDocumentProperties
Rem *******************************************************************

Sub SaveProperties()
    
    Dim ad As Object
    Dim af As Object
    
    Set ad = ActiveDocument
    Set af = AddressForm
    
    With ad
    .CustomDocumentProperties("anrede1").Value = af.TextBox1.Value
    .CustomDocumentProperties("name1").Value = af.TextBox2.Value
    .CustomDocumentProperties("firma1").Value = af.TextBox3.Value
    .CustomDocumentProperties("strass1").Value = af.TextBox4.Value
    .CustomDocumentProperties("plz1").Value = af.TextBox5.Value
    .CustomDocumentProperties("ort1").Value = af.TextBox6.Value
    .CustomDocumentProperties("land1").Value = af.TextBox7.Value
    .CustomDocumentProperties("ansprache1").Value = af.TextBox8.Value
    
    .CustomDocumentProperties("anrede2").Value = af.TextBox9.Value
    .CustomDocumentProperties("name2").Value = af.TextBox10.Value
    .CustomDocumentProperties("firma2").Value = af.TextBox11.Value
    .CustomDocumentProperties("strass2").Value = af.TextBox12.Value
    .CustomDocumentProperties("plz2").Value = af.TextBox13.Value
    .CustomDocumentProperties("ort2").Value = af.TextBox14.Value
    .CustomDocumentProperties("land2").Value = af.TextBox15.Value
    .CustomDocumentProperties("ansprache2").Value = af.TextBox16.Value
    
    .CustomDocumentProperties("anrede3").Value = af.TextBox17.Value
    .CustomDocumentProperties("name3").Value = af.TextBox18.Value
    .CustomDocumentProperties("firma3").Value = af.TextBox19.Value
    .CustomDocumentProperties("strass3").Value = af.TextBox20.Value
    .CustomDocumentProperties("plz3").Value = af.TextBox21.Value
    .CustomDocumentProperties("ort3").Value = af.TextBox22.Value
    Rem: .CustomDocumentProperties("land3").Value = af.TextBox23.Value
    .CustomDocumentProperties("ansprache3").Value = af.TextBox24.Value
    
    .CustomDocumentProperties("anrede4").Value = af.TextBox25.Value
    .CustomDocumentProperties("name4").Value = af.TextBox26.Value
    .CustomDocumentProperties("firma4").Value = af.TextBox27.Value
    .CustomDocumentProperties("strass4").Value = af.TextBox28.Value
    .CustomDocumentProperties("plz4").Value = af.TextBox29.Value
    .CustomDocumentProperties("ort4").Value = af.TextBox30.Value
    Rem: .CustomDocumentProperties("land4").Value = af.TextBox31.Value
    .CustomDocumentProperties("ansprache4").Value = af.TextBox32.Value
    
    .CustomDocumentProperties("anrede5").Value = af.TextBox33.Value
    .CustomDocumentProperties("name5").Value = af.TextBox34.Value
    .CustomDocumentProperties("firma5").Value = af.TextBox35.Value
    .CustomDocumentProperties("strass5").Value = af.TextBox36.Value
    .CustomDocumentProperties("plz5").Value = af.TextBox37.Value
    .CustomDocumentProperties("ort5").Value = af.TextBox38.Value
    Rem: .CustomDocumentProperties("land5").Value = af.TextBox39.Value
    .CustomDocumentProperties("ansprache5").Value = af.TextBox40.Value
    
    .CustomDocumentProperties("anrede6").Value = af.TextBox41.Value
    .CustomDocumentProperties("name6").Value = af.TextBox42.Value
    .CustomDocumentProperties("firma6").Value = af.TextBox43.Value
    .CustomDocumentProperties("strass6").Value = af.TextBox44.Value
    .CustomDocumentProperties("plz6").Value = af.TextBox45.Value
    .CustomDocumentProperties("ort6").Value = af.TextBox46.Value
    Rem:.CustomDocumentProperties("land6").Value = af.TextBox47.Value
    .CustomDocumentProperties("ansprache6").Value = af.TextBox48.Value
    
    .CustomDocumentProperties("anrede7").Value = af.TextBox49.Value
    .CustomDocumentProperties("name7").Value = af.TextBox50.Value
    .CustomDocumentProperties("firma7").Value = af.TextBox51.Value
    .CustomDocumentProperties("strass7").Value = af.TextBox52.Value
    .CustomDocumentProperties("plz7").Value = af.TextBox53.Value
    .CustomDocumentProperties("ort7").Value = af.TextBox54.Value
    Rem: .CustomDocumentProperties("land7").Value = af.TextBox55.Value
    .CustomDocumentProperties("ansprache7").Value = af.TextBox56.Value
    
    .Fields.Update
    End With
    Erstschriftlich
    
End Sub
Sub InsertDate()
'
' Eingefügt am 10.09.02 von HS
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "[DATUM]"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    If Selection = "[DATUM]" Then
        Selection.InsertDateTime DateTimeFormat:="dd. MMMM yyyy", InsertAsField:=False, _
        DateLanguage:=wdGerman, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    End If
    
End Sub



