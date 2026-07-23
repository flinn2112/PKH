Attribute VB_Name = "WC_Extended"
Option Base 0
Option Explicit

Global gcloLaborDaten As New Collection
Global gcolMedikation As New Collection
Global gcolErreger As New Collection
Global gcolBlutprod As New Collection
Global gcolLaborRanking As New Collection
Global gcolLaborAusschliessen As New Collection
Global gcolLaborUmbenennen As New Collection

Global gint_NO_LABInvestigation As Long
Global last_lab_date As String
Global last_lab_time As String
Global werte As Variant
Global wert_start As Long
Global wert_ende As Long
Global tabcount As Long
Global Meldungen As New Collection

Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
   ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "User32" () As Long
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) _
   As Long
Declare Function EmptyClipboard Lib "User32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
   ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
   As Long, ByVal hMem As Long) As Long
 
Const GHND = &H42
Const CF_TEXT = 1
Const MAXSIZE = 4096
Const dokvar = "00"

Global objWord As New cls_Print

Sub AutoOpen()
    Debug.Print "Starting AutoOpen."
    Set objWord.oApp = Word.Application
End Sub


Function ClipBoard_SetData(MyString As String)
   Dim hGlobalMemory As Long, lpGlobalMemory As Long
   Dim hClipMemory As Long, X As Long
 
   ' Allocate moveable global memory.
   '-------------------------------------------
   hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)
 
   ' Lock the block to get a far pointer
   ' to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)
 
   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)
 
   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      MsgBox "Could not unlock memory location. Copy aborted."
      GoTo OutOfHere2
   End If
 
   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      MsgBox "Could not open the Clipboard. Copy aborted."
      Exit Function
   End If
 
   ' Clear the Clipboard.
   X = EmptyClipboard()
 
   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
 
OutOfHere2:
 
   If CloseClipboard() = 0 Then
      MsgBox "Could not close Clipboard."
   End If
 
End Function

Public Sub berichte(wo, schluessel, wert)
    Meldungen.Add Array(wo, schluessel, wert)
End Sub

Public Sub meldungen_anzeigen()
    If Meldungen.Count = 0 Then Exit Sub
    Dim itm, m As MeldungenForm
    Set m = New MeldungenForm
    With m.MeldungsListe
        For Each itm In Meldungen
            Dim zeile
            zeile = .ListCount
            .AddItem
            .List(zeile, 0) = itm(0)
            .List(zeile, 1) = itm(1)
        Next
    End With
    m.Show 1
    Set m = Nothing
End Sub

Public Function getTextLine(col As Collection)
    Dim zeile, langtext, vorige As String
    Dim itm
    Dim t As String
    Dim i, j, l, anzahl As Long
       
    anzahl = col.Count
    If anzahl = 1 Then
        zeile = col(1)
    Else
        For i = 1 To anzahl - 1
            vorige = col(i)
            zeile = col(i + 1)
            j = InStr(1, zeile, " ")
            ' keine Space gefunden, dann muss die ganze zeile
            ' noch mit drauf passen.
            If j = 0 Then j = Len(zeile)
            If Len(vorige) + j > 72 Then
               langtext = langtext & vorige & " "
            Else
               langtext = langtext & vorige & vbCrLf
            End If
        Next
    End If
    getTextLine = langtext + zeile
End Function

Public Sub initglobValues()
'***********************************************************************
'* alle speicherinhalte werden initalisiert
'***********************************************************************
    Dim i As Long
    Dim sortierung As Variant
    Dim ausschliessen As Variant
    Dim umbenennen As Variant
    ' In dieser Reihenfolge werden die Laborwerte gewuenscht!
    sortierung = Array("SPECIMEN", "HB (HGB)", "RBC", "WBC", "NEUT %", "LYMPH %", "MONO %", "EOS %", "BASO %", _
            "NEUT #", "LYMPH #", "MONO #", "EOS #", "BASO #", "HCT", "MCV", "MPV", "MCH", "MCHC", _
            "RDW", "PLT", "MP", "KREATININ", "GFRKREA", "eGFR", "HARNSTOFF", "Harnsäure", "NATRIUM", "KALIUM", _
            "CALCIUM", "PO4", "CHLORID", "MAGNESIUM", "CRP", "TROPONIN I", "CK", "CK-MB", _
            "GOT", "GPT", "LDH", "freies Hb", "Haptoglobin", "Retikulocyten", "Reti -Hb", "Retikulozyten", _
            "BILI (TOT.)", "GGT", "ALP", _
            "AP", "LAP", "AMYLASE", "LIPASE", "TP", "ALBUMIN", "GLUCOSE", "HBA1C", "Eisen", "Ferritin", _
            "Transferrin", "CHOLESTERIN", "LDL-Cholesterin", "HDL-CHOL.", "TRIGLYZERIDE", _
            "Lipoprotein(a)", "TSH", "FT4", "FT3", "TPZ", "INR", "PTT", "FIBRINOGEN", "ATIII", _
            "DIGOXIN", "DIGITOXIN", "AMIODARON", "THEOPHYLLIN", "VANCOMYCIN", "GENTAMYCIN", _
            "TOBRAMYCIN", "CYCLOSPORIN A", "TACROLIMUS", "MYCOPHENOLAT", "Hepatitis B", "HBs-AG", _
            "Anti-HBs", "Anti-HBc", "HCV-AK", _
            "hier beginnt Labor Berlin", _
            "Eisen#", "Eisen Se", "Ferritin HP", "Transferrin HP", "Transferrin-Sättigung HP", _
            "Vitamin B12#", "Folsäure#")
    ausschliessen = Array("IG %", "IG #", "NRBC %", "NRBC #")
    umbenennen = Array(Array("_GFR", "eGFR"), Array("ANORGANISCHER PHOSPHOR", "PO4"))
    
  wert_start = LBound(sortierung)
  wert_ende = UBound(sortierung)
  
  With gcolLaborRanking
    For i = 1 To .Count
        .Remove 1
    Next
    For i = wert_start To wert_ende
      .Add i, sortierung(i)
    Next
  End With
  
  gint_NO_LABInvestigation = 0
  For i = 1 To gcloLaborDaten.Count
    gcloLaborDaten.Remove 1
  Next i
  
  With gcolLaborAusschliessen
    For i = 1 To .Count
      .Remove 1
    Next
    For i = LBound(ausschliessen) To UBound(ausschliessen)
        .Add ausschliessen(i), ausschliessen(i)
    Next
  End With
  
  With gcolLaborUmbenennen
    For i = 1 To .Count
        .Remove 1
    Next
    For i = LBound(umbenennen) To UBound(umbenennen)
        .Add umbenennen(i)(1), umbenennen(i)(0)
    Next
  End With
    
  For i = 1 To gcolMedikation.Count
    gcolMedikation.Remove 1
  Next i

  'Beginn (neu 11.11.05 von KD): Ergänzung um Init. des Objektes Blutprodukte/Antikörperstatus
  For i = 1 To gcolBlutprod.Count
    gcolBlutprod.Remove 1
  Next i
  'Ende Init. Blutprodukte/Antikörperstatus

  'Beginn (neu 28.06.12 von KD): Ergänzung um Init. des Objektes "Mikrobiologie (Erreger)"
  For i = 1 To gcolErreger.Count
    gcolErreger.Remove 1
  Next i
  'Ende Init. Mikrobiologie (Erreger)

'  Set gAllText = New clsAlleTexte
'  gAllText.ScanForText

End Sub

Private Function set_zahl(zahl As Variant)
Dim a, b As String
Dim c, d As Long

a = LTrim(zahl)
c = InStr(a, ".")
If (c > 0) Then
   set_zahl = Left$(a, c - 1) & "," & Right$(a, Len(a) - c)
Else
   set_zahl = a
End If
End Function

Private Sub anhaengen(ausgabe As String, neu As String)
    If ausgabe > "" Then ausgabe = ausgabe & ", "
    ausgabe = ausgabe & neu
End Sub


' Echokardiographie-Untersuchung, überarbeitet Soykas 28.9.11
' Alle leeren Felder und leere Abschnitte weitgehend unterdrückt
Public Sub DoEcho(gcolEcho As Collection, gcolMesswerte As Collection)
    Dim i As Long, itm As Variant
    Dim is_print As Boolean, ausgabe As String
    Dim werte As String, morphologie As String, prothese As String, langtext As String, langtext1 As String, lokalisation As String
    Dim AORKLIND As String, AORKLKLTX As String, AORKLPROB As String, AORTAAIGR As String, _
        AORTAAOEF As String, AORTAASC As String, _
        AORTAKLMO As String, AORTAKLOB As String, AORTAPMAX As String, _
        AORTAPMIT As String, AORTAWUOB As String, AORTAWUR As String, _
        AORTAWUTX As String, datum As String, DIASTLVTE As String, _
        ECHOBEUTX As String, echoBef As String, fragest As String, FS As String, HFFELD As String, _
        IVSDS As String, IVSDS2 As String, LA As String, LVEDD As String, LVEF As String, _
        LVESD As String, LVHW As String, LVHWDS2 As String, LVLW As String, LVLWDS2 As String, _
        LVVW As String, LVVWDS2 As String, MITRALPHT As String, MITRALPMI As String, _
        MITRAMIGR As String, MITRAMOE2 As String, MITRAMOEF As String, MITRINDEX As String, _
        MITRKLOTX As String, MITRKLPOB As String, MITRMOEF As String, MITRMOEF2 As String, _
        MITRMOROB As String, MITRMORTX As String, PERICLOKA As String, PERICLOKS As String, _
        PERICLOKW As String, PERICPEJN As String, PULMKLOB As String, PULMKLTXT As String, _
        PULMLPIGR As String, PWDIEM As String, PWDISM As String, PWDITE As String, _
        PWDITEM As String, PWDITSM As String, REGKINOB As String, REGKINTXT As String, _
        RVEDD As String, RVEF As String, RVESD As String, TRICSYSMX As String, TRICUPIGR As String, _
        TRICUSMOB As String, TRICUSMTX As String, TTETEERAD As String, lvldwds2 As String, _
        dbl As Double, UNAUFTRAG As String, is_empty As Boolean
    
    getDocVar gcolEcho, "ZUNAUFTRAG", UNAUFTRAG, is_empty, False
    
    If getDocVar(gcolEcho, "ZAORKLIND", AORKLIND, is_empty, True) Then
        AORKLIND = "AÖF/KO " & Format(AORKLIND, "#0.00") & " qcm/qm"
    End If
    getDocVar gcolEcho, "ZAORKLPROB", AORKLPROB, is_empty, False
    getDocVar gcolEcho, "ZAORTAAIGR", AORTAAIGR, is_empty, False
    If getDocVar(gcolEcho, "ZAORTAAOEF", AORTAAOEF, is_empty, True) Then
        AORTAAOEF = "AÖF " & Format(AORTAAOEF, "#0.00") & " qcm"
    End If
    If getDocVar(gcolEcho, "ZAORTAASC", AORTAASC, is_empty, True) Then
        AORTAASC = Format(AORTAASC) & " mm"
    End If
    getDocVar gcolEcho, "ZAORTAKLOB", AORTAKLOB, is_empty, False
    If getDocVar(gcolEcho, "ZAORTAPMAX", AORTAPMAX, is_empty, True) Then
        AORTAPMAX = "Delta p max " & Format(AORTAPMAX, "#0.0") & " mmHg"
    End If
    If getDocVar(gcolEcho, "ZAORTAPMIT", AORTAPMIT, is_empty, True) Then
        AORTAPMIT = "Delta p mittel " & Format(AORTAPMIT, "#0.0") & " mmHg"
    End If
    getDocVar gcolEcho, "ZAORTAWUOB", AORTAWUOB, is_empty, True
    If getDocVar(gcolEcho, "ZAORTAWUR", AORTAWUR, is_empty, True) Then
        AORTAWUR = Format(AORTAWUR, "##0") & " mm"
    End If
    getDocVar gcolEcho, "ZDATUM", datum, is_empty, False
    If getDocVar(gcolEcho, "ZDIASTLVTE", DIASTLVTE, is_empty, True) Then
        DIASTLVTE = "diast. LVfunktion. TE " & Format(DIASTLVTE, "##0 msec")
    End If
    getDocVar gcolEcho, "ZFRAGEST", fragest, is_empty, False
    If getDocVar(gcolEcho, "ZFS", FS, is_empty, True) Then
        FS = "FS " & Format(FS, "##0 ") & "%"
    End If
    getDocVar gcolEcho, "ZHFFELD", HFFELD, is_empty, False
    If getDocVar(gcolEcho, "ZIVSDS", IVSDS, is_empty, False) Then
        IVSDS = Format(IVSDS, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZIVSDS2", IVSDS2, is_empty, False) Then
        IVSDS2 = Format(IVSDS2, "##0 mm ")
    End If
    If getDocVar(gcolEcho, "ZLA", LA, is_empty, False) Then
        LA = "LA " & Format(LA, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZLVEDD", LVEDD, is_empty, False) Then
        LVEDD = "LV-EDD " & Format(LVEDD, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZLVEF", LVEF, is_empty, False) Then
        LVEF = "LV-EF " & Format(LVEF, "##0 ") & "%"
    End If
    If getDocVar(gcolEcho, "ZLVESD", LVESD, is_empty, False) Then
        LVESD = "LV-ESD " & Format(LVESD, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZLVHW", LVHW, is_empty, False) Then
        LVHW = Format(LVHW, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZLVHWDS2", LVHWDS2, is_empty, False) Then
        LVHWDS2 = Format(LVHW, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZLVLW", LVLW, is_empty, False) Then
        LVLW = Format(LVLW, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZLVLWDS2", LVLWDS2, is_empty, False) Then
        LVLWDS2 = Format(LVLW, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZLVVW", LVVW, is_empty, False) Then
        LVVW = Format(LVVW, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZLVVWDS2", LVVWDS2, is_empty, False) Then
        LVVWDS2 = Format(LVVWDS2, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZMITRALPHT", MITRALPHT, is_empty, True) Then
        MITRALPHT = "PHT " & Format(MITRALPHT, "#0.0") & " msec"
    End If
    If getDocVar(gcolEcho, "ZMITRALPMI", MITRALPMI, is_empty, True) Then
        MITRALPMI = "Delta p mittel " & Format(MITRALPMI, "#0.0") & " mmHg"
    End If
    If getDocVar(gcolEcho, "ZMITRAMIGR", MITRAMIGR, is_empty, False) Then
        MITRAMIGR = "MI " & MITRAMIGR
    End If
    If getDocVar(gcolEcho, "ZMITRAMOE2", MITRAMOE2, is_empty, True) Then
        MITRAMOE2 = "MÖF(2D) " & Format(MITRAMOE2, "#0.0") & " qcm"
    End If
    getDocVar gcolEcho, "ZMITRAMOEF", MITRAMOEF, is_empty, False
    getDocVar gcolEcho, "ZMITRMOEF2", MITRMOEF2, is_empty, False
    If getDocVar(gcolEcho, "ZMITRINDEX", MITRINDEX, is_empty, True) Then
        MITRINDEX = "MÖF/KO " & Format(MITRINDEX, "#0.00") & " qcm/qm"
    End If
    getDocVar gcolEcho, "ZMITRKLPOB", MITRKLPOB, is_empty, False
    If getDocVar(gcolEcho, "ZMITRMOEF", MITRMOEF, is_empty, True) Then
        MITRMOEF = "MÖF " & Format(MITRMOEF, "#0.0") & " qcm"
    End If
    getDocVar gcolEcho, "ZMITRMOROB", MITRMOROB, is_empty, False
    If getDocVar(gcolEcho, "ZPERICLOKA", PERICLOKA, is_empty, False) Then
        PERICLOKA = LCase(PERICLOKA)
    End If
    If getDocVar(gcolEcho, "ZPERICLOKW", PERICLOKW, is_empty, True) Then
        PERICLOKW = Format(PERICLOKW, "#00") & " mm"
    End If
    getDocVar gcolEcho, "ZPERICPEJN", PERICPEJN, is_empty, False
    getDocVar gcolEcho, "ZPULMKLOB", PULMKLOB, is_empty, False
    If getDocVar(gcolEcho, "ZPULMLPIGR", PULMLPIGR, is_empty, False) Then
        PULMLPIGR = "PI " & PULMLPIGR
    End If
    If getDocVar(gcolEcho, "ZPWDIEM", PWDIEM, is_empty, False) Then
        PWDIEM = "Em " & Format(PWDIEM, "#0.0") & " cm/sec"
    End If
    If getDocVar(gcolEcho, "ZPWDISM", PWDISM, is_empty, True) Then
        PWDISM = "Sm " & Format(PWDISM, "#0.0") & " cm/sec"
    End If
    If getDocVar(gcolEcho, "ZPWDITE", PWDITE, is_empty, True) Then
        PWDITE = "TE " & Format(PWDITE, "#0.0") & " msec"
    End If
    If getDocVar(gcolEcho, "ZPWDITEM", PWDITEM, is_empty, True) Then
        PWDITEM = "Tem " & Format(PWDITEM, "#0.0") & " msec"
    End If
    If getDocVar(gcolEcho, "ZPWDITSM", PWDITSM, is_empty, True) Then
        PWDITSM = "Tsm " & Format(PWDITSM, "#0.0") & " msec"
    End If
    getDocVar gcolEcho, "ZREGKINOB", REGKINOB, is_empty, False
    If getDocVar(gcolEcho, "ZRVEDD", RVEDD, is_empty, False) Then
        RVEDD = "RV-EDD " & Format(RVEDD, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZRVEF", RVEF, is_empty, False) Then
        RVEF = "RV-EF " & Format(RVEF, "##0") & " %"
    End If
    If getDocVar(gcolEcho, "ZRVESD", RVESD, is_empty, False) Then
        RVESD = "RV-ESD " & Format(RVEF, "##0 mm")
    End If
    If getDocVar(gcolEcho, "ZTRICSYSMX", TRICSYSMX, is_empty, True) Then
        TRICSYSMX = "sys Delta p max " & Format(TRICSYSMX, "#00") & " mmHg"
    End If
    If getDocVar(gcolEcho, "ZTRICUPIGR", TRICUPIGR, is_empty, False) Then
        TRICUPIGR = "TI " & TRICUPIGR
    End If
    getDocVar gcolEcho, "ZTRICUSMOB", TRICUSMOB, is_empty, False
    getDocVar gcolEcho, "ZTTETEERAD", TTETEERAD, is_empty, True
    getDocVar gcolEcho, "ZREGKINTXT<L>", REGKINTXT, is_empty, False
    If getDocVar(gcolEcho, "ZAORTAWUTX<L>", AORTAWUTX, is_empty, False) Then
        AORTAWUTX = "Morphologie der A. asc./desc.: " & AORTAWUTX & vbCrLf
    End If
    getDocVar gcolEcho, "ZAORTAKLMO<L>", AORTAKLMO, is_empty, False
    getDocVar gcolEcho, "ZAORKLKLTX<L>", AORKLKLTX, is_empty, False
    getDocVar gcolEcho, "ZMITRMORTX<L>", MITRMORTX, is_empty, False
    getDocVar gcolEcho, "ZMITRKLOTX<L>", MITRKLOTX, is_empty, False
    getDocVar gcolEcho, "ZPULMKLTXT<L>", PULMKLTXT, is_empty, False
    getDocVar gcolEcho, "ZTRICUSMTX<L>", TRICUSMTX, is_empty, False
    getDocVar gcolEcho, "ZPERICLOKS<L>", PERICLOKS, is_empty, False
    ' Im Dokument zwei unterschiedliche Fehler in zwei verschiedenen
    ' Dokumentversionen, wird aber gleich an der gleichen Stelle ausgegeben
    If Not getDocVar(gcolEcho, "ZECHOBEUTX<L>", ECHOBEUTX, is_empty, False) Then
        getDocVar gcolEcho, "ZEBEURTEIL<L>", ECHOBEUTX, is_empty, False
    End If
    getDocVar gcolEcho, "ZEBEFUND<L>", echoBef, is_empty, False
        
    ' Befund vor der Textmarke einfuegen, dann wird der naechste folgende
    ' Befund auch richtig hinter dem ersten Befund ausgegeben -
    ' vorausgesetzt, die Reihenfolge war vorher richtig
    ActiveDocument.Bookmarks("Echokard").Select
    Selection.MoveLeft , 1
    set_format
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    If UNAUFTRAG > "" Then
        Selection.TypeText UNAUFTRAG & " vom "
    ElseIf IsNumeric(TTETEERAD) Then
        If TTETEERAD = 1 Then
            Selection.TypeText "TTE vom "
        ElseIf TTETEERAD = 2 Then
            Selection.TypeText "TEE vom "
        End If
    End If
    Selection.TypeText text:=ISHMED_Datum(datum & "") & ":"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    If fragest <> "" Then
        Selection.Font.Italic = True
        Selection.TypeText "Fragestellung: "
        Selection.Font.Italic = False
        Selection.TypeParagraph
        Selection.TypeText fragest
        Selection.TypeParagraph
    End If
        
    is_print = False
    werte = ""
    Selection.Paragraphs.TabStops.ClearAll
    'Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft

    If LVEDD > "" Then
        werte = LVEDD
    End If
    If LVESD > "" Then
        anhaengen werte, LVESD
    End If
    If FS > "" Then
        anhaengen werte, FS
    End If
    If LA > "" Then
        anhaengen werte, LA
    End If
    If LVVW > "" Or LVVWDS2 > "" Then
        If LVVW = "" Then
            anhaengen werte, "LV-VW Syst. " & LVVWDS2
        ElseIf LVVWDS2 = "" Then
            anhaengen werte, "LV-VW Diast. " & LVVW
        Else
            anhaengen werte, "LV-VW D/S " & LVVW & " / " & LVVWDS2
        End If
    End If
    If LVEF > "" Then
        anhaengen werte, LVEF
    End If
    If werte <> "" Then
      Selection.TypeText werte
      Selection.TypeParagraph
      werte = ""
    End If
    
    lvldwds2 = LVLWDS2
    If LVLW > "" Or lvldwds2 > "" Then
        If LVLW = "" Then
            werte = "LV-LW Syst. " & lvldwds2
        ElseIf lvldwds2 = "" Then
            werte = "LV-LW Diast. " & LVLW
        Else
            werte = "LV-LW D/S " & LVLW & " / " & lvldwds2
        End If
    End If
        
    If DIASTLVTE > "" Then
        anhaengen werte, DIASTLVTE
    End If
    If LVHW > "" Or LVHWDS2 > "" Then
        If LVHW = "" Then
           anhaengen werte, "LV-HW Syst. " & LVHWDS2
        ElseIf LVHWDS2 = "" Then
           anhaengen werte, "LV-HW Diast. " & LVHW
        Else
           anhaengen werte, "LV-HW D/S " & LVHW & " / " & LVHWDS2
        End If
    End If
    If IVSDS > "" And IVSDS2 > "" Then
        anhaengen werte, "IVS D/S " & IVSDS & " / " & IVSDS2
    ElseIf IVSDS > "" Then
        anhaengen werte, "IVS Diast. " & IVSDS
    ElseIf IVSDS2 > "" Then
        anhaengen werte, "IVS Syst. " & IVSDS2
    End If
    If RVEDD > "" Then
        anhaengen werte, RVEDD
    End If
    If RVESD > "" Then
        anhaengen werte, RVESD
    End If
    If RVEF > "" Then
        anhaengen werte, RVEF
    End If
    If werte <> "" Then
      Selection.TypeText werte
      Selection.TypeParagraph
      werte = ""
    End If
    
    langtext = REGKINTXT
    werte = ""
    If REGKINOB = "X" Then
        werte = "keine"
    End If
    If langtext > "" Then
        anhaengen werte, langtext
    End If
    If werte > "" Then
        werte = "Regionale Kinetikstörungen: " & werte
        Selection.TypeText werte
        Selection.TypeParagraph
    End If
    
    ' Aortenwurzel
    werte = ""
    If AORTAWUOB = "X" Then
        werte = "o.B."
    End If
    If AORTAWUR > "" Then
        anhaengen werte, AORTAWUR
    End If
    If werte > "" Then
        Selection.TypeText "Aortenwurzel " & werte & vbCrLf
        werte = ""
    End If
    ' Ende Abschnitt Aortenwurzel
    
    ' Abschnitt A. asc./desc.
    If AORTAWUTX > "" Then
        Selection.TypeText AORTAWUTX
    End If
    If AORTAASC > "" Then
        Selection.TypeText "Aorta ascendens " & AORTAASC & vbCrLf
    End If
    ' Ende Abschnitt A. asc./desc.
    
    ' Abschnitt Aortenklappe
    werte = ""
    langtext = AORTAKLMO
    langtext1 = AORKLKLTX
    If AORTAPMIT > "" Then
      werte = AORTAPMIT
    End If
    If AORTAAOEF > "" Then
        anhaengen werte, AORTAAOEF
    End If
    If AORTAPMAX > "" Then
        anhaengen werte, AORTAPMAX
    End If
    If AORKLIND > "" Then
        anhaengen werte, AORKLIND
    End If
    If AORTAAIGR > "" Then
        anhaengen werte, "AI " & AORTAAIGR
    End If
    
    morphologie = ""
    If AORTAKLOB = "X" Then
       morphologie = "o.B."
    End If
    If langtext > "" Then
      anhaengen morphologie, langtext
    End If
    If morphologie > "" Then
        morphologie = "Morphologie " & morphologie & vbCrLf
    End If
    
    prothese = ""
    If AORKLPROB = "X" Then
       prothese = prothese & "o.B."
    End If
    If langtext1 > "" Then
        anhaengen prothese, langtext1
    End If
    If prothese > "" Then
        prothese = "Klappenprothese: " & prothese & vbCrLf
    End If

    If morphologie > "" Or prothese > "" Or werte > "" Then
        Selection.TypeText "Aortenklappe: " & _
            morphologie & prothese & werte & vbCrLf
    End If
    ' Ende Abschnitt Aortenklappe

    ' Abschnitt Mitralklappe
    langtext = MITRMORTX
    langtext1 = MITRKLOTX
    werte = ""
    If MITRALPMI > "" Then
        werte = MITRALPMI
    End If
    If MITRALPHT > "" Then
        anhaengen werte, MITRALPHT
    End If
    If MITRMOEF > "" Then
        anhaengen werte, MITRMOEF
    End If
    If MITRAMOE2 > "" Then
        anhaengen werte, MITRAMOE2
    End If
    If MITRAMIGR > "" Then
        anhaengen werte, MITRAMIGR
    End If
    If MITRINDEX > "" Then
        anhaengen werte, MITRINDEX
    End If
    
    morphologie = ""
    If MITRMOROB = "X" Then
       morphologie = morphologie & "o.B."
    End If
    If langtext > "" Then
        anhaengen morphologie, langtext
    End If
    If morphologie > "" Then
        morphologie = "Morphologie " & morphologie & vbCrLf
    End If
    
    prothese = ""
    If MITRKLPOB = "X" Then
       prothese = prothese & "o.B."
    Else
    End If
    If langtext1 > "" Then
        anhaengen prothese, langtext1
    End If
    If prothese > "" Then
        prothese = "Klappenprothese: " & prothese & vbCrLf
    End If
    ausgabe = morphologie & prothese & werte
    
    If ausgabe > "" Then
        If Right(Trim(ausgabe), 2) <> vbCrLf Then ausgabe = ausgabe & vbCrLf
        Selection.TypeText "Mitralklappe: " & ausgabe
    End If
    ' Ende Abschnitt Mitralklappe
    
    ' Abschnitt Pulmonalklappe
    langtext = PULMKLTXT
    
    morphologie = ""
    If PULMKLOB = "X" Then
       morphologie = morphologie & "o.B."
    End If
    If langtext > "" Then
        anhaengen morphologie, langtext
    End If
    If morphologie > "" Then
        morphologie = "Morphologie " & morphologie & vbCrLf
    End If
    
    werte = ""
    If PULMLPIGR > "" Then
        werte = PULMLPIGR
    End If
    ausgabe = morphologie & werte
    If ausgabe > "" Then
        If Right(Trim(ausgabe), 2) <> vbCrLf Then ausgabe = ausgabe & vbCrLf
        Selection.TypeText "Pulmonalklappe: " & ausgabe
    End If
    ' Ende Abschnitt Pulmonalklappe
    
    ' Abschnitt Tricuspidalklappe
    langtext = TRICUSMTX
    werte = ""
    If TRICUPIGR > "" Then
        werte = TRICUPIGR
    End If
    If TRICSYSMX > "" Then
        anhaengen werte, TRICSYSMX
    End If
    morphologie = ""
    If TRICUSMOB = "X" Then
       morphologie = morphologie & "o.B."
    End If
    If langtext > "" Then
        anhaengen morphologie, langtext
    End If
    If morphologie > "" Then
        morphologie = " Morphologie " & morphologie & vbCrLf
    End If
    ausgabe = morphologie & werte
    If ausgabe > "" Then
        If Right(Trim(ausgabe), 2) <> vbCrLf Then ausgabe = ausgabe & vbCrLf
        Selection.TypeText "Trikuspidalklappe: " & ausgabe
    End If
    ' Ende Abschnitt Tricuspidalklappe
    
    ' Abschnitt PWDI
    werte = ""
    If PWDISM > "" Then
        werte = PWDISM
    End If
    If PWDIEM > "" Then
        anhaengen werte, PWDIEM
    End If
    If PWDITSM > "" Then
        anhaengen werte, PWDITSM
    End If
    If PWDITEM > "" Then
        anhaengen werte, PWDITEM
    End If
    If PWDITE > "" Then
        anhaengen werte, PWDITE
    End If
    If werte > "" Then
        Selection.TypeText "PWDI: " & werte & vbCrLf
    End If
    ' Ende Abschnitt PWDI

    ' Abschnitt Perikard
    werte = ""
    If IsNumeric(PERICPEJN) Then
        If PERICPEJN = 1 Then
           werte = "n.best."
        End If
        If PERICPEJN = 2 Then
           werte = "ja"
        End If
        If PERICPEJN = 3 Then
           werte = "nein"
        End If
    End If
    
    lokalisation = ""
    If PERICLOKA > "" Then
        lokalisation = PERICLOKA
    End If
    If PERICLOKW > "" Then
        anhaengen lokalisation, PERICLOKW
    End If
    If lokalisation > "" Then
        lokalisation = "Lokalisation " & lokalisation & vbCrLf
    End If
    
    langtext = PERICLOKS
    If langtext > "" Then
        anhaengen werte, langtext & vbCrLf
    End If
    ausgabe = werte & lokalisation
    If ausgabe > "" Then
        If Right(Trim(ausgabe), 2) <> vbCrLf Then ausgabe = ausgabe & vbCrLf
         Selection.TypeText "Perikarderguss: " & ausgabe
    End If
    ' Ende Abschnitt Perikard
    
    ' Neue Messwertetabelle
    If Not gcolMesswerte Is Nothing Then
        If gcolMesswerte.Count > 0 Then
            Dim line As Collection, gefunden As Boolean, _
                initialized As Boolean, kapitel As String, tabrow As Collection
            Selection.Font.Italic = True
            Selection.TypeText "Messwerte:"
            Selection.Font.Italic = False
            Selection.TypeParagraph
            
        End If
    End If

    ' Abschnitt Befund
    If echoBef > "" Then
        Selection.Font.Italic = True
        Selection.TypeText "Befund:"
        Selection.Font.Italic = False
        Selection.TypeParagraph
        print_langtext echoBef
    End If
    
    ' Abschnitt Beurteilung
    If ECHOBEUTX > "" Then
        Selection.Font.Italic = True
        Selection.TypeText "Beurteilung:"
        Selection.Font.Italic = False
        Selection.TypeParagraph
        print_langtext ECHOBEUTX
    End If
   
End Sub

Public Sub DoRechtsherz(gcolRechtsherz As Collection)
    Dim i As Long, itm As Variant
    Dim is_print As Boolean
    Dim zweite_tabelle As Boolean
    Dim werte As String
    Dim datum As String, zeit, UNTERSUCH, RECHTSRU, RECHTRR_1, RECHTRR_2, RECHTRR_3, RECHTRR_T, RECHTRA_1, _
        RECHTRA_T, RECHTRV_1, RECHTRV_2, RECHTRV_3, RECHTRV_T, RECHTPA_1, RECHTPA_2, RECHTPA_3, _
        RECHTPA_T, RECHTPCM, RECHTPCMT, RECHTHZV, RECHTHZVT, RECHTCI, RECHTCI_T, RECHTPVR, _
        RECHTPVRT, RECHTSVO2, RECHTSVOT, RECHTSSVR, RECHTSVRT, RECHTTPG, RECHTTPGT, RECHTSTAB, _
        RECHTRETX, RECHTRET1, O2GABE, O2GABE1, O2GABE2, O2GABE3, O2GABE4, O2GABE5, O2GABE6, _
        NITRO, NITRO1, NITRO2, NITRO3, NITRO4, NITRO5, NITRO6, RECHTRE1, RTABDOSIS, RTABDOSI1, _
        RTABDOSI2, RTABDOSI3, RTABDOSI4, RTABDOSI5, RTABDOSI6, RECHTRE2, RTABBELR, RTABBEL2, _
        RTABBEL3, RTABBEL4, RTABBEL5, RTABBEL6, RTABBEL7, RTABRR_1, RTABRR_2, RTABRR_3, RTABRR_4, _
        RTABRR_5, RTABRR_6, RTABRR_7, RTABRA_1, RTABRA_2, RTABRA_3, RTABRA_4, RTABRA_5, RTABRA_6, _
        RTABRA_7, RTABRV_1, RTABRV_2, RTABRV_3, RTABRV_4, RTABRV_5, RTABRV_6, RTABRV_7, RTABPA_1, _
        RTABPA_2, RTABPA_3, RTABPA_4, RTABPA_5, RTABPA_6, RTABPA_7, RTABPC_1, RTABPC_2, RTABPC_3, _
        RTABPC_4, RTABPC_5, RTABPC_6, RTABPC_7, RTABHZV1, RTABHZV2, _
        RTABHZV3, RTABHZV4, RTABHZV5, RTABHZV6, RTABHZV7, RTABCI_1, RTABCI_2, RTABCI_3, RTABCI_4, _
        RTABCI_5, RTABCI_6, RTABCI_7, RTABSVO_1, RTABSVO_2, RTABSVO_3, RTABSVO_4, RTABSVO_5, _
        RTABSVO_6, RTABSVO_7, RTABPVR_1, RTABPVR_2, RTABPVR_3, RTABPVR_4, RTABPVR_5, RTABPVR_6, _
        RTABPVR_7, RTABSVR_1, RTABSVR_2, RTABSVR_3, RTABSVR_4, RTABSVR_5, RTABSVR_6, RTABSVR_7, _
        RTABTPG_1, RTABTPG_2, RTABTPG_3, RTABTPG_4, RTABTPG_5, RTABTPG_6, RTABTPG_7, TXTUNTERS, _
        TBEURTEIL, RECHTUNT, HFFELD, HFTXT
        
    zweite_tabelle = False
    Dim is_empty As Boolean
    getDocVar gcolRechtsherz, "ZHFTXT", HFTXT, is_empty, True
    getDocVar gcolRechtsherz, "ZDATUM", datum, is_empty, True
    getDocVar gcolRechtsherz, "ZHFFELD", HFFELD, is_empty, True
    getDocVar gcolRechtsherz, "ZNITRO", NITRO, is_empty, True
    getDocVar gcolRechtsherz, "ZNITRO1", NITRO1, is_empty, True
    getDocVar gcolRechtsherz, "ZNITRO2", NITRO2, is_empty, True
    getDocVar gcolRechtsherz, "ZNITRO3", NITRO3, is_empty, True
    getDocVar gcolRechtsherz, "ZNITRO4", NITRO4, is_empty, True
    getDocVar gcolRechtsherz, "ZNITRO5", NITRO5, is_empty, True
    getDocVar gcolRechtsherz, "ZNITRO6", NITRO6, is_empty, True
    getDocVar gcolRechtsherz, "ZO2GABE", O2GABE, is_empty, True
    getDocVar gcolRechtsherz, "ZO2GABE1", O2GABE1, is_empty, True
    getDocVar gcolRechtsherz, "ZO2GABE2", O2GABE2, is_empty, True
    getDocVar gcolRechtsherz, "ZO2GABE3", O2GABE3, is_empty, True
    getDocVar gcolRechtsherz, "ZO2GABE4", O2GABE4, is_empty, True
    getDocVar gcolRechtsherz, "ZO2GABE5", O2GABE5, is_empty, True
    getDocVar gcolRechtsherz, "ZO2GABE6", O2GABE6, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTCI", RECHTCI, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTCI_T", RECHTCI_T, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTHZV", RECHTHZV, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTHZVT", RECHTHZVT, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTPA_1", RECHTPA_1, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTPA_2", RECHTPA_2, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTPA_3", RECHTPA_3, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTPA_T", RECHTPA_T, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTPCM", RECHTPCM, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTPCMT", RECHTPCMT, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTPVR", RECHTPVR, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTPVRT", RECHTPVRT, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRA_1", RECHTRA_1, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRA_T", RECHTRA_T, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRE1", RECHTRE1, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRE2", RECHTRE2, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRET1", RECHTRET1, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRETX", RECHTRETX, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRR_1", RECHTRR_1, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRR_2", RECHTRR_2, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRR_3", RECHTRR_3, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRR_T", RECHTRR_T, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRV_1", RECHTRV_1, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRV_2", RECHTRV_2, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRV_3", RECHTRV_3, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTRV_T", RECHTRV_T, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTSRU", RECHTSRU, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTSSVR", RECHTSSVR, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTSTAB", RECHTSTAB, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTSVO2", RECHTSVO2, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTSVOT", RECHTSVOT, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTSVRT", RECHTSVRT, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTTPG", RECHTTPG, is_empty, True
    getDocVar gcolRechtsherz, "ZRECHTTPGT", RECHTTPGT, is_empty, True
    If getDocVar(gcolRechtsherz, "ZRTABPA_1", RTABPA_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPA_2", RTABPA_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPA_3", RTABPA_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPA_4", RTABPA_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPA_5", RTABPA_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPA_6", RTABPA_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPA_7", RTABPA_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPC_1", RTABPC_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPC_2", RTABPC_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPC_3", RTABPC_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPC_4", RTABPC_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPC_5", RTABPC_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPC_6", RTABPC_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPC_7", RTABPC_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABHZV1", RTABHZV1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABHZV2", RTABHZV2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABHZV3", RTABHZV3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABHZV4", RTABHZV4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABHZV5", RTABHZV5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABHZV6", RTABHZV6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABHZV7", RTABHZV7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABCI_1", RTABCI_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABCI_2", RTABCI_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABCI_3", RTABCI_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABCI_4", RTABCI_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABCI_5", RTABCI_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABCI_6", RTABCI_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABCI_7", RTABCI_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPVR_1", RTABPVR_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPVR_2", RTABPVR_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPVR_3", RTABPVR_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPVR_4", RTABPVR_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPVR_5", RTABPVR_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPVR_6", RTABPVR_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABPVR_7", RTABPVR_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRA_1", RTABRA_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRA_2", RTABRA_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRA_3", RTABRA_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRA_4", RTABRA_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRA_5", RTABRA_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRA_6", RTABRA_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRA_7", RTABRA_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRR_1", RTABRR_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRR_2", RTABRR_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRR_3", RTABRR_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRR_4", RTABRR_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRR_5", RTABRR_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRR_6", RTABRR_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRR_7", RTABRR_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRV_1", RTABRV_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRV_2", RTABRV_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRV_3", RTABRV_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRV_4", RTABRV_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRV_5", RTABRV_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRV_6", RTABRV_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABRV_7", RTABRV_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVO_1", RTABSVO_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVO_2", RTABSVO_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVO_3", RTABSVO_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVO_4", RTABSVO_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVO_5", RTABSVO_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVO_6", RTABSVO_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVO_7", RTABSVO_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVR_1", RTABSVR_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVR_2", RTABSVR_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVR_3", RTABSVR_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVR_4", RTABSVR_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVR_5", RTABSVR_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVR_6", RTABSVR_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABSVR_7", RTABSVR_7, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABTPG_1", RTABTPG_1, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABTPG_2", RTABTPG_2, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABTPG_3", RTABTPG_3, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABTPG_4", RTABTPG_4, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABTPG_5", RTABTPG_5, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABTPG_6", RTABTPG_6, is_empty, True) Then zweite_tabelle = Not is_empty
    If getDocVar(gcolRechtsherz, "ZRTABTPG_7", RTABTPG_7, is_empty, True) Then zweite_tabelle = Not is_empty
    getDocVar gcolRechtsherz, "ZUNTERSUCH", UNTERSUCH, is_empty, True
    getDocVar gcolRechtsherz, "ZZEIT", zeit, is_empty, True
    getDocVar gcolRechtsherz, "ZTXTUNTERS<L>", TXTUNTERS, is_empty, False
    getDocVar gcolRechtsherz, "ZTBEURTEIL<L>", TBEURTEIL, is_empty, False
    getDocVar gcolRechtsherz, "ZRECHTUNT", RECHTUNT, is_empty, False
    ActiveDocument.Bookmarks("Rechtsherz").Select
    set_format
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="Rechtsherzkatheter vom "
    Selection.TypeText text:=ISHMED_Datum(datum) & ":"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    
    If TXTUNTERS > "" Then
        Selection.TypeParagraph
        Selection.TypeText "Ablauf der Untersuchung: "
        print_langtext TXTUNTERS
    End If
        
    is_print = False
    Selection.Paragraphs.TabStops.ClearAll
    Selection.Paragraphs.TabStops.Add InchesToPoints(0.78)
    
    'Eintragen der Werte für RR art.:
    werte = ""
    If (RECHTRR_1 <> 0) Then
       werte = "RR art " & Chr(9) & Format(RECHTRR_1, "#,##0")
    End If
    If (RECHTRR_2 <> 0) Then
       werte = werte & " / " & Format(RECHTRR_2, "#,##0")
    End If
    If (RECHTRR_3 <> 0) Then
        werte = werte & " / " & Format(RECHTRR_3, "#,##0")
    End If
    If werte <> "" Then
        werte = werte & " mmHG"
        Selection.TypeText werte
        Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für RA:
    werte = ""
    If (RECHTRA_1 <> 0) Then
       werte = "RA " & Chr(9) & Format(RECHTRA_1, "#,##0")
    End If
    If werte <> "" Then
      werte = werte & " mmHG"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für RV:
    werte = ""
    If (RECHTRV_1 <> 0) Then
       werte = "RV " & Chr(9) & Format(RECHTRV_1, "#,##0")
    End If
    If (RECHTRV_2 <> 0) Then
       werte = werte & " / " & Format(RECHTRV_2, "#,##0")
    End If
    If (RECHTRV_3 <> 0) Then
        werte = werte & " / " & Format(RECHTRV_3, "#,##0")
    End If
    If werte <> "" Then
      werte = werte & " mmHG"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für PA:
    werte = ""
    If (RECHTPA_1 <> 0) Then
       werte = "PA " & Chr(9) & Format(RECHTPA_1, "#,##0")
    End If
    If (RECHTPA_2 <> 0) Then
       werte = werte & " / " & Format(RECHTPA_2, "#,##0")
    End If
    If (RECHTPA_3 <> 0) Then
        werte = werte & " / " & Format(RECHTPA_3, "#,##0")
    End If
    If werte <> "" Then
      werte = werte & " mmHG"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für PCm:
    werte = ""
    If (RECHTPCM <> 0) Then
       werte = "PCm " & Chr(9) & Format(RECHTPCM, "#,##0")
    End If
    If werte <> "" Then
      werte = werte & " mmHG"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für HZV:
    werte = getDecimalFormat(RECHTHZV)
    If werte <> "" Then
      werte = "HZV" & Chr(9) & werte & " l/min"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für CI:
    werte = getDecimalFormat(RECHTCI)
    If werte <> "" Then
      werte = "CI" & Chr(9) & werte & " l/min/m²"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für SVO2:
    werte = ""
    If RECHTSVO2 <> 0 Then
      werte = "SVO2" & Chr(9) & getDecimalFormat(RECHTSVO2) & " %"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für PVR:
    werte = ""
    If (RECHTPVR <> 0) Then
      werte = "PVR " & Chr(9) & Format(RECHTPVR, "#,##0") & " dyn sec cm"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
     'Eintragen der Werte für SVR:
    werte = ""
    If (RECHTSSVR <> 0) Then
        werte = "SVR " & Chr(9) & Format(RECHTSSVR, "#,##0") & " dyn sec cm"
        Selection.TypeText werte
        Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für TPG:
    werte = ""
    If (RECHTTPG <> 0) Then
      werte = "TPG " & Chr(9) & Format(RECHTTPG, "#,##0") & " mmHG"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für HF:
    werte = ""
    If (HFFELD <> 0) Then
      Selection.TypeParagraph
      werte = "HF " & Chr(9) & Format(HFFELD, "##0") & "/min"
      Selection.TypeText werte
      Selection.TypeParagraph
    End If
       
    'Eintragen der Werte für HF-Text
    If HFTXT > "" Then
        print_langtext HFTXT
        Selection.TypeParagraph
    End If
    
    If TBEURTEIL > "" Then
        Selection.TypeParagraph
        Selection.TypeText "Beurteilung: "
        print_langtext TBEURTEIL
    End If
    
    If zweite_tabelle Then
        Selection.TypeParagraph
        Selection.TypeText "Reversibilitätsprüfung einer pulmonalen Hypertonie mit Flolan " & _
            "(Epoprostenol) in steigender Dosis oder Hämodynamikprüfung unter Belastung:"
        Selection.TypeParagraph
    
        Dim tbl As Word.Table, row As Word.row, breiten(), ueberschriften()
        breiten = Array(InchesToPoints(1.2), InchesToPoints(0.61), InchesToPoints(0.61), InchesToPoints(0.61), _
                        InchesToPoints(0.61), InchesToPoints(0.61), InchesToPoints(0.61), InchesToPoints(0.61))
        'Eintragen der Spaltenköpfe für die nachfolgenden Tabellenwerte:
        ueberschriften = Array("", "Ruhe", "2", "4", "6", "8", "10", "12")
        anlegenTabelle tbl, breiten, ueberschriften, True
    
        Set row = tbl.Rows(2)
        row.Cells(1).Range.text = "RR mmHg"
        
        'Eintragen der Werte für RR mmHG:
        row.Cells(2).Range.text = Format(RTABRR_1, "##0")
        row.Cells(3).Range.text = Format(RTABRR_2, "##0")
        row.Cells(4).Range.text = Format(RTABRR_3, "##0")
        row.Cells(5).Range.text = Format(RTABRR_4, "##0")
        row.Cells(6).Range.text = Format(RTABRR_5, "##0")
        row.Cells(7).Range.text = Format(RTABRR_6, "##0")
        row.Cells(8).Range.text = Format(RTABRR_7, "##0")
    
    
        'Eintragen der Werte für RA mmHG:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "RAm mmHg"
        row.Cells(2).Range.text = Format(RTABRA_1, "##0")
        row.Cells(3).Range.text = Format(RTABRA_2, "##0")
        row.Cells(4).Range.text = Format(RTABRA_3, "##0")
        row.Cells(5).Range.text = Format(RTABRA_4, "##0")
        row.Cells(6).Range.text = Format(RTABRA_5, "##0")
        row.Cells(7).Range.text = Format(RTABRA_6, "##0")
        row.Cells(8).Range.text = Format(RTABRA_7, "##0")
        
        'Eintragen der Werte für RV mmHG:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "RV mmHg"
        row.Cells(2).Range.text = Format(RTABRV_1, "##0")
        row.Cells(3).Range.text = Format(RTABRV_2, "##0")
        row.Cells(4).Range.text = Format(RTABRV_3, "##0")
        row.Cells(5).Range.text = Format(RTABRV_4, "##0")
        row.Cells(6).Range.text = Format(RTABRV_5, "##0")
        row.Cells(7).Range.text = Format(RTABRV_6, "##0")
        row.Cells(8).Range.text = Format(RTABRV_7, "##0")
    
        'Eintragen der Werte für PA mmHG:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "PA mmHg"
        row.Cells(2).Range.text = Format(RTABPA_1, "##0")
        row.Cells(3).Range.text = Format(RTABPA_2, "##0")
        row.Cells(4).Range.text = Format(RTABPA_3, "##0")
        row.Cells(5).Range.text = Format(RTABPA_4, "##0")
        row.Cells(6).Range.text = Format(RTABPA_5, "##0")
        row.Cells(7).Range.text = Format(RTABPA_6, "##0")
        row.Cells(8).Range.text = Format(RTABPA_7, "##0")
        
        'Eintragen der Werte für PC mmHG:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "PCm mmHg"
        row.Cells(2).Range.text = Format(RTABPC_1, "##0")
        row.Cells(3).Range.text = Format(RTABPC_2, "##0")
        row.Cells(4).Range.text = Format(RTABPC_3, "##0")
        row.Cells(5).Range.text = Format(RTABPC_4, "##0")
        row.Cells(6).Range.text = Format(RTABPC_5, "##0")
        row.Cells(7).Range.text = Format(RTABPC_6, "##0")
        row.Cells(8).Range.text = Format(RTABPC_7, "##0")
        
        'Eintragen der Werte für HZV l/min:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "HZV l/min"
        row.Cells(2).Range.text = getDecimalFormat(RTABHZV1)
        row.Cells(3).Range.text = getDecimalFormat(RTABHZV2)
        row.Cells(4).Range.text = getDecimalFormat(RTABHZV3)
        row.Cells(5).Range.text = getDecimalFormat(RTABHZV4)
        row.Cells(6).Range.text = getDecimalFormat(RTABHZV5)
        row.Cells(7).Range.text = getDecimalFormat(RTABHZV6)
        row.Cells(8).Range.text = getDecimalFormat(RTABHZV7)
    
        'Eintragen der Werte für CI l/min/qcm:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "CI l/min/m²"
        row.Cells(2).Range.text = getDecimalFormat(RTABCI_1)
        row.Cells(3).Range.text = getDecimalFormat(RTABCI_2)
        row.Cells(4).Range.text = getDecimalFormat(RTABCI_3)
        row.Cells(5).Range.text = getDecimalFormat(RTABCI_4)
        row.Cells(6).Range.text = getDecimalFormat(RTABCI_5)
        row.Cells(7).Range.text = getDecimalFormat(RTABCI_6)
        row.Cells(8).Range.text = getDecimalFormat(RTABCI_7)
    
        'Eintragen der Werte für SVO2 %:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "SVO2 %"
        row.Cells(2).Range.text = getDecimalFormat(RTABSVO_1)
        row.Cells(3).Range.text = getDecimalFormat(RTABSVO_2)
        row.Cells(4).Range.text = getDecimalFormat(RTABSVO_3)
        row.Cells(5).Range.text = getDecimalFormat(RTABSVO_4)
        row.Cells(6).Range.text = getDecimalFormat(RTABSVO_5)
        row.Cells(7).Range.text = getDecimalFormat(RTABSVO_6)
        row.Cells(8).Range.text = getDecimalFormat(RTABSVO_7)
    
        'Eintragen der Werte für PVR dynseccm:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "PVR dynseccm"
        row.Cells(2).Range.text = Format(RTABPVR_1, "##0")
        row.Cells(3).Range.text = Format(RTABPVR_2, "##0")
        row.Cells(4).Range.text = Format(RTABPVR_3, "##0")
        row.Cells(5).Range.text = Format(RTABPVR_4, "##0")
        row.Cells(6).Range.text = Format(RTABPVR_5, "##0")
        row.Cells(7).Range.text = Format(RTABPVR_6, "##0")
        row.Cells(8).Range.text = Format(RTABPVR_7, "##0")
    
        'Eintragen der Werte für SVR dynseccm:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "SVR dynseccm"
        row.Cells(2).Range.text = Format(RTABSVR_1, "##0")
        row.Cells(3).Range.text = Format(RTABSVR_2, "##0")
        row.Cells(4).Range.text = Format(RTABSVR_3, "##0")
        row.Cells(5).Range.text = Format(RTABSVR_4, "##0")
        row.Cells(6).Range.text = Format(RTABSVR_5, "##0")
        row.Cells(7).Range.text = Format(RTABSVR_6, "##0")
        row.Cells(8).Range.text = Format(RTABSVR_7, "##0")
    
        'Eintragen der Werte für TPG mmHG:
        Set row = tbl.Rows.Add
        row.Cells(1).Range.text = "TPG mmHg"
        row.Cells(2).Range.text = Format(RTABTPG_1, "##0")
        row.Cells(3).Range.text = Format(RTABTPG_2, "##0")
        row.Cells(4).Range.text = Format(RTABTPG_3, "##0")
        row.Cells(5).Range.text = Format(RTABTPG_4, "##0")
        row.Cells(6).Range.text = Format(RTABTPG_5, "##0")
        row.Cells(7).Range.text = Format(RTABTPG_6, "##0")
        row.Cells(8).Range.text = Format(RTABTPG_7, "##0")
         ' aus der Tabelle rausgehen
    End If

End Sub ' DoRechtsherz

Sub set_format()
    Selection.Find.ClearFormatting
    Selection.Paragraphs.TabStops.ClearAll
    With Selection.Find
        .text = ""
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
'    Selection.MoveUp Unit:=wdLine, count:=1
'    Selection.EndKey Unit:=wdLine
End Sub

Public Sub DoLufu(gcolLufu As Collection)
    Dim i As Long, itm As Variant
    Dim werte, werte1, mybookmark, mysplitbookmark As String
    Dim tabelle As Variant

    tabelle = Array(Array("", "Ist", "% Soll", "Ist", "% Soll"), _
            Array("VC", "ZVCISTL", "ZVCSOLL", "ZVCIST2L", "ZVCSOLL2"), _
            Array("FVC", Array("ZFVC", "ZFVCL"), "ZFVCSOLL", Array("ZFVCIST2", "ZFVCIST2L"), "ZFVCSOLL2"), _
            Array("FV1", Array("ZFV1IST", "ZFV1ISTL"), "ZFV1SOLL", Array("ZFV1IST2", "ZFV1IST2L"), "ZFV1SOLL2"), _
            Array("FV1%", Array("ZFV1PRIST", "ZFV1PRISTL"), "ZFV1PRSOLL", Array("ZFV1PRIST2", "ZFV1PRIST3"), "ZFV1PRSOL2"), _
            Array("FV1/VC", Array("ZFV1VCIST", "ZFV1VCISTL"), "ZFV1VCSOLL", Array("ZFV1VCIST2", "ZFV1VCIST3"), "ZFV1VCSOL2"), _
            Array("PF", "ZPEFISTL", "ZPEFSOLL", "ZPEFIST2L", "ZPEFSOLL2"), _
            Array("MF 75%", Array("ZMF75IST", "ZMF75ISTL"), "ZMF75SOLL", Array("ZMF75IST2", "ZMF75IST2L"), "ZMF75SOLL2"), _
            Array("MF 50%", Array("ZMF50IST", "ZMF50ISTL"), "ZMF50SOLL", Array("ZMF50IST2", "ZMF50IST2L"), "ZMF50SOLL2"), _
            Array("MF 25%", Array("ZMF25IST", "ZMF25ISTL"), "ZMF25SOLL", Array("ZMF25IST2", "ZMF25IST2L"), "ZMF25SOLL2"), _
            Array("MMF", "ZMMFIST3L", "ZMMFSOLL3", "ZMMFIST23L", "ZMMFSOLL23"), _
            Array("MVV", "ZMVVIST3L", "ZMVVSOLL3", "ZMVVIST23L", "ZMVVSOLL23"))

    Dim BE, datum, fragest, zeit
    
    Dim is_empty As Boolean
    'Stop
   
    getDocVar gcolLufu, "ZBE", BE, is_empty, True
    If Not is_empty Then BE = "BE: " & Format(BE, "#####0.0 mmol/l")
    getDocVar gcolLufu, "ZDATUM", datum, is_empty, False
    If Not is_empty Then datum = ISHMED_Datum(datum & "")
    getDocVar gcolLufu, "ZFRAGEST", fragest, is_empty, False
    getDocVar gcolLufu, "ZZEIT", zeit, is_empty, True

    ActiveDocument.Bookmarks("Lufu").Select
    set_format
    Selection.MoveLeft , 1
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="Lungenfunktion vom "
    Selection.TypeText text:=datum & ":"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
       
    Selection.Paragraphs.TabStops.ClearAll
    Selection.Paragraphs.TabStops.Add InchesToPoints(1)
    Selection.Paragraphs.TabStops.Add InchesToPoints(2.2)
    Selection.Paragraphs.TabStops.Add InchesToPoints(3.4)
    Selection.Paragraphs.TabStops.Add InchesToPoints(4.6)
    Selection.Paragraphs.TabStops.Add InchesToPoints(5.8)
    Selection.Paragraphs.TabStops.Add InchesToPoints(7)
    Selection.Paragraphs.TabStops.Add InchesToPoints(8.2)
    
    ' Generische Ausgabe der Tabelle, in der ersten Zeile stehen die Überschriften
    ' in der ersten Spalte stehen die Bezeichnungen
    ' in den anderen Feldern stehen die Variablennamen. Wenn es für ein Feld mehrere
    ' Variablennamen gibt (weil der Datentyp der Variablen sich im Laufe der
    ' Dokumentversionen geändert hat) steht dort statt des Strings ein Array mit den
    ' alternativen Variablennamen.
    
    Dim zeile As Integer, spalte As Integer, textzeile As String
    Dim zahlenwert, tabtext As String, mit_medi As Boolean
    Dim spalte_eins As Integer, zeile_eins As Integer, letzte_zeile As Integer, letzte_spalte As Integer

    Dim O2, PCO2, PH, PO2, SO2, STB, TBEURTEIL, TXTUNTERS
    
    tabtext = ""
    mit_medi = False
    zeile_eins = LBound(tabelle)
    letzte_zeile = UBound(tabelle)
    For zeile = LBound(tabelle) + 1 To UBound(tabelle)
        spalte_eins = LBound(tabelle(zeile))
        letzte_spalte = UBound(tabelle(zeile))
        textzeile = ""
        For spalte = spalte_eins + 1 To letzte_spalte
            If IsArray(tabelle(zeile)(spalte)) Then
                getAnyDocVar gcolLufu, tabelle(zeile)(spalte), zahlenwert, is_empty, True
            Else
                Dim w As String
                w = tabelle(zeile)(spalte)
                getDocVar gcolLufu, w, zahlenwert, is_empty, True
            End If
            If Not is_empty Then
                textzeile = textzeile & Chr(9) & zahlenwert
                If spalte > spalte_eins + 2 Then
                    ' Wenn ein Wert im hinteren Teil der Tabelle steht,
                    ' dann spaeter auch dafuer die Ueberschriften ausgeben
                    mit_medi = True
                End If
            End If
        Next spalte
        If textzeile > "" Then
            ' Wenn Werte in der Zeile stehen, Bezeichnung aus der ersten Spalte
            ' hinfuegen und ausgeben, sonst keine Spalte ausgeben
            textzeile = tabelle(zeile)(spalte_eins) & textzeile
            tabtext = tabtext & textzeile & vbCrLf
        End If
    Next zeile
    ' Zwei Zeilen Ueberschriften ausgeben, Spalte 1,2,3 immer, 4+5 nur wenn Werte darunter stehen
    spalte_eins = LBound(tabelle(zeile_eins))
    Selection.TypeText Chr(9) & "ohne Bronchospasmolyse"
    If mit_medi Then
        Selection.TypeText Chr(9) & "mit Bronchospasmolyse"
    End If
    Selection.TypeParagraph
    Selection.TypeText tabelle(zeile_eins)(spalte_eins) & Chr(9) & tabelle(zeile_eins)(spalte_eins + 1) & Chr(9) & tabelle(zeile_eins)(spalte_eins + 2)
    If mit_medi Then
        Selection.TypeText Chr(9) & tabelle(zeile_eins)(spalte_eins + 3) & Chr(9) & tabelle(zeile_eins)(spalte_eins + 4)
    End If
    Selection.TypeParagraph
    ' Rest der Tabelle ausgeben
    Selection.TypeText tabtext
    
    werte1 = ""
    If (O2 <> 0) Then
      werte1 = "O2: " & Format(O2, "##0 l/min")
    End If
    If (PO2 <> 0) Then
       werte1 = werte1 + Chr(9) & PO2
    End If
    If PCO2 <> 0 Then
     werte1 = werte1 + Chr(9) & PCO2
    End If
    
    If SO2 <> 0 Then
     werte1 = werte1 + Chr(9) & "SO2: " & Format(SO2, "##0 ") & " %"
    End If
    If werte1 <> "" Then
        Selection.TypeParagraph
        Selection.TypeText "Blutgase:"
        Selection.TypeParagraph
        Selection.TypeText werte1
        Selection.TypeParagraph
    End If
    
    werte = ""
    If (PH > "") Then
         werte = PH
    End If
    If (BE > "") Then
        werte = werte + Chr(9) & BE
    End If
    If STB > "" Then
     werte = werte + Chr(9) & "Stb.: " & STB
    End If
    
    If werte <> "" Then
       If werte1 = "" Then
        Selection.TypeText "Blutgase:"
        Selection.TypeParagraph
       End If
       Selection.TypeText werte
       Selection.TypeParagraph
    End If
    
    DoBefund gcolLufu, "Befund", "Lufu"
    
End Sub

Private Sub print_langtext(was As Variant)
 
 If (RTrim(was) <> "") Then
   Selection.TypeText was
   Selection.TypeParagraph
 End If
 
End Sub

Public Sub DoKoerper(gcolKoerper As Collection)
Dim i As Long, itm As Variant
Dim isprint As Boolean
Dim is_empty As Boolean
Dim groesse, bmi, ko, gewicht As Double
Dim KU_ALGEWI, KU_ALGROE, KU_ALPS_AG, KU_ALPS_DE, KU_ALPS_KL, KU_ALPS_UN, KU_ALSESO, KU_ALSE_AP, _
    KU_ALSE_DE, KU_ALSE_KO, KU_ALSE_UN, KU_ALSE_UR, KU_ALTEMP, KU_AL_AZU, KU_AL_EZU, OBERFLAECH, _
    ZDATUM As String, ZHFFELD, ZTXTUNTERS

    getDocVar gcolKoerper, "KU_ALGEWI", KU_ALGEWI, is_empty, False
    getDocVar gcolKoerper, "KU_ALGROE", KU_ALGROE, is_empty, False
    getDocVar gcolKoerper, "KU_ALPS_AG", KU_ALPS_AG, is_empty, False
    getDocVar gcolKoerper, "KU_ALPS_DE", KU_ALPS_DE, is_empty, False
    getDocVar gcolKoerper, "KU_ALPS_KL", KU_ALPS_KL, is_empty, False
    getDocVar gcolKoerper, "KU_ALPS_UN", KU_ALPS_UN, is_empty, False
    getDocVar gcolKoerper, "KU_ALSESO", KU_ALSESO, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_AP", KU_ALSE_AP, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_DE", KU_ALSE_DE, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_KO", KU_ALSE_KO, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_UN", KU_ALSE_UN, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_UR", KU_ALSE_UR, is_empty, False
    getDocVar gcolKoerper, "KU_ALTEMP", KU_ALTEMP, is_empty, False
    getDocVar gcolKoerper, "KU_AL_AZU", KU_AL_AZU, is_empty, False
    getDocVar gcolKoerper, "KU_AL_EZU", KU_AL_EZU, is_empty, False
    getDocVar gcolKoerper, "ZOBERFLAECH", OBERFLAECH, is_empty, False
    getDocVar gcolKoerper, "ZZDATUM", ZDATUM, is_empty, False
    getDocVar gcolKoerper, "ZZHFFELD", ZHFFELD, is_empty, False
    getDocVar gcolKoerper, "ZZTXTUNTERS", ZTXTUNTERS, is_empty, False
    getDocVar gcolKoerper, "KU_ALGEWI", KU_ALGEWI, is_empty, False
    getDocVar gcolKoerper, "KU_ALGROE", KU_ALGROE, is_empty, False
    getDocVar gcolKoerper, "KU_ALPS_AG", KU_ALPS_AG, is_empty, False
    getDocVar gcolKoerper, "KU_ALPS_DE", KU_ALPS_DE, is_empty, False
    getDocVar gcolKoerper, "KU_ALPS_KL", KU_ALPS_KL, is_empty, False
    getDocVar gcolKoerper, "KU_ALPS_UN", KU_ALPS_UN, is_empty, False
    getDocVar gcolKoerper, "KU_ALSESO", KU_ALSESO, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_AP", KU_ALSE_AP, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_DE", KU_ALSE_DE, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_KO", KU_ALSE_KO, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_UN", KU_ALSE_UN, is_empty, False
    getDocVar gcolKoerper, "KU_ALSE_UR", KU_ALSE_UR, is_empty, False
    getDocVar gcolKoerper, "KU_ALTEMP", KU_ALTEMP, is_empty, False
    getDocVar gcolKoerper, "KU_AL_AZU", KU_AL_AZU, is_empty, False
    getDocVar gcolKoerper, "KU_AL_EZU", KU_AL_EZU, is_empty, False
    getDocVar gcolKoerper, "ZOBERFLAECH", OBERFLAECH, is_empty, False
    getDocVar gcolKoerper, "ZDATUM", ZDATUM, is_empty, False
    getDocVar gcolKoerper, "ZHFFELD", ZHFFELD, is_empty, False
    getAnyDocVar gcolKoerper, Array("ZTXTUNTERS<L>", "ZETXTUNTER<L>"), ZTXTUNTERS, is_empty, False

    ActiveDocument.Bookmarks("koerper").Select
    set_format
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="Körperliche Untersuchung vom "
    Selection.TypeText text:=ISHMED_Datum(ZDATUM) & ":"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    ' Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    
    isprint = False
    
    If (KU_ALGROE > 0) Then
       Selection.TypeText "Größe: " & Format(KU_ALGROE, "##0 ") & "cm  "
       isprint = True
    End If
    If (KU_ALGEWI > 0) Then
       Selection.TypeText "Gewicht: " & Format(KU_ALGEWI, "##0 ") & "kg  "
       isprint = True
    End If
    If (KU_ALGEWI > 0 And KU_ALGROE > 0) Then
        groesse = KU_ALGROE / 100
        groesse = groesse * groesse
        bmi = KU_ALGEWI / groesse
        If bmi < 100 Then
            Selection.TypeText "BMI: " & Format(bmi, "##0.0 kg/m²  ")
        End If
        
        'Selection.TypeText "Temperatur: " & Format(KU_ALTEMP, "#0 ") & " ºC"
        'isprint = True
        
        groesse = KU_ALGROE
        gewicht = KU_ALGEWI
        ko = Exp(0.725 * Log(groesse)) * Exp(0.425 * Log(gewicht)) * 71.84
        ko = ko / 10000
        If ko < 10 Then
            Selection.TypeText "KO: " & Format(ko, "##0.0 m²")
        End If
    End If
    If isprint = True Then
        Selection.TypeParagraph
    End If
    ' Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    print_langtext ZTXTUNTERS
End Sub

Private Sub print_datum(col As Collection)
    Dim itm As Variant
    For Each itm In col
        If itm(0) = "ZDATUM" Then
            Selection.TypeText text:=Right(itm(1), 2) & "." & Mid(itm(1), 5, 2) & "." & Left(itm(1), 4)
            Exit For
        End If
    Next
End Sub

Sub DoBefund(konsil As Collection, Text_ueberschrift As String, Text_marke As String)
    Dim is_empty As Boolean
    Dim datum As String, befund, beurteilung, procedere, anamn, fragest, emedik

    getAnyDocVar konsil, Array("AN_SPEZ<L>", "ZTXTUNTERS<L>", "ZETXTUNTER<L>", "ZBERICHTTX<L>", "ZBEFTEXT<L>", "ZEBEFUND<L>", "ZEBERICHTX<L>"), befund, is_empty, False
    getAnyDocVar konsil, Array("ZEBEURTEIL<L>", "ZTBEURTEIL<L>", "ZEBEURTEIL"), beurteilung, is_empty, False
    getAnyDocVar konsil, Array("ZDATUM", "X00TUM"), datum, is_empty, False
    getDocVar konsil, "ZEANAMNESE<L>", anamn, is_empty, False
    getDocVar konsil, "ZFRAGEST", fragest, is_empty, False
    getDocVar konsil, "ZEMEDIK<L>", emedik, is_empty, False
    getAnyDocVar konsil, Array("ISHMLEIST", "ZUNTERSART"), Text_ueberschrift, is_empty, False
    getDocVar konsil, "ZEPROCEDER<L>", procedere, is_empty, False
    
    Rem If befund = "" And beurteilung = "" Then Exit Sub
    ActiveDocument.Bookmarks(Text_marke).Select
    Selection.MoveLeft , 1
    set_format
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:=Text_ueberschrift
    If Right(RTrim(Text_ueberschrift), 1) <> ":" Then
        If datum > "00000000" Then
            If Right(Text_ueberschrift, 5) <> " vom " Then
                Selection.TypeText " vom "
            End If
            Selection.TypeText text:=ISHMED_Datum(datum)
        End If
    End If
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    If fragest > "" Then
        Selection.Font.Italic = True
        Selection.TypeText text:="Indikation: "
        Selection.Font.Italic = False
        Selection.TypeText text:=fragest
        Selection.TypeParagraph
    End If
    If emedik > "" Then
        Selection.Font.Italic = True
        Selection.TypeText text:="Medikation: "
        Selection.Font.Italic = False
        Selection.TypeText text:=emedik
        Selection.TypeParagraph
    End If
    If befund > "" Then
        Selection.Font.Italic = True
        Selection.TypeText text:="Befund: "
        Selection.Font.Italic = False
        Selection.TypeText text:=befund
        Selection.TypeParagraph
    End If
    If beurteilung > "" Then
        Selection.Font.Italic = True
        Selection.TypeText text:="Beurteilung: "
        Selection.Font.Italic = False
        Selection.TypeText text:=beurteilung
        Selection.TypeParagraph
    End If
    If procedere > "" Then
        Selection.Font.Italic = True
        If Text_ueberschrift <> "Diätempfehlung" Then
            Selection.TypeText text:="Procedere: "
        End If
        Selection.Font.Italic = False
        Selection.TypeText text:=procedere
        Selection.TypeParagraph
    End If
    If anamn > "" Then
        Selection.TypeText text:=anamn
        Selection.TypeParagraph
    End If
End Sub

' Schrittmacher
Sub DoSchrittmacher(konsil As Collection, tb As Collection, werte As Collection)
    Dim tabrow As Collection, liste As Boolean, is_empty As Boolean
    Dim datum As String, befund, beurteilung
    
    liste = False
    If Not tb Is Nothing Then liste = tb.Count > 0
    
    getAnyDocVar konsil, Array("AN_SPEZ<L>", "ZTXTUNTERS<L>", "ZBERICHTTX<L>", "ZBEFTEXT<L>", "ZEBEFUND<L>", "ZHSMBEFUND<L>"), befund, is_empty, False
    getAnyDocVar konsil, Array("ZEBEURTEIL<L>", "ZTBEURTEIL<L>", "ZEBEURTEIL"), beurteilung, is_empty, False
    getDocVar konsil, "ZDATUM", datum, is_empty, False
    
    ' Hier besser klammern wegen Praezedenz der bool'schen Operatoren
    If (befund = "") And (beurteilung = "") And (liste = False) Then Exit Sub
    
    ActiveDocument.Bookmarks("Schrittmacher").Select
    Selection.MoveLeft , 1
    set_format
    
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText "Schrittmacherkontrolle vom " & ISHMED_Datum(datum)
    Selection.Font.Underline = wdUnderlineNone
    
    If befund > "" Then
        Selection.TypeParagraph
        Selection.TypeText text:=befund
    End If
    If beurteilung > "" Then
        Selection.TypeParagraph
        Selection.TypeText text:=beurteilung
    End If
    
    ' todo!! 3.5.2018
    
    Dim tbl As Word.Table
    Dim i, row As Word.row, ZART, ZPRODUKT, ZSERIENNO
    If liste Then
        Selection.TypeParagraph
        anlegenSchrittmacherTabelle tbl
        i = 0
        With tbl
            For Each tabrow In tb
                If i = 0 Then
                    Set row = tbl.Rows(2)
                Else
                    Set row = tbl.Rows.Add
                End If
                i = i + 1
                getTabVar tabrow, "ZART", ZART, is_empty, False
                getTabVar tabrow, "ZPRODUKT", ZPRODUKT, is_empty, False
                getTabVar tabrow, "ZSERIENNO", ZSERIENNO, is_empty, False

                row.Cells(1).Range.text = ZART
                row.Cells(2).Range.text = ZPRODUKT
                row.Cells(3).Range.text = ZSERIENNO
            Next
        End With
    End If
    
    If Not werte Is Nothing Then
        If werte.Count > 0 Then
            Selection.TypeParagraph
            Dim breiten()
            Dim ueberschriften()
            breiten = Array(CentimetersToPoints(10), CentimetersToPoints(5))
            ueberschriften = Array("Bezeichnung", "Wert")
            anlegenTabelle tbl, breiten, ueberschriften, True
            i = 0
            With tbl
                For Each tabrow In werte
                    If i = 0 Then
                        Set row = tbl.Rows(2)
                    Else
                        Set row = tbl.Rows.Add
                    End If
                    i = i + 1
                    getTabVar tabrow, "BEZEICHNER", ZART, is_empty, False
                    getTabVar tabrow, "WERT", ZPRODUKT, is_empty, False
                    row.Cells(1).Range.text = ZART
                    row.Cells(2).Range.text = ZPRODUKT
                Next
            End With
            abschliessenTabelle tbl, False
            Set tbl = Nothing
        End If
    End If
    
End Sub

Private Sub anlegenSchrittmacherTabelle(mytable As Word.Table)
    Selection.TypeParagraph
    Set mytable = Selection.Range.Tables.Add(Selection.Range, 2, 3, wdWord9TableBehavior, wdAutoFitFixed)
    With mytable
        .Rows(1).HeadingFormat = True
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderHorizontal)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
        .Rows(1).Cells.Shading.BackgroundPatternColor = wdColorGray10
        .Rows(1).Cells(1).Range.text = "Aggregat/Sonde"
        .Rows(1).Cells(2).Range.text = "Hersteller/Produkt/Typ"
        .Rows(1).Cells(3).Range.text = "Seriennummer"
    End With
    
    With Options
       .DefaultBorderLineStyle = wdLineStyleSingle
       .DefaultBorderLineWidth = wdLineWidth050pt
       .DefaultBorderColor = wdColorAutomatic
    End With
    
    ' Aggregat/Sonde
    mytable.Columns(1).Select
    'Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(10)
    
    ' Hersteller/Produkt/Typ
    mytable.Columns(2).Select
    'Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(10)
    
    ' Seriennummer
    mytable.Columns(3).Select
    'Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(5)
End Sub


' Selection soll vorher schon an der richtigen Stelle stehen, jetzt wird nur noch eingefuegt
Private Sub anlegenTabelle(mytable As Word.Table, breiten(), ueberschriften(), mit_ueberschrift As Boolean)
    Dim initiale_zeilen As Long, i As Integer, spalten As Integer
    If mit_ueberschrift Then
        initiale_zeilen = 2
    Else
        initiale_zeilen = 1
    End If
    spalten = UBound(breiten) - LBound(breiten) + 1
    ' Neue Zeile, nicht im Fliesstext
    Selection.TypeParagraph
    Selection.Collapse Direction:=wdCollapseEnd
    Set mytable = Selection.Range.Tables.Add(Selection.Range, initiale_zeilen, spalten, wdWord9TableBehavior, wdAutoFitFixed)
    mytable.AllowAutoFit = False
    With mytable
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderHorizontal)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
        Dim spalte As Integer
        spalte = 1
        For i = LBound(breiten) To UBound(breiten)
           .Columns(spalte).SetWidth CSng(breiten(i)), wdAdjustNone
           spalte = spalte + 1
        Next i
        If mit_ueberschrift Then
            .Rows(1).HeadingFormat = True ' Ueberschriftenzeile auf der neuen Seite wiederholen
            .Rows(1).Cells.Shading.BackgroundPatternColor = wdColorGray10
            spalte = 1
            For i = LBound(ueberschriften) To UBound(ueberschriften)
                .Rows(1).Cells(spalte).Range.text = ueberschriften(i)
                spalte = spalte + 1
            Next i
        End If
    End With
    
    With Options
       .DefaultBorderLineStyle = wdLineStyleSingle
       .DefaultBorderLineWidth = wdLineWidth050pt
       .DefaultBorderColor = wdColorAutomatic
    End With

End Sub

Private Sub anlegenLaborTabelle(datum As String, uhrzeit As String, mytable As Word.Table, ueberschrift As Boolean)
    Dim initiale_zeilen As Long
    ' Nur wenn die Labortabelle eine Ueberschrift bekommen soll,
    ' gleich zu Anfang zwei Zeilen anlegen, sonst nur eine Zeile anlegen
    If ueberschrift Then
        initiale_zeilen = 2
    Else
        initiale_zeilen = 1
    End If
    
    ActiveDocument.Bookmarks("Labor").Select
    Selection.MoveLeft 1 ' Vor der Marke einsetzen
    Selection.TypeParagraph
    Selection.TypeText "Labor vom " & datum & " / " & uhrzeit
    Selection.TypeParagraph
    Set mytable = Selection.Range.Tables.Add(Selection.Range, initiale_zeilen, 7, wdWord9TableBehavior, wdAutoFitFixed)
    With mytable
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    
    With Options
       .DefaultBorderLineStyle = wdLineStyleSingle
       .DefaultBorderLineWidth = wdLineWidth050pt
       .DefaultBorderColor = wdColorAutomatic
    End With
    
    If ueberschrift Then
        Selection.TypeText text:="Bezeichnung"
        Selection.MoveRight UNIT:=wdCell
        Selection.TypeText text:="Wert"
        Selection.MoveRight UNIT:=wdCell
        Selection.TypeText text:="Einheit"
        Selection.MoveRight UNIT:=wdCell
        Selection.TypeText text:="Normalwert"
        
        ' ---------------------neu 17.11.02 --------------------------
        Selection.MoveRight UNIT:=wdCell
        Selection.TypeText text:="Datum"
        Selection.MoveRight UNIT:=wdCell
        Selection.TypeText text:="Zeit"
        Selection.MoveRight UNIT:=wdCell
        Selection.TypeText text:="SortNr"
        '-------------------------------------------------------------
        
        Selection.MoveLeft UNIT:=wdCell
        Selection.MoveLeft UNIT:=wdCell
        Selection.MoveLeft UNIT:=wdCell
        
        ' ---------------------neu 17.11.02 --------------------------
        Selection.MoveLeft UNIT:=wdCell
        Selection.MoveLeft UNIT:=wdCell
        Selection.MoveLeft UNIT:=wdCell
        '--------------------------------------------------------------
        
        ' Selection.MoveRight Unit:=wdCharacter, Count:=5, Extend:=wdExtend
        Selection.MoveRight UNIT:=wdCharacter, Count:=8, Extend:=wdExtend
        
        With Selection.Cells
           With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorGray10
            End With
        End With
    End If ' ueberschrift
    Selection.MoveRight UNIT:=wdCharacter, Count:=8, Extend:=wdExtend
    With Selection.Cells
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders.Shadow = False
    End With
    With Options
       .DefaultBorderLineStyle = wdLineStyleSingle
       .DefaultBorderLineWidth = wdLineWidth050pt
       .DefaultBorderColor = wdColorAutomatic
    End With
    
    ' Bezeichnung
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(5.2)
    Selection.Collapse Direction:=wdCollapseStart
    
    ' Wert
    Selection.Move UNIT:=wdColumn, Count:=1
    Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(2.2)
    Selection.Collapse Direction:=wdCollapseStart
    
    ' Einheit
    Selection.Move UNIT:=wdColumn, Count:=1
    Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(2.2)
    Selection.Collapse Direction:=wdCollapseStart
    
    ' Normalwert
    Selection.Move UNIT:=wdColumn, Count:=1
    Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(3.3)
    
    ' Datum
    Selection.Move UNIT:=wdColumn, Count:=1
    Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(3.3)
    
    ' Zeit
    Selection.Move UNIT:=wdColumn, Count:=1
    Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(3.3)
    
    ' Sortnr
    Selection.Move UNIT:=wdColumn, Count:=1
    Selection.SelectColumn
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(3.3)
End Sub
Public Sub abschliessenTabelle(mytable As Word.Table, mit_ueberschrift As Boolean)
    With mytable
        If .Rows.Count > 1 Then
            With .Rows.Borders(wdBorderHorizontal)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
        End If
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    End With
End Sub

Public Sub abschliessenLaborTabelle(mytable As Word.Table, mit_ueberschrift As Boolean)
    With mytable
        If .Rows.Count > 1 Then
            With .Rows.Borders(wdBorderHorizontal)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
        End If
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        
        .Sort mit_ueberschrift, 7, wdSortFieldNumeric, wdSortOrderAscending
        .Columns(7).Delete
        .Columns(6).Delete
        .Columns(5).Delete
        .Columns(1).Select
        Selection.SelectColumn
        Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
        Selection.Columns.PreferredWidth = CentimetersToPoints(5.2)
        .Columns(2).Select
        Selection.SelectColumn
        Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
        Selection.Columns.PreferredWidth = CentimetersToPoints(3.2)
        .Columns(3).Select
        Selection.SelectColumn
        Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
        Selection.Columns.PreferredWidth = CentimetersToPoints(3.2)
        .Columns(4).Select
        Selection.SelectColumn
        Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
        Selection.Columns.PreferredWidth = CentimetersToPoints(3.2)
    End With
End Sub

Public Sub DoDiagnosen(tb As Collection)
    If tb Is Nothing Then Exit Sub
    If tb.Count = 0 Then Exit Sub
    Dim tabrow As Collection, key As String, langtext As String, is_empty As Boolean, mytable As Word.Table, tr As Word.row
    Dim breiten(), ueberschriften()
    breiten = Array(CentimetersToPoints(2), CentimetersToPoints(14.5))
    Selection.StartOf UNIT:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[USER:DIATAB-DIA_DKEY]"
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
    If Selection.Find.Found Then
        For Each tabrow In tb
            getTabVar tabrow, "DIA_DKEY", key, is_empty, False
            getTabVar tabrow, "DIA_LANGTEXT<L>", langtext, is_empty, False
            langtext = Replace(langtext, vbCrLf, vbCrLf & vbTab)
            Selection.TypeText key & vbTab & langtext
            Selection.TypeParagraph
        Next
    
    
'        anlegenTabelle mytable, breiten(), ueberschriften(), False
'        Dim erste_zeile As Boolean
'        erste_zeile = True
'        For Each tabrow In tb
'            getTabVar tabrow, "DIA_DKEY", key, is_empty, False
'            getTabVar tabrow, "DIA_LANGTEXT<L>", langtext, is_empty, False
'            If erste_zeile = False Then
'                Set tr = mytable.Rows.Add()
'            Else
'                Set tr = mytable.Rows(1)
'                erste_zeile = False
'            End If
'            tr.Cells(1).Range.text = key
'            tr.Cells(2).Range.text = langtext
'        Next
'        abschliessenTabelle mytable, False
'        With mytable
'            .Rows.SetLeftIndent LeftIndent:=0, RulerStyle:=wdAdjustNone
'            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
'            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
'            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
'            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
'            .Borders.Shadow = False
'        End With
'        With Options
'            .DefaultBorderLineStyle = wdLineStyleSingle
'            .DefaultBorderLineWidth = wdLineWidth050pt
'            .DefaultBorderColor = wdColorAutomatic
'        End With
    End If
End Sub

' eingefügt von x-tention im Oktober 2022
Public Sub DoProzeduren(tb As Collection)
    If tb Is Nothing Then Exit Sub
    If tb.Count = 0 Then Exit Sub
    Dim tabrow As Collection, key As String, langtext As String, is_empty As Boolean, mytable As Word.Table, tr As Word.row
    Dim breiten(), ueberschriften()
    breiten = Array(CentimetersToPoints(2), CentimetersToPoints(14.5))
    Selection.StartOf UNIT:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[USER:PROCTAB-PROC_ICPML]"
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
    If Selection.Find.Found Then
        For Each tabrow In tb
            getTabVar tabrow, "PROC_ICPML", key, is_empty, False
            getTabVar tabrow, "PROC_LANGTEXT", langtext, is_empty, False
            langtext = Replace(langtext, vbCrLf, vbCrLf & vbTab)
            Selection.TypeText key & vbTab & langtext
            Selection.TypeParagraph
        Next
    End If
End Sub

Public Sub DoBlutkomp(tb As Collection)
    'Todo, soykas, 4.5.2018
    If tb Is Nothing Then Exit Sub
    If tb.Count = 0 Then Exit Sub
    
    ActiveDocument.Bookmarks("Blutkomp").Select
    Selection.MoveLeft Count:=1
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText "Blutkomponenten:"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.MoveLeft Count:=1
    Dim tbl As Word.Table, breiten(), ueberschriften(), tr As Word.row, size As Integer
    breiten = Array(CentimetersToPoints(2.3), CentimetersToPoints(7), CentimetersToPoints(2), CentimetersToPoints(2), CentimetersToPoints(4))
    ueberschriften = Array("Datum", "Blutkomponente", "Menge", "Einheit", "Charge")
    anlegenTabelle tbl, breiten, ueberschriften, True
    Dim tabrow As Collection, i As Integer, is_empty As Boolean
    For Each tabrow In tb
        Dim ZEINHEIT As String, ZBLUTKOMP As String, ZCHARGE As String, ZGEBENAM As String, ZMENGE As String
        i = i + 1
        getTabVar tabrow, "ZGEBENAM", ZGEBENAM, is_empty, False
        getTabVar tabrow, "ZBLUTKOMP", ZBLUTKOMP, is_empty, False
        getTabVar tabrow, "ZMENGE", ZMENGE, is_empty, False
        getTabVar tabrow, "ZEINHEIT", ZEINHEIT, is_empty, False
        getTabVar tabrow, "ZCHARGE", ZCHARGE, is_empty, False
        If i = 1 Then
            Set tr = tbl.Rows(2)
        Else
            Set tr = tbl.Rows.Add
        End If
        If ZGEBENAM > "" And ZGEBENAM <> "00000000" Then
            ZGEBENAM = ISHMED_Datum(ZGEBENAM)
        End If
        If IsNumeric(ZMENGE) Then
            ZMENGE = Format(getDecimalFormat(ZMENGE), "##,#0")
        End If
        tr.Cells(1).Range.text = ZGEBENAM
        tr.Cells(2).Range.text = ZBLUTKOMP
        tr.Cells(3).Range.text = ZMENGE
        tr.Cells(4).Range.text = ZEINHEIT
        tr.Cells(5).Range.text = ZCHARGE
    Next
    abschliessenTabelle tbl, True
End Sub

Public Sub DoLabor(cp As Collection, tb As Collection, ueberschrift As Boolean)
    '***********************************************************************
    '* Ausgabe der Laborwerte
    '***********************************************************************
    Dim anzahl_gewaehlt As Long
    Dim is_empty As Boolean
    Dim itm
    Dim i As Long, y As Long, Tag_2 As String
    Dim tbl As Word.Table
    Dim normal As Boolean
    Dim tabrow As String
    Dim dbl_normalwert, dbl_normalwert1, dbl_wert As Double
    
    Dim datum As String, uhrzeit As String, zeile As Collection, zeilen As Long
    Dim LADATUM As String, LATIME As String, KATTEXT As String, NORMALWERT, ZUNIT, rank, Value, gewaehlt As Integer, ausschluss As Boolean

    If tb Is Nothing Then Exit Sub

    zeilen = 0
    For Each zeile In tb
        LADATUM = ""
        LATIME = ""
        KATTEXT = ""
        NORMALWERT = ""
        ZUNIT = ""
        Value = ""
        rank = ""
        ausschluss = False
        tabrow = zeile.Item(1).Item("row_key")
        gewaehlt = zeile.Item("gewaehlt")
        getTabVar zeile, "N2LADATUM", LADATUM, is_empty, False
        getTabVar zeile, "N2LATIME", LATIME, is_empty, False
        getTabVar zeile, "N2KATTEXT", KATTEXT, is_empty, False
        ausschluss = getFromCol(gcolLaborAusschliessen, KATTEXT, Value)
        getTabVar zeile, "N2NORMAL", NORMALWERT, is_empty, False
        getTabVar zeile, "N2UNIT", ZUNIT, is_empty, False
        getTabVar zeile, "N2VALUE", Value, is_empty, False
        getTabVar zeile, "RANK1", rank, is_empty, False
        
        If gewaehlt And Not ausschluss Then
            Dim row As Word.row
            If datum <> LADATUM Or uhrzeit <> LATIME Then
                ' Datum hat gewechselt, neue Untersuchung, neue Tabelle
                datum = LADATUM
                uhrzeit = LATIME
                If Not tbl Is Nothing Then
                    ' Wenn es nicht die erste Tabelle ist, sortieren
                    abschliessenLaborTabelle tbl, ueberschrift
                End If
                anlegenLaborTabelle ISHMED_Datum(datum), ISHMED_Uhrzeit(uhrzeit), tbl, ueberschrift
                zeilen = 0
            End If
            If zeilen = 0 Then
                If ueberschrift Then
                    ' Wenn ueberschrift, dann ist die erste Inhaltszeile die
                    ' zweite, sonst direkt Werte in die erste (und in die folgenden)
                    ' Zeilen eintragen. Ueberschrift oder nicht wird von "ganz oben"
                    ' gesteuert und nach dem ersten Labor auf "False" gesetzt.
                    Set row = tbl.Rows(2)
                Else
                    Set row = tbl.Rows(1)
                End If
            Else
                Set row = tbl.Rows.Add
            End If
            
            ' Pruefen ob der Wert im Normbereich liegt und
            ' die Variable "normal" entsprechend setzen
            'Stop
            normal = True
            If IsNumeric(getDecimalFormat(Value)) Then
                dbl_wert = CDbl(getDecimalFormat(Value))
                If Left(NORMALWERT, 1) = ">" Then
                    If IsNumeric(getDecimalFormat(Mid(NORMALWERT, 2))) Then
                        dbl_normalwert = CDbl(getDecimalFormat(Mid(NORMALWERT, 2)))
                        normal = dbl_wert > dbl_normalwert
                    End If
                ElseIf Left(NORMALWERT, 1) = "<" Then
                    If IsNumeric(getDecimalFormat(Mid(NORMALWERT, 2))) Then
                        dbl_normalwert = CDbl(getDecimalFormat(Mid(NORMALWERT, 2)))
                        normal = dbl_wert < dbl_normalwert
                    End If
                Else
                    Dim strpos As Integer
                    ' Fuehrendes Minuszeichen nicht finden
                    strpos = InStr(2, NORMALWERT, "-")
                    If strpos > 0 Then
                        If IsNumeric(getDecimalFormat(Mid(NORMALWERT, 1, strpos - 1))) _
                            And IsNumeric(getDecimalFormat(Mid(NORMALWERT, strpos + 1))) Then
                            dbl_normalwert = CDbl(getDecimalFormat(Mid(NORMALWERT, 1, strpos - 1)))
                            dbl_normalwert1 = CDbl(getDecimalFormat(Mid(NORMALWERT, strpos + 1)))
                            normal = dbl_wert >= dbl_normalwert And dbl_wert <= dbl_normalwert1
                        End If
                    End If
                End If
            End If
                    
            zeilen = zeilen + 1
            row.Cells(1).Range.text = KATTEXT
            row.Cells(2).Range.text = Value
            row.Cells(3).Range.text = ZUNIT
            row.Cells(4).Range.text = NORMALWERT
            row.Cells(5).Range.text = tabrow
            row.Cells(7).Range.text = rank
            row.Select
            Selection.Font.Bold = Not normal
        End If
    Next
    If Not tbl Is Nothing Then
        ' Auch die letzte (oder einzige) Tabelle sortieren
        abschliessenLaborTabelle tbl, ueberschrift
    End If
End Sub

Public Sub DoEKG(gcolEKG As Collection, rhythmus As Collection)
    ' tb sind Diagnosen, werden in SAP hier nicht gepflegt, wird ignoriert
    Dim werte As String, tabrow As Collection, i As Long
    Dim pwelle As Double
    Dim zeile As Collection
    Dim datum, HERZFRERA, HERZFREWE, HERZRHLA2, HERZRHLA3, HERZRHLAG, HERZRHSON, HFFELD, PWELFORM, _
        PWELHOEH, STASC_AVF, STASC_AVL, STASC_AVR, STASC_I, STASC_II, STASC_III, STASC_V1, STASC_V2, _
        STASC_V3, STASC_V4, STASC_V5, STASC_V6, STBOG_AVF, STBOG_AVL, STBOG_AVR, STBOG_I, STBOG_II, _
        STBOG_III, STBOG_V1, STBOG_V2, STBOG_V3, STBOG_V4, STBOG_V5, STBOG_V6, STDESCAVF, STDESCAVL, _
        STDESCAVR, STDESCIII, STDESC_I, STDESC_II, STDESC_V1, STDESC_V2, STDESC_V3, STDESC_V4, STDESC_V5, _
        STDESC_V6, STHORIAVF, STHORIAVL, STHORIAVR, STHORIIII, STHORI_I, STHORI_II, STHORI_V1, STHORI_V2, _
        STHORI_V3, STHORI_V4, STHORI_V5, STHORI_V6, STMULDAVF, STMULDAVL, STMULDAVR, STMULDIII, STMULD_I, _
        STMULD_II, STMULD_V1, STMULD_V2, STMULD_V3, STMULD_V4, STMULD_V5, STMULD_V6, STNORMAVF, STNORMAVL, _
        STNORMAVR, STNORMIII, STNORM_I, STNORM_II, STNORM_V1, STNORM_V2, STNORM_V3, STNORM_V4, STNORM_V5, _
        STNORM_V6, TBEURTEIL, TXTUNTERS, UNTERSCHF, UNTERSUCH, ZEIRELQT, zeit, ZEITP, ZEITPQ, ZEITQRSV1, _
        ZEITQRSV6, ZEITQT, HERZRHART
    Dim hat_rhythmus As Boolean, is_empty As Boolean
    
    hat_rhythmus = False
    If Not rhythmus Is Nothing Then
        hat_rhythmus = rhythmus.Count > 1
    End If
    
    getDocVar gcolEKG, "ZDATUM", datum, is_empty, False
    getDocVar gcolEKG, "ZHERZFRERA", HERZFRERA, is_empty, False
    getDocVar gcolEKG, "ZHERZFREWE", HERZFREWE, is_empty, False
    getDocVar gcolEKG, "ZHERZRHLA2", HERZRHLA2, is_empty, False
    getDocVar gcolEKG, "ZHERZRHLA3", HERZRHLA3, is_empty, False
    getDocVar gcolEKG, "ZHERZRHLAG", HERZRHLAG, is_empty, False
    getAnyDocVar gcolEKG, Array("ZHERZRHSON<L>", "ZEHERZRHSO<L>"), HERZRHSON, is_empty, False
    getDocVar gcolEKG, "ZHFFELD", HFFELD, is_empty, False
    getDocVar gcolEKG, "ZPWELFORM", PWELFORM, is_empty, False
    getDocVar gcolEKG, "ZPWELHOEH", PWELHOEH, is_empty, False
    getDocVar gcolEKG, "ZSTASC_AVF", STASC_AVF, is_empty, False
    getDocVar gcolEKG, "ZSTASC_AVL", STASC_AVL, is_empty, False
    getDocVar gcolEKG, "ZSTASC_AVR", STASC_AVR, is_empty, False
    getDocVar gcolEKG, "ZSTASC_I", STASC_I, is_empty, False
    getDocVar gcolEKG, "ZSTASC_II", STASC_II, is_empty, False
    getDocVar gcolEKG, "ZSTASC_III", STASC_III, is_empty, False
    getDocVar gcolEKG, "ZSTASC_V1", STASC_V1, is_empty, False
    getDocVar gcolEKG, "ZSTASC_V2", STASC_V2, is_empty, False
    getDocVar gcolEKG, "ZSTASC_V3", STASC_V3, is_empty, False
    getDocVar gcolEKG, "ZSTASC_V4", STASC_V4, is_empty, False
    getDocVar gcolEKG, "ZSTASC_V5", STASC_V5, is_empty, False
    getDocVar gcolEKG, "ZSTASC_V6", STASC_V6, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_AVF", STBOG_AVF, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_AVL", STBOG_AVL, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_AVR", STBOG_AVR, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_I", STBOG_I, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_II", STBOG_II, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_III", STBOG_III, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_V1", STBOG_V1, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_V2", STBOG_V2, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_V3", STBOG_V3, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_V4", STBOG_V4, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_V5", STBOG_V5, is_empty, False
    getDocVar gcolEKG, "ZSTBOG_V6", STBOG_V6, is_empty, False
    getDocVar gcolEKG, "ZSTDESCAVF", STDESCAVF, is_empty, False
    getDocVar gcolEKG, "ZSTDESCAVL", STDESCAVL, is_empty, False
    getDocVar gcolEKG, "ZSTDESCAVR", STDESCAVR, is_empty, False
    getDocVar gcolEKG, "ZSTDESCIII", STDESCIII, is_empty, False
    getDocVar gcolEKG, "ZSTDESC_I", STDESC_I, is_empty, False
    getDocVar gcolEKG, "ZSTDESC_II", STDESC_II, is_empty, False
    getDocVar gcolEKG, "ZSTDESC_V1", STDESC_V1, is_empty, False
    getDocVar gcolEKG, "ZSTDESC_V2", STDESC_V2, is_empty, False
    getDocVar gcolEKG, "ZSTDESC_V3", STDESC_V3, is_empty, False
    getDocVar gcolEKG, "ZSTDESC_V4", STDESC_V4, is_empty, False
    getDocVar gcolEKG, "ZSTDESC_V5", STDESC_V5, is_empty, False
    getDocVar gcolEKG, "ZSTDESC_V6", STDESC_V6, is_empty, False
    getDocVar gcolEKG, "ZSTHORIAVF", STHORIAVF, is_empty, False
    getDocVar gcolEKG, "ZSTHORIAVL", STHORIAVL, is_empty, False
    getDocVar gcolEKG, "ZSTHORIAVR", STHORIAVR, is_empty, False
    getDocVar gcolEKG, "ZSTHORIIII", STHORIIII, is_empty, False
    getDocVar gcolEKG, "ZSTHORI_I", STHORI_I, is_empty, False
    getDocVar gcolEKG, "ZSTHORI_II", STHORI_II, is_empty, False
    getDocVar gcolEKG, "ZSTHORI_V1", STHORI_V1, is_empty, False
    getDocVar gcolEKG, "ZSTHORI_V2", STHORI_V2, is_empty, False
    getDocVar gcolEKG, "ZSTHORI_V3", STHORI_V3, is_empty, False
    getDocVar gcolEKG, "ZSTHORI_V4", STHORI_V4, is_empty, False
    getDocVar gcolEKG, "ZSTHORI_V5", STHORI_V5, is_empty, False
    getDocVar gcolEKG, "ZSTHORI_V6", STHORI_V6, is_empty, False
    getDocVar gcolEKG, "ZSTMULDAVF", STMULDAVF, is_empty, False
    getDocVar gcolEKG, "ZSTMULDAVL", STMULDAVL, is_empty, False
    getDocVar gcolEKG, "ZSTMULDAVR", STMULDAVR, is_empty, False
    getDocVar gcolEKG, "ZSTMULDIII", STMULDIII, is_empty, False
    getDocVar gcolEKG, "ZSTMULD_I", STMULD_I, is_empty, False
    getDocVar gcolEKG, "ZSTMULD_II", STMULD_II, is_empty, False
    getDocVar gcolEKG, "ZSTMULD_V1", STMULD_V1, is_empty, False
    getDocVar gcolEKG, "ZSTMULD_V2", STMULD_V2, is_empty, False
    getDocVar gcolEKG, "ZSTMULD_V3", STMULD_V3, is_empty, False
    getDocVar gcolEKG, "ZSTMULD_V4", STMULD_V4, is_empty, False
    getDocVar gcolEKG, "ZSTMULD_V5", STMULD_V5, is_empty, False
    getDocVar gcolEKG, "ZSTMULD_V6", STMULD_V6, is_empty, False
    getDocVar gcolEKG, "ZSTNORMAVF", STNORMAVF, is_empty, False
    getDocVar gcolEKG, "ZSTNORMAVL", STNORMAVL, is_empty, False
    getDocVar gcolEKG, "ZSTNORMAVR", STNORMAVR, is_empty, False
    getDocVar gcolEKG, "ZSTNORMIII", STNORMIII, is_empty, False
    getDocVar gcolEKG, "ZSTNORM_I", STNORM_I, is_empty, False
    getDocVar gcolEKG, "ZSTNORM_II", STNORM_II, is_empty, False
    getDocVar gcolEKG, "ZSTNORM_V1", STNORM_V1, is_empty, False
    getDocVar gcolEKG, "ZSTNORM_V2", STNORM_V2, is_empty, False
    getDocVar gcolEKG, "ZSTNORM_V3", STNORM_V3, is_empty, False
    getDocVar gcolEKG, "ZSTNORM_V4", STNORM_V4, is_empty, False
    getDocVar gcolEKG, "ZSTNORM_V5", STNORM_V5, is_empty, False
    getDocVar gcolEKG, "ZSTNORM_V6", STNORM_V6, is_empty, False
    getAnyDocVar gcolEKG, Array("ZTXTUNTERS<L>", "ZETXTUNTER<L>", "ZEBEFUND<L>"), TXTUNTERS, is_empty, False
    getDocVar gcolEKG, "ZUNTERSCHF", UNTERSCHF, is_empty, False
    getDocVar gcolEKG, "ZUNTERSUCH", UNTERSUCH, is_empty, False
    getDocVar gcolEKG, "ZZEIRELQT", ZEIRELQT, is_empty, False
    getDocVar gcolEKG, "ZZEIT", zeit, is_empty, False
    getDocVar gcolEKG, "ZZEITP", ZEITP, is_empty, False
    getDocVar gcolEKG, "ZZEITPQ", ZEITPQ, is_empty, False
    getDocVar gcolEKG, "ZZEITQRSV1", ZEITQRSV1, is_empty, False
    getDocVar gcolEKG, "ZZEITQRSV6", ZEITQRSV6, is_empty, False
    getDocVar gcolEKG, "ZZEITQT", ZEITQT, is_empty, False
    getAnyDocVar gcolEKG, Array("ZTBEURTEIL<L>", "ZEBEURTEIL<L>"), TBEURTEIL, is_empty, False
    
    ActiveDocument.Bookmarks("EKG").Select
    Selection.MoveLeft , 1
    set_format
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="EKG-Befund vom "
    Selection.TypeText text:=ISHMED_Datum(datum & "") & ":"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    Selection.Paragraphs.TabStops(1).Clear
    Selection.Paragraphs.TabStops(2).Clear
    Selection.Paragraphs.TabStops(3).Clear
    Selection.Paragraphs.TabStops(4).Clear
    Selection.Paragraphs.TabStops.Add InchesToPoints(1)
    Selection.Paragraphs.TabStops.Add InchesToPoints(1.8)
    Selection.Paragraphs.TabStops.Add InchesToPoints(2.4)
    Selection.Paragraphs.TabStops.Add InchesToPoints(3)
    Selection.Paragraphs.TabStops.Add InchesToPoints(3.4)
    
    If (HERZFREWE <> 0) Then
      Selection.TypeText "Herzfrequenz: " & Format(HERZFREWE, "##0 ") & "/min "
    Else
      Selection.TypeText "Herzfrequenz: "
      Select Case HERZFRERA
        Case 1
          Selection.TypeText " normalfrequent "
        Case 2
          Selection.TypeText " Sinustachykardie "
        Case 3
          Selection.TypeText " Sinusbradykardie "
        Case 4
          Selection.TypeText " Tachyarrhtythmie "
        Case 5
          Selection.TypeText " Bradyarrhtythmie "
      End Select
      Selection.TypeParagraph
    End If
   
    If (RTrim(HERZRHART) <> "") Then
      Selection.TypeText "Herzrhythmus: " & HERZRHART
      Selection.TypeParagraph
    End If
    
    If hat_rhythmus Then
        Dim t As String
        For Each tabrow In rhythmus
            getTabVar tabrow, "ZHERZRHART", HERZRHART, is_empty, False
            If Not is_empty Then
                If t > "" Then t = t & ", "
                t = t & HERZRHART
            End If
        Next
        Selection.TypeText "Rhythmus: "
        Selection.TypeText t
        Selection.TypeParagraph
    End If
    
    If (RTrim(HERZRHLAG) <> "") Then
       Selection.TypeText "Lagetyp: " & HERZRHLAG
         Selection.TypeParagraph
    End If
    
    If (RTrim(HERZRHLA2) <> "") Then
       Selection.TypeText "Lagetyp: " & HERZRHLA2
         Selection.TypeParagraph
    End If
    
    If (RTrim(HERZRHLA3) <> "") Then
       Selection.TypeText "Lagetyp: " & HERZRHLA3
       Selection.TypeParagraph
    End If
    
    print_langtext HERZRHSON
    
    werte = ""
    
    If (ZEITP <> 0) Then
         werte = "P-Zeit: " & Format(ZEITP, "##0 msec ") & Chr(9)
    End If
    
    If ZEITPQ <> 0 Then
         werte = werte & "PQ-Zeit: " & Format(ZEITPQ, "##0 msec ") & Chr(9)
    End If
    
    If ZEITQRSV1 <> 0 Then
         werte = werte & "QRS-Zeit V1: " & Format(ZEITQRSV1, "##0 msec ")
    End If
    
    If (RTrim(werte) <> "") Then
      Selection.TypeText Chr(9) & werte
      Selection.TypeParagraph
    End If
    werte = ""
    If ZEITQRSV6 <> 0 Then
        werte = "QRS-Zeit V6: " & Format(ZEITQRSV6, "##0 msec") & Chr(9)
    End If
    
    If ZEITQT <> 0 Then
        werte = werte & "QT-Zeit: " & Format(ZEITQT, "##0 msec") & Chr(9)
    End If
    
    If ZEIRELQT <> 0 Then
        werte = werte & "rel. QT-Zeit: " & Format(ZEIRELQT, "##0 msec")
    End If
    
    If (RTrim(werte) <> "") Then
      Selection.TypeText Chr(9) & werte
      Selection.TypeParagraph
    End If
    
    
    If PWELHOEH <> 0 Then
       Selection.TypeText "P-Welle Höhe: " & Left$(PWELHOEH, 1) & "," & Right$(PWELHOEH, 1) & " mV" & Chr(9)
       Selection.TypeParagraph
    End If
    
    If RTrim(PWELFORM) <> "" Then
        Selection.TypeText "Form: " & PWELFORM
        Selection.TypeParagraph
    End If
   
    print_langtext TXTUNTERS
    print_langtext TBEURTEIL
End Sub


' EM_20250701, Ausgabe des Tomtec Echokardio-Befund
' Hier die relevanten Felder:
' ZTECHOKARD  1   Z00000000000000J    ZUNAUFTRAG        --> befTitel
' ZTECHOKARD  1   Z00000000000000J    ZDATUM
' ZTECHOKARD  1   Z00000000000000J    ZEBEFUND
' ZTECHOKARD  1   Z00000000000000J    ZEBEURTEIL
' ZTECHOKARD  1   Z00000000000000J    ZEPROCEDER
' ZTECHOKARD  1   Z00000000000000J    ZFRAGEST

' ZTECHOKARD  1   Z00000000000000K    ECHOEINH
' ZTECHOKARD  1   Z00000000000000K    ECHOKAPIT
' ZTECHOKARD  1   Z00000000000000K    ECHONORM
' ZTECHOKARD  1   Z00000000000000K    ECHOUEBER2
' ZTECHOKARD  1   Z00000000000000K    ECHOWERT
'
' EM_20250930: Ausgabe Titel geändert in befTitel & " vom " & ISHMED_Datum(datum & "") & ":"
'
'Public Sub DoTEchokard(tb As Collection, counter As Integer)
Public Sub DoTEchokard(gcolEcho As Collection, gcolMesswerte As Collection)
    
    If gcolEcho Is Nothing Then Exit Sub
    If gcolEcho.Count = 0 Then Exit Sub
    
    Dim tabrow As Collection, i As Integer, is_empty As Boolean
    Dim datum As String, befTitel As String, collText As String
    Dim echoBef As String, befundTextLen As Long
    Dim echoBeur As String, beurTextLen As Long
    Dim echoProz As String, prozTextLen As Long
    Dim echoFrag As String, fragTextLen As Long
    
    For Each tabrow In gcolEcho
        getTabVar tabrow, "ZUNAUFTRAG", befTitel, is_empty, False
        getTabVar tabrow, "ZDATUM", datum, is_empty, False
        getTabVar tabrow, "ZEBEFUND<L>", echoBef, is_empty, False
        getTabVar tabrow, "ZEBEURTEIL<L>", echoBeur, is_empty, False
        getTabVar tabrow, "ZEPROCEDER<L>", echoProz, is_empty, False
        getTabVar tabrow, "ZFRAGEST", echoFrag, is_empty, False
    Next
    
   ' Kopf
    ActiveDocument.Bookmarks("EKG").Select
    Selection.MoveLeft , 1
    set_format
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    ' Selection.TypeText text:="Echokardio-Befund (Tomtec) vom "
    Selection.TypeText befTitel & " vom " & ISHMED_Datum(datum & "") & ":"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
        
    ' Abschnitt Messwerte
    If Not gcolMesswerte Is Nothing Then
        If gcolMesswerte.Count > 0 Then
            Dim line As Collection, gefunden As Boolean, _
                initialized As Boolean, kapitel As String
            Selection.TypeParagraph
            Selection.Font.Italic = True
            Selection.TypeText "Messwerte:"
            Selection.Font.Italic = False
            Selection.TypeParagraph
'

'
            For Each tabrow In gcolMesswerte
                Dim ECHOUEBER2, ECHOWERT, ECHOEINH
                getTabVar tabrow, "ECHOUEBER2", ECHOUEBER2, is_empty, False
                getTabVar tabrow, "ECHOWERT", ECHOWERT, is_empty, False
                getTabVar tabrow, "ECHOEINH", ECHOEINH, is_empty, False
                If Len(Trim$(ECHOWERT)) > 0 Then
                    Debug.Print "xt---" & collText
                    If Len(collText) = 0 Then 'on new
                        collText = collText & ECHOUEBER2 & " " & ECHOWERT & " " & ECHOEINH & ";"
                    Else
                        collText = collText & " " & ECHOUEBER2 & " " & ECHOWERT & " " & ECHOEINH & ";"
                    End If
                End If
            Next
            Selection.TypeText collText
            Selection.TypeParagraph
        End If
    End If
    
    ' Abschnitt Befund
    befundTextLen = Len(echoBef)
    If befundTextLen > 0 Then
        Selection.TypeParagraph
        Selection.Font.Italic = True
        Selection.TypeText text:="Befund: "
        Selection.Font.Italic = False
        Selection.TypeParagraph
        Selection.TypeText text:=echoBef
        Selection.TypeParagraph
    End If
    
    ' Abschnitt Beurteilung
    beurTextLen = Len(echoBeur)
    If beurTextLen > 0 Then
        Selection.TypeParagraph
        Selection.Font.Italic = True
        Selection.TypeText text:="Beurteilung: "
        Selection.Font.Italic = False
        Selection.TypeParagraph
        Selection.TypeText text:=echoBeur
        Selection.TypeParagraph
    End If
    
    ' Abschnitt Prozedere
    prozTextLen = Len(echoProz)
    If prozTextLen > 0 Then
        Selection.TypeParagraph
        Selection.Font.Italic = True
        Selection.TypeText text:="Prozedere: "
        Selection.Font.Italic = False
        Selection.TypeParagraph
        Selection.TypeText text:=echoProz
        Selection.TypeParagraph
    End If
    
    ' Abschnitt Fragestellung
    fragTextLen = Len(echoFrag)
    If fragTextLen > 0 Then
        Selection.TypeParagraph
        Selection.Font.Italic = True
        Selection.TypeText text:="Fragestellung: "
        Selection.Font.Italic = False
        Selection.TypeParagraph
        Selection.TypeText text:=echoFrag
        Selection.TypeParagraph
    End If
    
End Sub

' EM_20250701, Ausgabe des Tomcat Sono-Befund
Public Sub DoTSonograf(tb As Collection, counter As Integer)
    
    If tb Is Nothing Then Exit Sub
    If tb.Count = 0 Then Exit Sub
    
    Dim datum As String
    Dim sonoBef As String, befTextLen As Long
    Dim sonoBeur As String, beurTextLen As Long
    Dim sonoProz As String, prozTextLen As Long
    Dim tabrow As Collection, i As Integer, is_empty As Boolean

    For Each tabrow In tb
        getTabVar tabrow, "ZDATUM", datum, is_empty, False
        getTabVar tabrow, "ZEBEFUND<L>", sonoBef, is_empty, False
        getTabVar tabrow, "ZEBEURTEIL<L>", sonoBeur, is_empty, False
        getTabVar tabrow, "ZEPROCEDER<L>", sonoProz, is_empty, False
    Next

    ' Kopf
    ActiveDocument.Bookmarks("EKG").Select
    Selection.MoveLeft , 1
    set_format
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="Sonographie-Befund (Tomtec) vom "
    Selection.TypeText text:=ISHMED_Datum(datum & "") & ":"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
        
    ' Abschnitt Befund
    befTextLen = Len(sonoBef)
    If befTextLen > 0 Then
        Selection.TypeParagraph
        Selection.Font.Italic = True
        Selection.TypeText text:="Befund: "
        Selection.Font.Italic = False
        Selection.TypeParagraph
        Selection.TypeText text:=sonoBef
        Selection.TypeParagraph
        
    End If
    
    ' Abschnitt Beurteilung
    beurTextLen = Len(sonoBeur)
    If beurTextLen > 0 Then
        Selection.TypeParagraph
        Selection.Font.Italic = True
        Selection.TypeText text:="Beurteilung: "
        Selection.Font.Italic = False
        Selection.TypeParagraph
        Selection.TypeText text:=sonoBeur
        Selection.TypeParagraph
    End If
    
    ' Abschnitt Prozedere
    prozTextLen = Len(sonoProz)
    If prozTextLen > 0 Then
        Selection.TypeParagraph
        Selection.Font.Italic = True
        Selection.TypeText text:="Prozedere: "
        Selection.Font.Italic = False
        Selection.TypeParagraph
        Selection.TypeText text:=sonoProz
        Selection.TypeParagraph
    End If
    
End Sub

Public Sub DoDialyse(gcolDialyse As Collection)
    Dim itm As Variant, is_empty As Boolean
    Dim datum As String, DIALART, DIALFILT, DIALINSU, DIAL_ZEI, DIA_FITXT, HFFELD, TXTUNTERS
    
    getDocVar gcolDialyse, "ZDATUM", datum, is_empty, False
    getDocVar gcolDialyse, "ZDIALART", DIALART, is_empty, False
    getDocVar gcolDialyse, "ZDIALFILT", DIALFILT, is_empty, False
    getDocVar gcolDialyse, "ZDIALINSU", DIALINSU, is_empty, False
    getDocVar gcolDialyse, "ZDIAL_ZEI", DIAL_ZEI, is_empty, False
    getDocVar gcolDialyse, "ZDIA_FITXT", DIA_FITXT, is_empty, False
    getDocVar gcolDialyse, "ZHFFELD", HFFELD, is_empty, False
    getAnyDocVar gcolDialyse, Array("ZTXTUNTERS<L>", "ZETXTUNTER<L>"), TXTUNTERS, is_empty, False

    ActiveDocument.Bookmarks("Dialyse").Select
    set_format
    Selection.Paragraphs.TabStops.Add InchesToPoints(1.8)
    
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="Dialyse vom "
    Selection.TypeText text:=ISHMED_Datum(datum) & ":"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    
    Selection.ParagraphFormat.TabStops.ClearAll
    ' tab bei 3cm
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    
    Selection.TypeText "Niereninsuffizienz: " & vbTab
    If (DIALINSU = 0) Then
      Selection.TypeText "chronisch"
    Else
      Selection.TypeText "akut"
    End If
    Selection.TypeParagraph
    Selection.TypeText "Dialyse: " & vbTab
    Select Case DIALART
        Case "0"
        Selection.TypeText "HD "
        Case "1"
        Selection.TypeText "HF "
        Case "2"
        Selection.TypeText "UF "
        Case "3"
        Selection.TypeText "PF "
    End Select
    
    Selection.TypeText DIAL_ZEI & " Std."
    
    Selection.TypeParagraph
    Selection.TypeText "Filter: " & vbTab
    Select Case DIALFILT
        Case "0"
        Selection.TypeText "OP 05 "
        Case "1"
        Selection.TypeText "F6HPS "
        Case "2"
        Selection.TypeText "F8HPS "
    End Select
    Selection.TypeParagraph
    print_langtext DIA_FITXT
    Selection.TypeParagraph
    print_langtext TXTUNTERS
    
End Sub

Public Sub FillProperties(content As String, property As String)

Dim bolIsValid As Boolean, prop As Object
        
 bolIsValid = False
 For Each prop In ActiveDocument.CustomDocumentProperties
     If Trim(prop.name) = property Then
        bolIsValid = True
         Exit For
     End If
 Next
 If Not bolIsValid Then
    ActiveDocument.CustomDocumentProperties.Add _
    name:=property, LinkToContent:=False, Value:=" ", _
    Type:=msoPropertyTypeString
 End If
 ActiveDocument.CustomDocumentProperties(property).Value = content

End Sub

Sub changeValueName(c)
    Dim name As String, newname As String
    Dim col As Collection
    Dim gefunden As Boolean
    On Local Error GoTo exitChangeValueName
    gefunden = False
    Set col = c.Item("N2KATTEXT")
    name = col.Item("var")
    newname = gcolLaborUmbenennen(name)
    If name <> newname Then
        col.Remove "var"
        col.Add newname, "var"
    End If
    gefunden = True
exitChangeValueName:
End Sub

Function getValueRank(c, key As String) As Long
    Dim col As Object
    Dim itm As Variant
    Dim rank As Long
    Set itm = Nothing
    rank = -1
    On Local Error GoTo exitGetValueRank
    Set col = c.Item("N2KATTEXT")
    itm = col.Item("var")
    rank = gcolLaborRanking(itm)
exitGetValueRank:
    If rank = -1 Then
        wert_ende = wert_ende + 1
        rank = wert_ende
    End If
    getValueRank = rank
End Function

Sub DoAerzte()
Dim name_n(7) As String          ' Für den Nachnamen
Dim anrede(7) As String          ' Für die Anrede
Dim name_k(7) As String          ' Für den kompletten Namen (alte Dokumente)
Dim key As String
Dim lfn As String
Dim content As String
Dim bolIsValid As Boolean
Dim i As Long

For i = 1 To R3Table.RowCount
  key = ""
  content = ""
  lfn = ""
  key = R3Table(i, 1)
  If (Left$(key, 6) = "[ZADRE") Then
    content = R3Table(i, 2)
    lfn = R3Table(i, 3)
    Select Case key
'--------------------------------------------------
'hier beginnt der Hausarzt (erstschriftlich)
'--------------------------------------------------
       Case "[ZADRESSE:Y000001B-ZNAME_H]"
             name_n(1) = content
             
       Case "[ZADRESSE:Y000001B-ZARZTHANRE]"
            
            If Trim(content) = "03" Then
                FillProperties "Drs.", "anrede1"
                anrede(1) = "Sehr geehrte Kollegen "
            End If
            If Trim(content) = "02" Then
                FillProperties "Frau", "anrede1"
                anrede(1) = "Sehr geehrte Frau Kollegin "
            ElseIf Trim(content) = "01" Then
                FillProperties "Herr", "anrede1"
                anrede(1) = "Sehr geehrter Herr Kollege "
            End If
         
        Case "[ZADRESSE:Y000001B-ZISHMHANAM]"
            FillProperties content, "name1"
            name_k(1) = content
            If Left$(content, 4) = "Herr" Then
               anrede(1) = "Sehr geehrter Herr Kollege "
            End If
            If Left$(content, 4) = "Frau" Then
               anrede(1) = "Sehr geehrte Frau Kollegin "
            End If
                   
        Case "[ZADRESSE:Y000001B-ZNAME2_H]"
            FillProperties content, "Firma1"
        
        Case "[ZADRESSE:Y000001B-ZISHMHASTR]"
            FillProperties content, "strass1"
         
        Case "[ZADRESSE:Y000001B-ZISHMHAPLZ]"
            FillProperties content, "plz1"
        
        Case "[ZADRESSE:Y000001B-ZISHMHAORT]"
            FillProperties content, "ort1"
'--------------------------------------------------
'hier beginnt der Einweisende Arzt
'--------------------------------------------------
       Case "[ZADRESSE:Y000001B-ZNAME_E]"
             name_n(2) = content
        
        Case "[ZADRESSE:Y000001B-ZARZTEANRE]"
            If Trim(content) = "03" Then
                FillProperties "Drs.", "anrede1"
                anrede(2) = "Sehr geehrte Kollegen "
            End If
            If Trim(content) = "02" Then
                FillProperties "Frau", "anrede2"
                anrede(2) = "Sehr geehrte Frau Kollegin "
            ElseIf Trim(content) = "01" Then
                FillProperties "Herr", "anrede2"
                anrede(2) = "Sehr geehrter Herr Kollege "
            End If
         
        Case "[ZADRESSE:Y000001B-ZISHMEANAM]"
            FillProperties content, "name2"
            name_k(2) = content
            If Left$(content, 4) = "Herr" Then
               anrede(2) = "Sehr geehrter Herr Kollege "
            End If
            If Left$(content, 4) = "Frau" Then
               anrede(2) = "Sehr geehrte Frau Kollegin "
            End If
            
        Case "[ZADRESSE:Y000001B-ZNAME1_E]"
            FillProperties content, "Firma2"
        
        Case "[ZADRESSE:Y000001B-ZISHMEASTR]"
            FillProperties content, "strass2"
         
        Case "[ZADRESSE:Y000001B-ZISHMEAPLZ]"
            FillProperties content, "plz2"
        
        Case "[ZADRESSE:Y000001B-ZISHMEAORT]"
            FillProperties content, "ort2"
    
'--------------------------------------------------
'hier beginnt Arzt1
'--------------------------------------------------
       Case "[ZADRESSE:Y000001B-ZNAME_1]"
           name_n(3) = content

        Case "[ZADRESSE:Y000001B-ZARZT1ANRE]"
            If Trim(content) = "03" Then
                FillProperties "Drs.", "anrede1"
                anrede(3) = "Sehr geehrte Kollegen "
            End If
            If Trim(content) = "02" Then
                FillProperties "Frau", "anrede3"
                anrede(3) = "Sehr geehrte Frau Kollegin "
            ElseIf Trim(content) = "01" Then
                FillProperties "Herr", "anrede3"
                anrede(3) = "Sehr geehrter Herr Kollege "
            End If
        
        Case "[ZADRESSE:Y000001B-ZARZT1NAME]"
            FillProperties content, "name3"
            name_k(3) = content
            If Left$(content, 4) = "Herr" Then
               anrede(3) = "Sehr geehrter Herr Kollege "
            End If
            If Left$(content, 4) = "Frau" Then
               anrede(3) = "Sehr geehrte Frau Kollegin "
            End If
            
        Case "[ZADRESSE:Y000001B-ZNAME2_1]"
            FillProperties content, "Firma3"
        
        Case "[ZADRESSE:Y000001B-ZARZT1STRA]"
            FillProperties content, "strass3"
         
        Case "[ZADRESSE:Y000001B-ZARZT1PLZ]"
            FillProperties content, "plz3"
        
        Case "[ZADRESSE:Y000001B-ZARZT1ORT]"
            FillProperties content, "ort3"
'-------------------------------------------------------
'hier beginnt Arzt2
'-------------------------------------------------------
        Case "[ZADRESSE:Y000001B-ZNAME_2]"
             name_n(4) = content

        Case "[ZADRESSE:Y000001B-ZARZT2ANRE]"
            If Trim(content) = "03" Then
                FillProperties "Drs.", "anrede1"
                anrede(4) = "Sehr geehrte Kollegen "
            End If
            If Trim(content) = "02" Then
                FillProperties "Frau", "anrede4"
                anrede(4) = "Sehr geehrte Frau Kollegin "
            ElseIf Trim(content) = "01" Then
                FillProperties "Herr", "anrede4"
                anrede(4) = "Sehr geehrter Herr Kollege "
            End If
        
        Case "[ZADRESSE:Y000001B-ZARZT2NAME]"
            FillProperties content, "name4"
            name_k(4) = content
            If Left$(content, 4) = "Herr" Then
               anrede(4) = "Sehr geehrter Herr Kollege "
            End If
            If Left$(content, 4) = "Frau" Then
               anrede(4) = "Sehr geehrte Frau Kollegin "
            End If
            
        Case "[ZADRESSE:Y000001B-ZNAME2_2]"
            FillProperties content, "Firma4"
            
        Case "[ZADRESSE:Y000001B-ZARZT2STRA]"
            FillProperties content, "strass4"
         
        Case "[ZADRESSE:Y000001B-ZARZT2PLZ]"
            FillProperties content, "plz4"
        
        Case "[ZADRESSE:Y000001B-ZARZT2ORT]"
            FillProperties content, "ort4"
'-------------------------------------------------------
'hier beginnt Arzt3
'-------------------------------------------------------
        Case "[ZADRESSE:Y000001B-ZNAME_3]"
             name_n(5) = content

        Case "[ZADRESSE:Y000001B-ZARZT3ANRE]"
            If Trim(content) = "03" Then
                FillProperties "Drs.", "anrede1"
                anrede(5) = "Sehr geehrte Kollegen "
            End If
            If Trim(content) = "02" Then
                FillProperties "Frau", "anrede5"
                anrede(5) = "Sehr geehrte Frau Kollegin "
            ElseIf Trim(content) = "01" Then
                FillProperties "Herr", "anrede5"
                anrede(5) = "Sehr geehrter Herr Kollege "
            End If
 
        Case "[ZADRESSE:Y000001B-ZARZT3NAME]"
            FillProperties content, "name5"
            name_k(5) = content
            If Left$(content, 4) = "Herr" Then
               anrede(5) = "Sehr geehrter Herr Kollege "
            End If
            If Left$(content, 4) = "Frau" Then
               anrede(5) = "Sehr geehrte Frau Kollegin "
            End If
            
        Case "[ZADRESSE:Y000001B-ZNAME2_3]"
            FillProperties content, "Firma5"
            
        Case "[ZADRESSE:Y000001B-ZARZT3STRA]"
            FillProperties content, "strass5"
         
        Case "[ZADRESSE:Y000001B-ZARZT3PLZ]"
            FillProperties content, "plz5"
        
        Case "[ZADRESSE:Y000001B-ZARZT3ORT]"
            FillProperties content, "ort5"
'-------------------------------------------------------
'hier beginnt Arzt4
'-------------------------------------------------------
         Case "[ZADRESSE:Y000001B-ZNAME_4]"
             name_n(6) = content

          Case "[ZADRESSE:Y000001B-ZARZT4ANRE]"
            If Trim(content) = "03" Then
                FillProperties "Drs.", "anrede1"
                anrede(6) = "Sehr geehrte Kollegen "
            End If
            If Trim(content) = "02" Then
                FillProperties "Frau", "anrede6"
                anrede(6) = "Sehr geehrte Frau Kollegin "
            ElseIf Trim(content) = "01" Then
                FillProperties "Herr", "anrede6"
                anrede(6) = "Sehr geehrter Herr Kollege "
            End If
        Case "[ZADRESSE:Y000001B-ZARZT4NAME]"
            FillProperties content, "name6"
            name_k(6) = content
            If Left$(content, 4) = "Herr" Then
               anrede(6) = "Sehr geehrter Herr Kollege "
            End If
            If Left$(content, 4) = "Frau" Then
               anrede(6) = "Sehr geehrte Frau Kollegin "
            End If
            
        Case "[ZADRESSE:Y000001B-ZNAME2_4]"
            FillProperties content, "Firma6"
        
        Case "[ZADRESSE:Y000001B-ZARZT4STRA]"
            FillProperties content, "strass6"
         
        Case "[ZADRESSE:Y000001B-ZARZT4PLZ]"
            FillProperties content, "plz6"
        
        Case "[ZADRESSE:Y000001B-ZARZT4ORT]"
            FillProperties content, "ort6"
'-------------------------------------------------------
'hier beginnt Arzt5
'-------------------------------------------------------
        Case "[ZADRESSE:Y000001B-ZNAME_5]"
             name_n(7) = content

         Case "[ZADRESSE:Y000001B-ZARZT5ANRE]"
            If Trim(content) = "03" Then
                FillProperties "Drs.", "anrede1"
                anrede(7) = "Sehr geehrte Kollegen "
            End If
            If Trim(content) = "02" Then
                FillProperties "Frau", "anrede7"
                anrede(7) = "Sehr geehrte Frau Kollegin "
            ElseIf Trim(content) = "01" Then
                FillProperties "Herr", "anrede7"
                anrede(7) = "Sehr geehrter Herr Kollege "
            End If
        Case "[ZADRESSE:Y000001B-ZARZT5NAME]"
            FillProperties content, "name7"
            name_k(7) = content
            If Left$(content, 4) = "Herr" Then
               anrede(7) = "Sehr geehrter Herr Kollege "
            End If
            If Left$(content, 4) = "Frau" Then
               anrede(7) = "Sehr geehrte Frau Kollegin "
            End If
            
        Case "[ZADRESSE:Y000001B-ZNAME2_5]"
            FillProperties content, "Firma7"
            
        Case "[ZADRESSE:Y000001B-ZARZT5STRA]"
            FillProperties content, "strass7"
         
        Case "[ZADRESSE:Y000001B-ZARZT5PLZ]"
            FillProperties content, "plz7"
        
        Case "[ZADRESSE:Y000001B-ZARZT5ORT]"
            FillProperties content, "ort7"
    End Select
  End If
 Next
  
'----------------------------------------------------------------------
' Jetzt wird die Anrede Sehr geehrter Herr Kollege ... zusammengesetzt
'----------------------------------------------------------------------
 For i = 1 To 6
     If RTrim(name_n(i)) = "" Then
        name_n(i) = name_k(i)
     End If
     If Trim$(name_n(i)) = "Prof. Dr. med. Volkmar Falk" Or Trim$(name_n(i)) = "Falk" Then
        anrede(i) = "Sehr geehrter Herr Professor Falk"
     Else
        anrede(i) = anrede(i) & name_n(i)
     End If
 Next i
 
 ' Position 7: Reha-Klinik, Name des Chefarztes steht in der ersten Zeile
 If RTrim(name_n(7)) = "" Then
    name_n(7) = name_k(7)
 End If
 anrede(7) = anrede(7) & name_k(7)
 
 FillProperties anrede(1), "ansprache1"
 FillProperties anrede(2), "ansprache2"
 FillProperties anrede(3), "ansprache3"
 FillProperties anrede(4), "ansprache4"
 FillProperties anrede(5), "ansprache5"
 FillProperties anrede(6), "ansprache6"
 FillProperties anrede(7), "ansprache7"
End Sub

Public Sub DoMedikation(gcolMedikation As Collection, medis As Collection)
Dim y As Long
Dim s As String
Dim itm As Collection, is_empty As Boolean
Dim MEDITXT As String
Dim ABENDS As Double, MITTAGS As Double, MTABDOSIS As Double
Dim MTABEINHE, MTABPRAEP, MTABZEIT
Dim MTAPPL, NACHTS As Double
    
    ActiveDocument.Bookmarks("MEDIKAMENT").Select
    set_format
    If gcolMedikation.Count = 0 Then
        Selection.Tables(1).Select
        Selection.Tables(1).Delete
        Exit Sub
    Else
        getDocVar gcolMedikation, "ZMEDITXT<L>", MEDITXT, is_empty, False
        Dim i As Long, j As Long, zeile As Word.row, amount As Double
        i = 2
        With Selection.Tables(1)
            For Each itm In medis
                j = 0
                If i = 2 Then
                    Set zeile = .Rows(2)
                Else
                    Set zeile = .Rows.Add
                End If
                j = j + 1
                getTabVar itm, "ZMTABPRAEP", MTABPRAEP, is_empty, False
                zeile.Cells(j).Range.text = MTABPRAEP
                j = j + 1
                getTabVar itm, "ZMTABDOSIS", MTABDOSIS, is_empty, True
                zeile.Cells(j).Range.text = MTABDOSIS
                j = j + 1
                getTabVar itm, "ZMTABEINHE", MTABEINHE, is_empty, False
                zeile.Cells(j).Range.text = MTABEINHE
                j = j + 1
                getTabVar itm, "ZMTAPPL", MTAPPL, is_empty, False
                zeile.Cells(j).Range.text = MTAPPL
                j = j + 1
                getTabVar itm, "ZMTABZEIT", MTABZEIT, is_empty, True
                zeile.Cells(j).Range.text = MTABZEIT
                j = j + 1
                getTabVar itm, "ZMITTAGS", MITTAGS, is_empty, True
                zeile.Cells(j).Range.text = MITTAGS
                j = j + 1
                getTabVar itm, "ZABENDS", ABENDS, is_empty, True
                zeile.Cells(j).Range.text = ABENDS
                j = j + 1
                getTabVar itm, "ZNACHTS", NACHTS, is_empty, True
                zeile.Cells(j).Range.text = NACHTS
                i = i + 1
            Next
        End With
    End If
    
    If MEDITXT > "" Then
        ActiveDocument.Bookmarks("MEDIKA_TEXT").Select
        set_format
        print_langtext MEDITXT
        Selection.TypeParagraph
    End If

End Sub

Public Sub DoTherapie(col As Collection, ab As Collection, am As Collection, mre As Collection) ' ab = Antibiotika, am = Antimyotika
    Dim tabrow As Collection, gefunden As Boolean, feld, i As Integer, is_empty As Boolean
    Dim datum, uhrzeit, INDIKATIO, VONDATUM, BISDATUM, WIRKSTOFF, grund
    
    getDocVar col, "ZUHRZEIT", uhrzeit, is_empty, False
    If Not is_empty Then uhrzeit = ISHMED_Uhrzeit(uhrzeit & "")
    getDocVar col, "ZDATUM", datum, is_empty, False
    If Not is_empty Then datum = ISHMED_Datum(datum & "")

    ActiveDocument.Bookmarks("Antibiotika").Select
    Selection.MoveLeft Count:=1
    set_format
    Selection.TypeParagraph
    Selection.Font.Underline = True
    Selection.TypeText "Antibiotika-/Antimykotika-Therapie"
    Selection.Font.Underline = False
    Selection.TypeParagraph
    Selection.TypeText "Daten zuletzt aktualisiert am " & datum & " um " & uhrzeit
    Selection.TypeParagraph
    
    'anlegenTabelle mytable As Word.table, breiten(), ueberschriften(), mit_ueberschrift As Boolean
    Dim tb As Word.Table, breiten(), ueberschriften(), tr As Word.row
    breiten = Array(CentimetersToPoints(5), CentimetersToPoints(2.5), CentimetersToPoints(2.5), CentimetersToPoints(7))
    ueberschriften = Array("Wirkstoff", "Erste Gabe", "Letzte Gabe", "Diagnose/Indikation")
    anlegenTabelle tb, breiten, ueberschriften, True
    '... Tabelle fuellen
    i = 0
    If Not ab Is Nothing Then
        For Each tabrow In ab
            i = i + 1
            getTabVar tabrow, "ZVONDATUM", VONDATUM, is_empty, False
            If Not is_empty Then VONDATUM = ISHMED_Datum(VONDATUM & "")
            getTabVar tabrow, "ZBISDATUM", BISDATUM, is_empty, False
            If Not is_empty Then BISDATUM = ISHMED_Datum(BISDATUM & "")
            getTabVar tabrow, "ZWIRKSTOFF", WIRKSTOFF, is_empty, False
            getAnyTabVar tabrow, Array("ZINDIKATIO", "ZINDIKATIE<L>"), grund, is_empty, False
            If i = 1 Then
                Set tr = tb.Rows(2)
            Else
                Set tr = tb.Rows.Add
            End If
            tr.Cells(1).Range.text = WIRKSTOFF
            tr.Cells(2).Range.text = VONDATUM
            tr.Cells(3).Range.text = BISDATUM
            tr.Cells(4).Range.text = grund
        Next
    End If
    If Not am Is Nothing Then
        For Each tabrow In am
            i = i + 1
            getTabVar tabrow, "ZVONDATUM", VONDATUM, is_empty, False
            If Not is_empty Then VONDATUM = ISHMED_Datum(VONDATUM & "")
            getTabVar tabrow, "ZBISDATUM", BISDATUM, is_empty, False
            If Not is_empty Then BISDATUM = ISHMED_Datum(BISDATUM & "")
            getTabVar tabrow, "ZWIRKSTOFF", WIRKSTOFF, is_empty, False
            getAnyTabVar tabrow, Array("ZINDIKATIO", "ZINDIKATIE<L>"), grund, is_empty, False
            If i = 1 Then
                Set tr = tb.Rows(2)
            Else
                Set tr = tb.Rows.Add
            End If
            tr.Cells(1).Range.text = WIRKSTOFF
            tr.Cells(2).Range.text = VONDATUM
            tr.Cells(3).Range.text = BISDATUM
            tr.Cells(4).Range.text = grund
        Next
    End If
    abschliessenTabelle tb, True
    
    If Not mre Is Nothing Then
        ActiveDocument.Bookmarks("Antibiotika").Select
        Selection.MoveLeft Count:=1
        set_format
        Selection.TypeParagraph
        Selection.Font.Underline = True
        Selection.TypeText "MRE-/MRSA-Sanierung"
        Selection.Font.Underline = False
        Selection.TypeParagraph
        breiten = Array(CentimetersToPoints(11), CentimetersToPoints(3), CentimetersToPoints(3))
        ueberschriften = Array("Maßnahme", "Erste Anw.", "Letzte Anw.")
        anlegenTabelle tb, breiten, ueberschriften, True
        i = 0
        For Each tabrow In mre
            i = i + 1
            getTabVar tabrow, "ZVONDATUM", VONDATUM, is_empty, False
            If Not is_empty Then VONDATUM = ISHMED_Datum(VONDATUM & "")
            getTabVar tabrow, "ZBISDATUM", BISDATUM, is_empty, False
            If Not is_empty Then BISDATUM = ISHMED_Datum(BISDATUM & "")
            getTabVar tabrow, "ZWIRKSTOFF", WIRKSTOFF, is_empty, False
            If i = 1 Then
                Set tr = tb.Rows(2)
            Else
                Set tr = tb.Rows.Add
            End If
            tr.Cells(1).Range.text = WIRKSTOFF
            tr.Cells(2).Range.text = VONDATUM
            tr.Cells(3).Range.text = BISDATUM
        Next
        abschliessenTabelle tb, True
    End If
End Sub

Public Sub DoMedikationNeu(kopf As Collection, medis As Collection)
    Dim y As Long
    Dim s As String
    Dim itm
    Dim MEDITXT, datum As String
    Dim is_empty As Boolean
    
    If Not kopf Is Nothing Then
        getDocVar kopf, "ZENTMEDTXT<L>", MEDITXT, is_empty, False
        If getDocVar(kopf, "ZDATUM", datum, is_empty, False) Then
            datum = ISHMED_Datum(datum & "")
        End If
        
        If datum > "" Then
            ActiveDocument.Bookmarks("StandMedikation").Select
            Selection.Font.Underline = wdUnderlineNone
            Selection.TypeText "(Stand vom " & datum & ")"
            Selection.TypeParagraph
        End If
        
        If MEDITXT > "" Then
            set_format
            ActiveDocument.Bookmarks("MEDIKA_TEXT").Select
            Selection.MoveLeft Count:=1
            Selection.TypeParagraph
            print_langtext MEDITXT
            Selection.TypeParagraph
        End If
    End If
    
    ActiveDocument.Bookmarks("MEDIKAMENT").Select

    If medis Is Nothing Then
        is_empty = True
    Else
        is_empty = medis.Count = 0
    End If
    If is_empty Then
        Selection.Tables(1).Select
        Selection.Tables(1).Delete
    Else
        Dim i As Long, feld, zeile As Word.row, wert As String, betrag As Variant
        Dim tabrow As Collection ', font As Word.font
        i = 2
        With Selection.Tables(1)
            Dim grund As String
'            .Select
'            Set font = Selection.font
'            Selection.font.size = 8
'            Selection.Collapse
            For i = 2 To medis.Count
                .Rows.Add
            Next
            i = 1
            For Each tabrow In medis
                ' Beginnt unter der Ueberschrift auf Zeile 2
                Dim kein_morgen As Boolean, kein_mittag As Boolean, kein_abend As Boolean, kein_nacht As Boolean
                Dim dosis_morgen As Double, dosis_mittag As Double, dosis_abend As Double, dosis_nacht As Double
                i = i + 1
                Set zeile = .Rows(i)
                getTabVar tabrow, "ZEMTPRAEP", wert, is_empty, False
                If Not is_empty Then zeile.Cells(1).Range.text = wert
                getTabVar tabrow, "ZEMTDOSIS", betrag, is_empty, True
                If Not is_empty Then
                    zeile.Cells(2).Range.text = betrag
                End If
                getTabVar tabrow, "ZEMTDOEINH", wert, is_empty, False
                If Not is_empty Then zeile.Cells(3).Range.text = wert
                getTabVar tabrow, "ZEMTAPPL", wert, is_empty, False
                If Not is_empty Then zeile.Cells(4).Range.text = wert
                getTabVar tabrow, "ZEMTMEINH", wert, is_empty, False
                If Not is_empty Then zeile.Cells(5).Range.text = wert
                getTabVar tabrow, "ZEMTMORGEN", dosis_morgen, kein_morgen, True
                getTabVar tabrow, "ZEMTMITTAG", dosis_mittag, kein_mittag, True
                getTabVar tabrow, "ZEMTABEND", dosis_abend, kein_abend, True
                getTabVar tabrow, "ZEMTNACHT", dosis_nacht, kein_nacht, True
                getTabVar tabrow, "ZEMTGRUND", grund, is_empty, False
                If (Not is_empty) And kein_morgen And kein_mittag And kein_abend And kein_nacht Then
                    Dim l_range As Word.Range
                    Set l_range = ActiveDocument.Range(.Cell(i, 6).Range.Start, .Cell(i, 9).Range.End)
                    l_range.Cells.Merge
                    .Cell(i, 6).Range.text = grund
                Else
                    If kein_morgen Then
                        zeile.Cells(6).Range.text = "-"
                    Else
                        zeile.Cells(6).Range.text = dosis_morgen
                    End If
                    If kein_mittag Then
                        zeile.Cells(7).Range.text = "-"
                    Else
                        zeile.Cells(7).Range.text = dosis_mittag
                    End If
                    If kein_abend Then
                        zeile.Cells(8).Range.text = "-"
                    Else
                        zeile.Cells(8).Range.text = dosis_abend
                    End If
                    If kein_nacht Then
                        zeile.Cells(9).Range.text = "-"
                    Else
                        zeile.Cells(9).Range.text = dosis_nacht
                    End If
                End If
           Next
        End With
        ActiveDocument.Bookmarks("MEDIKA_TEXT").Select
        Selection.MoveLeft Count:=1
        Selection.TypeParagraph
        Selection.TypeText "Oder wirkstoffgleiche Medikamente"
        Selection.TypeParagraph
        Selection.TypeParagraph
'        Selection.font.size = font.size
    End If

End Sub

' Transfusionen werden nur in der Summe uebernommen, nicht einzeln die Tabelle abgeklappert
Public Sub DoBlutprod(col As Collection, antikoerper As Collection, summen As Collection)
    Dim sumTrans As Long      'Summe der Transfusionen
    Dim is_empty As Boolean
    Dim i As Long
    Dim idx As Long
    Dim tabrow As Collection
    Dim d As String              'Datum, enthält jüngstes Datum der Antikörperstatusdaten
    Dim timeLastestEntry As String 'Uhrzeit, enthält jüngste Uhrzeit der Antikörperstatusdaten
    Dim p1, p2 As Long        'Position eines Leerzeichens in einem String
    Dim s As String              'Hilfstext, der den Wert eines Datenfeldes temp. aufnimmt
    Dim summen_vorhanden As Boolean
    Dim antikoerper_vorhanden As Boolean
    Dim BLUTGRUPP, CDEFORMEL, EBSUMME, EFFPSUM, EKLISUMME, EKLSUMME, FFPISUMME, FFPSUMME, _
        FFPVSUMME, RHESUSFKT, TKGPISUMM, TKGPSUMME, TKLISUMME, TKLSUMME, befund

    ActiveDocument.Bookmarks("Blutprodukte").Select
    set_format
    getDocVar col, "ZBLUTGRUPP", BLUTGRUPP, is_empty, False
    getDocVar col, "ZCDEFORMEL", CDEFORMEL, is_empty, False
    getDocVar col, "ZEBSUMME", EBSUMME, is_empty, False
    getDocVar col, "ZEFFPSUM", EFFPSUM, is_empty, False
    getDocVar col, "ZEKLISUMME", EKLISUMME, is_empty, False
    getDocVar col, "ZEKLSUMME", EKLSUMME, is_empty, False
    getDocVar col, "ZFFPISUMME", FFPISUMME, is_empty, False
    getDocVar col, "ZFFPSUMME", FFPSUMME, is_empty, False
    getDocVar col, "ZFFPVSUMME", FFPVSUMME, is_empty, False
    getDocVar col, "ZRHESUSFKT", RHESUSFKT, is_empty, False
    getDocVar col, "ZTKGPISUMM", TKGPISUMM, is_empty, False
    getDocVar col, "ZTKGPSUMME", TKGPSUMME, is_empty, False
    getDocVar col, "ZTKLISUMME", TKLISUMME, is_empty, False
    getDocVar col, "ZTKLSUMME", TKLSUMME, is_empty, False
    getDocVar col, "ZEBEFUND<L>", befund, is_empty, False
    
    summen_vorhanden = False
    If Not summen Is Nothing Then
        summen_vorhanden = summen.Count > 0
    End If
    
    antikoerper_vorhanden = False
    If Not antikoerper Is Nothing Then
        antikoerper_vorhanden = antikoerper.Count > 0
    End If

    ' Überschrift wird nicht ausgegeben, da dieser Text fest in Vorlage vorhanden ist.
    
    ' vor Ausgabe der Blutparameter die KELL-Formel "stutzen", d.h. die Werte für Blutgruppe
    ' und Rhesus-Faktor aus dem Datenelement entfernen. Diese beiden Werte werden durch das
    ' DHZB ins SAP-System transferiert und befinden sich am Anfang des Datenfeldes .ZBPCDEFORMEL,
    ' getrennt durch Leerzeichen. Es wird nachfolgend zum zweiten Leerzeichen gesprungen und nur
    ' der anschließende Text verwendet:
    p1 = Len(CDEFORMEL)
    
    If p1 > 0 Then
        s = CDEFORMEL                    'Verwendet den kompletten Wert aus KELL-Formel
    Else
        s = "keine Angabe"
    End If
    
    ' Ausgabe der Blutparameter (dabei wird die zuvor ggf. "gestutzte" KELL-Formel am Ende verwendet):
    Selection.TypeText "Blutgruppe: " & BLUTGRUPP & " Rhesus-Faktor: " & RHESUSFKT & vbNewLine & "CDE Formel mit Kell: " & s
    Selection.TypeParagraph
    
    If antikoerper_vorhanden Then
        ' Ermitteln des jüngsten Antikörperstatus-Eintrags:
        Dim t As String
        d = ""
   
        Dim datum As String, text As String, feld, zeile As Object, time As String
        
        For Each zeile In antikoerper
            getTabVar zeile, "ZAKSTDAT", datum, is_empty, False
            getTabVar zeile, "ZAKSTUHR", time, is_empty, False
            getTabVar zeile, "ZEAKSTTXT<L>", text, is_empty, False
            If text > "" And ((d = "") Or (d < datum)) Then
                t = text
                d = datum
                timeLastestEntry = time
            Else
                If text > "" And ((d = "") Or (d = datum)) Then
                    If timeLastestEntry = "" Or timeLastestEntry < time Then
                        t = text
                        d = datum
                        timeLastestEntry = time
                    End If
                End If
            End If
        Next
        
        ' Ausgabe des "jüngsten" Antikörperstatus-Eintrags:
        If d > "" Then
            Selection.TypeText "Letzter Antikörpersuchtest vom " & _
                ISHMED_Datum(d) & ": " & t
            Selection.TypeParagraph
        End If
        
        If befund > "" Then
            Selection.TypeParagraph
            Selection.TypeText befund
            Selection.TypeParagraph
        End If
    End If ' Antikoerper_vorhanden
    
    ' Ermittle Summe der Transfusionen:
    sumTrans = CLng(EBSUMME) + CLng(EFFPSUM) + CLng(EKLISUMME) + CLng(EKLSUMME) + CLng(FFPISUMME) + CLng(FFPSUMME _
               ) + CLng(FFPVSUMME) + CLng(TKGPISUMM) + CLng(TKGPSUMME) + CLng(TKLISUMME) + CLng(TKLSUMME)
    
    ' Ausgabe der Transfusionen oder - im Falle keiner Transfusion - eines Standardtextes:
    Selection.TypeParagraph
    Selection.TypeText "Transfusionen:"
    If sumTrans > 0 Then
        If EBSUMME > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(EBSUMME, "0") & " EB (Erythrozyten Eigenblut)"
        End If
        If EFFPSUM > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(EFFPSUM, "0") & " EFFP (FFP Eigenblut)"
        End If
        If EKLSUMME > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(EKLSUMME, "0") & " EK/L (Erythrozytenkonz. Leukozytendepletiert)"
        End If
        If EKLISUMME > 0 Then
            Selection.TypeText Format(EKLISUMME, "0") & " EK/LI (Erythrozytenkonz. Leukozytendepletiert und bestrahlt)"
            Selection.TypeParagraph
        End If
        If FFPSUMME > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(FFPSUMME, "0") & " FFP (Gefrorenes Frischplasma)"
        End If
        If FFPISUMME > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(FFPISUMME, "0") & " FFP/I (Gefrorenes Frischplasma bestrahlt)"
        End If
        If FFPVSUMME > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(FFPVSUMME, "0") & " FFP/V (Gefrorenes Frischplasma virusinaktiviert (SD))"
        End If
        If TKGPSUMME > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(TKGPSUMME, "0") & " TK/GP (Thrombozytenkonz. gepoolt)"
        End If
        If TKGPISUMM > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(TKGPISUMM, "0") & " TK/GPI (Thrombozytenkonz. gepoolt bestrahlt)"
        End If
        If TKLSUMME > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(TKLSUMME, "0") & " TK/L (Thrombozytenkonz. Leukozytendepletiert)"
        End If
        If TKLISUMME > 0 Then
            Selection.TypeParagraph
            Selection.TypeText Format(TKLISUMME, "0") & " TK/LI (Thrombozytenkonz. Leukozytendepletiert und bestrahlt)"
        End If
    ElseIf summen_vorhanden Then
        ' Neue Darstellung in Tabelle
        Dim tbl As Word.Table, row As Word.row, r As Integer, breiten(), ueberschriften()
        breiten = Array(InchesToPoints(1), InchesToPoints(1), InchesToPoints(4))
        'Eintragen der Spaltenköpfe für die nachfolgenden Tabellenwerte:
        ueberschriften = Array("Komponente", "Gegeben", "Beschreibung")
        anlegenTabelle tbl, breiten, ueberschriften, True
        r = 2
        For Each tabrow In summen
            Dim komp As String, beschr As String, anz As Integer
            getTabVar tabrow, "ZTRKOMP", komp, is_empty, False
            getTabVar tabrow, "ZTRSUMME", anz, is_empty, False
            getTabVar tabrow, "ZTRTEXT", beschr, is_empty, False
            If r = 2 Then
                Set row = tbl.Rows(r)
            Else
                Set row = tbl.Rows.Add
            End If
            row.Cells(1).Range.text = komp
            row.Cells(2).Range.text = anz
            row.Cells(3).Range.text = beschr
            r = r + 1
        Next
        abschliessenTabelle tbl, True
    Else
        Selection.TypeText " Es wurden im Paulinenkrankenhaus keine Transfusionen durchgeführt."
        Selection.TypeParagraph
        ' Falls keine Transfusionen vorhanden sind, verlassse diese Routine umgehend:
        Exit Sub
    End If

End Sub

Public Sub DoGehtest(gcolGehTest As Collection, tb As Collection)
    'Dim gt As clsGehtest
    'Dim gt6m As clsGehtest6Min
    Dim i As Long, itm As Variant, zeile As Collection, is_empty As Boolean
    Dim idx As Long
    Dim d As String              'Datum, enthält jüngstes Datum der Antikörperstatusdaten
    Dim p1, p2 As Long        'Position eines Leerzeichens in einem String
    Dim s As String              'Hilfstext, der den Wert eines Datenfeldes temp. aufnimmt
    Dim tabs
    Dim leere_tabelle As Boolean, leerer_befund As Boolean
    
    leere_tabelle = True
    If Not tb Is Nothing Then
        leere_tabelle = tb.Count = 0
    End If
    
    Dim ZBEFTEXT, ZDATUM As String, ZDOBUTAMIN, ZDOPAMIN, ZDOSISDOBU, ZDOSISDOPA, ZGEWICHT, ZHFNACH, ZHFVOR, _
        ZKOMMENTAR, ZRRNACH, ZRRVOR, ZSAOPERNAC, ZSAOPERVOR, ZSAOZENNAC, ZSAOZENVOR, ZSTRECKE, ZUHRZEIT
        
    getAnyDocVar gcolGehTest, Array("ZBEFTEXT<L>", "ZEBEFUND<L>"), ZBEFTEXT, leerer_befund, False
    
    If leere_tabelle And leerer_befund Then Exit Sub
    
    ' Achtung: mehrere Gehteste in einem Dokument, aber nur ein Befundtext, siehe unten,
    ' todo!
    
    'Textmarke für die Ausgabe anspringen und den Cursor positionieren
    ActiveDocument.Bookmarks("Gehtest").Select
    set_format
    
    'Überschrift ausgeben:
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="6-Minuten-Gehtest"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph

    If Not leere_tabelle Then
        tabs = Array(0.8, 2.3, 3.2, 4.1, 5, 6, 7, 8.3, 9.5, 10.7, 12.3, 13.8, 15, 16)
        'bei vorgenommenen Gehtests zunächst die Kopfzeile für die Datenspalten ausgeben:
        'Tabulatoren für die Ausgabe der Gehtests setzen
        Dim TS
        Selection.Font.size = 8
        Selection.ParagraphFormat.TabStops.ClearAll
        For i = 0 To UBound(tabs) - 1
            Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(i) + ((tabs(i + 1) - tabs(i)) / 2)) _
                  , Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
        Next
        
        Selection.TypeText "Gew." & vbTab & "Dopamin" & vbTab & "Dos.*" & vbTab & "Dobu" & vbTab & "Dos.*" & vbTab & "SAO2p" & vbTab & "SAO2z" & vbTab & _
                           "RR vor." & vbTab & "HF vor." & vbTab & "Strecke" & vbTab & "RR nach." & vbTab & "HF nach." & vbTab & "SAO2p" & vbTab & "SAO2z"
        Selection.TypeParagraph
        Selection.TypeText " (kg)" & vbTab & "(ml/h)" & vbTab & vbTab & "(ml/h)" & vbTab & vbTab & "(%)" & vbTab & "(%)" & vbTab & vbTab & vbTab & "(m)" & _
                           vbTab & vbTab & vbTab & "(%)" & vbTab & "(%)"
                           
        Selection.TypeParagraph
        Selection.ParagraphFormat.TabStops.ClearAll
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(0) + 0.5) _
              , Alignment:=wdAlignTabDecimal, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(1) + 0.5) _
              , Alignment:=wdAlignTabDecimal, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(2) + 0.5) _
              , Alignment:=wdAlignTabDecimal, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(3) + 0.5) _
              , Alignment:=wdAlignTabDecimal, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(5) - 0.2) _
              , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(6) - 0.2) _
              , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(6) + 0.5) _
              , Alignment:=wdAlignTabDecimal, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(8) - 0.2) _
              , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(9) - 0.2) _
              , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(10) - 0.2) _
              , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(11) - 0.2) _
              , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(12) - 0.2) _
              , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tabs(13) - 0.2) _
              , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    
        Selection.Font.size = 10
        idx = 0
        'Bei vorgenommenen Gehtests gib pro Gehtest eine tabellenartige Zeile aus:
        For Each zeile In tb
            idx = idx + 1
            getTabVar zeile, "ZDATUM", ZDATUM, is_empty, False
            getTabVar zeile, "ZDOBUTAMIN", ZDOBUTAMIN, is_empty, False
            getTabVar zeile, "ZDOPAMIN", ZDOPAMIN, is_empty, False
            getTabVar zeile, "ZDOSISDOBU", ZDOSISDOBU, is_empty, False
            getTabVar zeile, "ZDOSISDOPA", ZDOSISDOPA, is_empty, False
            getTabVar zeile, "ZGEWICHT", ZGEWICHT, is_empty, False
            getTabVar zeile, "ZHFNACH", ZHFNACH, is_empty, False
            getTabVar zeile, "ZHFVOR", ZHFVOR, is_empty, False
            getTabVar zeile, "ZKOMMENTAR", ZKOMMENTAR, is_empty, False
            getTabVar zeile, "ZRRNACH", ZRRNACH, is_empty, False
            getTabVar zeile, "ZRRVOR", ZRRVOR, is_empty, False
            getTabVar zeile, "ZSAOPERNAC", ZSAOPERNAC, is_empty, False
            getTabVar zeile, "ZSAOPERVOR", ZSAOPERVOR, is_empty, False
            getTabVar zeile, "ZSAOZENNAC", ZSAOZENNAC, is_empty, False
            getTabVar zeile, "ZSAOZENVOR", ZSAOZENVOR, is_empty, False
            getTabVar zeile, "ZSTRECKE", ZSTRECKE, is_empty, False
            getTabVar zeile, "ZUHRZEIT", ZUHRZEIT, is_empty, False
            ' das Datum des Test ausgeben:
            Selection.Font.Italic = True
            Selection.TypeText Format(idx, "0") & ". Test am: " & ISHMED_Datum(ZDATUM)
            Selection.Font.Italic = False
            Selection.TypeParagraph
            ' nun die einzelnen Datenfelder ausgeben:
            If ZGEWICHT <> "" Then
              Selection.TypeText getDecimalFormat(ZGEWICHT)
            End If
            Selection.TypeText vbTab
            If ZDOPAMIN <> "" Then
              Selection.TypeText getDecimalFormat(ZDOPAMIN)
            End If
            Selection.TypeText vbTab
            If ZDOSISDOPA <> "" Then
              Selection.TypeText getDecimalFormat(ZDOSISDOPA)
            End If
            Selection.TypeText vbTab
            If ZDOBUTAMIN <> "" Then
              Selection.TypeText getDecimalFormat(ZDOBUTAMIN)
            End If
            Selection.TypeText vbTab
            If ZDOSISDOBU <> "" Then
              Selection.TypeText getDecimalFormat(ZDOSISDOBU)
            End If
            Selection.TypeText vbTab
            If ZSAOPERVOR > 0 Then
              Selection.TypeText Format(ZSAOPERVOR, "0")
            End If
            Selection.TypeText vbTab
            If ZSAOZENVOR > 0 Then
              Selection.TypeText Format(ZSAOZENVOR, "0")
            End If
            Selection.TypeText vbTab
            If ZRRVOR <> "" Then
              Selection.TypeText ZRRVOR
            End If
            Selection.TypeText vbTab
            If ZHFVOR > 0 Then
              Selection.TypeText Format(ZHFVOR, "0")
            End If
            Selection.TypeText vbTab
            If ZSTRECKE > 0 Then
              Selection.TypeText Format(ZSTRECKE, "0")
            End If
            Selection.TypeText vbTab
            If ZRRNACH <> "" Then
              Selection.TypeText ZRRNACH
            End If
            Selection.TypeText vbTab
            If ZHFNACH > 0 Then
              Selection.TypeText Format(ZHFNACH, "0")
            End If
            Selection.TypeText vbTab
            If ZSAOPERNAC > 0 Then
              Selection.TypeText Format(ZSAOPERNAC, "0")
            End If
            Selection.TypeText vbTab
            If ZSAOZENNAC > 0 Then
              Selection.TypeText Format(ZSAOZENNAC, "0")
            End If
            ' abschließend in der nächsten Zeile den Kommentar zum Gehtest ausgeben:
            Selection.EndKey UNIT:=wdLine
            Selection.TypeParagraph
            If ZKOMMENTAR <> "" Then
              Selection.TypeText "Kommentar: " & ZKOMMENTAR
              Selection.TypeParagraph
            End If
            ' nach jedem einzeln Gehtest eine Leerzeile einfügen:
            Selection.TypeParagraph
        Next
        ' Ausgabe der Fußnote zur Dosis:
        Selection.Font.size = 8
        Selection.TypeText "* (µg/kg/Min.)"
        Selection.Font.size = 10
        Selection.TypeParagraph
    End If
    
    If Not leerer_befund Then
        Selection.TypeParagraph
        ' Ausgabe des Befundtextes:
        Selection.TypeText "Befund: " & ZBEFTEXT
    End If

    Selection.TypeParagraph
    
End Sub 'DoGehtest

Public Sub DoErregerDetails(col As Collection)

    DoErreger
    
    ' Erreger-Details
    Dim itm As Collection, tz As Collection
    Dim zeile As Long, spalte As Long, is_empty As Boolean
    Dim Abnahmetag, Medium, MaterialInfo, Erregertext, Keimzahl, Resistenzgruppe, _
        Meldepflicht, Penicillin, Oxacillin, Amicillin, Ampicillin_Sulbactam, _
        Piperacillin, Piperacillin_Tazobactam, Cefazolin_Gr1, Cefaclor_Oral_Gr1, _
        Cefalexin_Oral_Gr1, Cefuroxim_Oral_PE_Gr2, Cefotiam_Gr2, Cefixim_Oral_Gr3, _
        Cefotaxim_Gr3a, Ceftriaxon_Gr3a, Ceftazidim_Gr3b, Imipenem, Meropenem, _
        Ofloxacin, Ciprofloxacin, Moxifloxacin, Gentamycin, Gentamycin_HR, Tobramycin, _
        Amikacin, NEOMYCIN, TMP_SMZ, Doxycyclin, Erythromycin, Clarithromycin, Clindamycin, _
        Vancomycin, Teicoplanin, Rifampicin, LINEZOLID, Synercid, Fosfomycin, Nitrofurantoin, _
        Metronidazol, Fluconazol, Itraconazol, Amphotericin_B, Flucytosin, Voriconazol, _
        Caspofungin, Tetracyclin, Piperacillin_Sulbactam, COLISTIN, Tigecyclin, Mezlocillin, _
        Mezlocillin_Sulbactam, RESERR, MERLOC
    Dim tbl As Word.Table, rg As Word.Range

    Selection.GoTo What:=wdGoToBookmark, name:="MibiErregerHinweis"
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="Mikrobiologie (Erreger)"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    Selection.TypeText text:="Siehe Anlage."
    Selection.TypeParagraph
    Selection.GoTo What:=wdGoToBookmark, name:="MibiErregerTabelle1"
    Selection.TypeText "Mikrobiologie-Befund (Erreger)"
    
    ' Nach den Unterschriften kommt nur noch die Mibi-Tabelle
    ActiveDocument.Bookmarks("Unterschrift").Select
    ' Die naechste Tabelle finden
    Selection.GoTo (wdGoToTable)

    
    ' In diesem Range ist genau eine Tabelle, naemlich die gesuchte!
    Set tbl = Selection.Range.Tables(1)
    ' Fuer jede Tabellenzeile, starten bei der zweiten Zeile (erste Zeile hat die Ueberschriften)
    zeile = 1
    With tbl
        For Each tz In col
            ' Zeilenzaehler erhoehen, beginnt bei 1
            zeile = zeile + 1
            ' einen Zellensprung vornehmen, wenn es sich nicht um den ersten
            ' Datensatz handelt, um eine neue Tabellenzeile zu erzeugen
            If zeile > 2 Then .Rows.Add
            
            getTabVar tz, "ZUNTERSDAT", Abnahmetag, is_empty, False
            getTabVar tz, "ZMEDIUM", Medium, is_empty, False
            getTabVar tz, "ZEBMATINFO", MaterialInfo, is_empty, False
            getTabVar tz, "ZRESERRTXT", Erregertext, is_empty, False
            getTabVar tz, "ZKEIMZAHL", Keimzahl, is_empty, False
            getTabVar tz, "ZRESGRUPPE", Resistenzgruppe, is_empty, False
            getTabVar tz, "ZSPEZIELL", Meldepflicht, is_empty, False
            getTabVar tz, "ZPENICILLI", Penicillin, is_empty, False
            getTabVar tz, "ZOXA", Oxacillin, is_empty, False
            getTabVar tz, "ZAMPIC", Amicillin, is_empty, False
            getTabVar tz, "ZAMPI", Ampicillin_Sulbactam, is_empty, False
            getTabVar tz, "ZPIPER", Piperacillin, is_empty, False
            getTabVar tz, "ZTAB", Piperacillin_Tazobactam, is_empty, False
            getTabVar tz, "ZCEFAZOLIN", Cefazolin_Gr1, is_empty, False
            getTabVar tz, "ZCEFACLOR", Cefaclor_Oral_Gr1, is_empty, False
            getTabVar tz, "ZCEFALEXIN", Cefalexin_Oral_Gr1, is_empty, False
            getTabVar tz, "ZCEFU", Cefuroxim_Oral_PE_Gr2, is_empty, False
            getTabVar tz, "ZCEFO", Cefotiam_Gr2, is_empty, False
            getTabVar tz, "ZCEFIX", Cefixim_Oral_Gr3, is_empty, False
            getTabVar tz, "ZCTX", Cefotaxim_Gr3a, is_empty, False
            getTabVar tz, "ZCEFT", Ceftriaxon_Gr3a, is_empty, False
            getTabVar tz, "ZCTM", Ceftazidim_Gr3b, is_empty, False
            getTabVar tz, "ZIMP", Imipenem, is_empty, False
            getTabVar tz, "ZMER", Meropenem, is_empty, False
            getTabVar tz, "ZOFLO", Ofloxacin, is_empty, False
            getTabVar tz, "ZCIP", Ciprofloxacin, is_empty, False
            getTabVar tz, "ZMOXIFLOXA", Moxifloxacin, is_empty, False
            getTabVar tz, "ZGM", Gentamycin, is_empty, False
            getTabVar tz, "ZGMHOCH", Gentamycin_HR, is_empty, False
            getTabVar tz, "ZTOB", Tobramycin, is_empty, False
            getTabVar tz, "ZAN", Amikacin, is_empty, False
            getTabVar tz, "ZNEOMYCIN", NEOMYCIN, is_empty, False
            getTabVar tz, "ZTS", TMP_SMZ, is_empty, False
            getTabVar tz, "ZDOXYCYCLI", Doxycyclin, is_empty, False
            getTabVar tz, "ZERY", Erythromycin, is_empty, False
            getTabVar tz, "ZCLARITMYZ", Clarithromycin, is_empty, False
            getTabVar tz, "ZCLINDA", Clindamycin, is_empty, False
            getTabVar tz, "ZVAN", Vancomycin, is_empty, False
            getTabVar tz, "ZTEIC", Teicoplanin, is_empty, False
            getTabVar tz, "ZRIFAMI", Rifampicin, is_empty, False
            getTabVar tz, "ZLINEZOLID", LINEZOLID, is_empty, False
            getTabVar tz, "ZSYNERZID", Synercid, is_empty, False
            getTabVar tz, "ZFOS", Fosfomycin, is_empty, False
            getTabVar tz, "ZNITRO", Nitrofurantoin, is_empty, False
            getTabVar tz, "ZMETRO", Metronidazol, is_empty, False
            getTabVar tz, "ZFLU", Fluconazol, is_empty, False
            getTabVar tz, "ZITRACONAZ", Itraconazol, is_empty, False
            getTabVar tz, "ZAMPHOTERB", Amphotericin_B, is_empty, False
            getTabVar tz, "ZFLUCYTOSI", Flucytosin, is_empty, False
            getTabVar tz, "ZVORICONAZ", Voriconazol, is_empty, False
            getTabVar tz, "ZCASPOFUNG", Caspofungin, is_empty, False
            getTabVar tz, "ZTETRA", Tetracyclin, is_empty, False
            getTabVar tz, "ZPIPESULBA", Piperacillin_Sulbactam, is_empty, False
            getTabVar tz, "ZCOLISTIN", COLISTIN, is_empty, False
            getTabVar tz, "ZTIGECYCLI", Tigecyclin, is_empty, False
            getTabVar tz, "ZMEZLO", Mezlocillin, is_empty, False
            getTabVar tz, "ZMEZLOC", Mezlocillin_Sulbactam, is_empty, False
            getTabVar tz, "ZRESERR", RESERR, is_empty, False
            getTabVar tz, "ZMERLOC", MERLOC, is_empty, False
'            getAnyDocVar  tz, Array("ZMEBEFID", "ZMEEBMAT"), , is_empty,   False
           
            spalte = 0
            
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Format(ISHMED_Datum(Abnahmetag & ""), "DD.MM.YY")
            spalte = spalte + 1
            '.Cell(zeile, spalte).Range.text = Station
            spalte = spalte + 1
            .Cell(zeile, 3).Range.text = Medium & IIf(Len(Trim(MaterialInfo)) = 0, "", " (" & MaterialInfo & ")")
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Erregertext & IIf(Len(Trim(Keimzahl)) = 0, "", " (" & Keimzahl & ")")
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Resistenzgruppe
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Meldepflicht
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Penicillin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Oxacillin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Amicillin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Ampicillin_Sulbactam
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Piperacillin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Piperacillin_Tazobactam
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Cefazolin_Gr1
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Cefaclor_Oral_Gr1
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Cefalexin_Oral_Gr1
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Cefuroxim_Oral_PE_Gr2
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Cefotiam_Gr2
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Cefixim_Oral_Gr3
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Cefotaxim_Gr3a
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Ceftriaxon_Gr3a
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Ceftazidim_Gr3b
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Imipenem
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Meropenem
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Ciprofloxacin
            spalte = spalte + 1
            '.Cell(zeile, spalte).Range.text = Levofloxacin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Moxifloxacin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Gentamycin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Gentamycin_HR
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Tobramycin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Amikacin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = COLISTIN
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Erythromycin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Clarithromycin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Doxycyclin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Tetracyclin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Tigecyclin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Clindamycin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Vancomycin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Rifampicin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = LINEZOLID
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Fosfomycin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Nitrofurantoin
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = TMP_SMZ
            spalte = spalte + 1
            .Cell(zeile, spalte).Range.text = Metronidazol
        Next
    End With
End Sub

'' KD, 28.06.12, Ausgabe der Erregerdaten (Mikrobilogie)
Public Sub DoErreger()
' das Hinweis-Lesezeichen anspringen und dort einen Hinweis auf den Anhang einfügen
    If gcolErreger.Count = 0 Then Exit Sub
    ActiveDocument.Bookmarks("MibiErregerHinweis").Select
    Selection.TypeParagraph
    Selection.Font.Underline = wdUnderlineSingle
    Selection.TypeText text:="Mikrobiologie (Erreger)"
    Selection.Font.Underline = wdUnderlineNone
    Selection.TypeParagraph
    Selection.TypeText text:="Siehe Anlage."
    Selection.TypeParagraph
    
    ActiveDocument.Bookmarks("MibiErregerTabelle1").Select
' Überschrift zur Tabelle einfügen
    Selection.TypeText text:="Mikrobiologie-Befund (Erreger)" '& " vom " & ISHMED_Datum(bef.gDatum)
End Sub


Public Sub DoOPBericht(tb As Collection, counter As Integer)
    If tb Is Nothing Then Exit Sub
    If tb.Count = 0 Then Exit Sub
    
    Dim zdiagnose As String, zop As String, zopDatum As String
    Dim tabrow As Collection, i As Integer, is_empty As Boolean
    For Each tabrow In tb
        getTabVar tabrow, "ZOPDATUM", zopDatum, is_empty, False
        getTabVar tabrow, "ZDIAGNOSE<L>", zdiagnose, is_empty, False
        getTabVar tabrow, "ZOP<L>", zop, is_empty, False
    Next
    
    Dim zopDatumLen, opTextLen, diagTextLen As Long
    
    zopDatumLen = Len(zopDatum)
    opTextLen = Len(zop)
    diagTextLen = Len(zdiagnose)
    
    
    If zopDatumLen > 0 Then
        ActiveDocument.Bookmarks("OPBerichtPMD").Select
        Selection.TypeParagraph
        Selection.Font.Underline = wdUnderlineSingle
        Selection.TypeText text:="Operation, " & ISHMED_Datum(zopDatum)
        Selection.Font.Underline = wdUnderlineNone
        
        If opTextLen > 0 Or diagTextLen > 0 Then
            If diagTextLen > 0 Then
                Selection.TypeParagraph
                Selection.Font.Bold = True
                Selection.TypeText text:="Diagnosen: "
                Selection.Font.Bold = False
                Selection.TypeText text:=zdiagnose
            End If
            If opTextLen > 0 Then
                Selection.TypeParagraph
                Selection.Font.Bold = True
                Selection.TypeText text:="OP: "
                Selection.Font.Bold = False
                Selection.TypeText text:=zop
            End If
            Selection.TypeParagraph
        Else
            Selection.TypeParagraph
        End If
    End If
    
End Sub

Public Function getDecimalFormat(val As Variant) As String
    Dim s As String, s1 As String   'Hilfsstring

    ' Im internen Zahlenformat von SAP sind die Dezimalstellen durch einen Punkt
    ' abgesetzt. Tausenderstellen könnten ggf. durch Kommata abgetrennt sein (us-amerikanisches
    ' Zahlenformat). Zur Umwandlung in deutsches Zahlenformat werden Punkte durch Kommas und Kommas
    ' durch Punkte ersetzt.
   
    On Error GoTo Err_GetDecimalFormat
    s1 = Replace(val, ".", ",")
    If IsNumeric(s1) Then
        s = Trim(s1)
    Else
        s = Trim(val)
    End If
    getDecimalFormat = s
Exit_GetDecimalFormat:
    Exit Function
Err_GetDecimalFormat:
    MsgBox "Fehler in Routine GetDecimalFormat: " & Err.Description
    Resume Exit_GetDecimalFormat
End Function

 
 












