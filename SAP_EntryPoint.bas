Attribute VB_Name = "SAP_EntryPoint"

Option Explicit
Option Base 0
'2026 xt: taking over
' Die Variablen des Dokuments stehen in der (einzigen) Zeile 0 des Dokuments
' Die Variablen in Tabellen eines Dokuments muessen durch ihre Zeilennummer identifiziert werden
Const dokvar = "0000"
    
Rem IS-H*MED: WinWord Muster MAKRO für IS-H*MED Version 4.63 Wordcontainer
'
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
    "OpenPrinterA" (ByVal pPrinterName As String, phPrinter _
    As Long, ByVal pDefault As Long) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" ( _
    ByVal hPrinter As Long) As Long

Private Declare Function DeviceCapabilities Lib "winspool.drv" _
    Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, _
    ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
    ByVal dev As Long) As Long

Dim KeyWordFound    As Long
Global R3Table         As Object
'****************************************************************
' Word 2010 hat keine Dateifunktionen (Speichern, Drucken, etc.),
' wenn es aus einem OLE-Container (wie z. B. SAP-GUI) gestartet
' wird. Die Menüfunktionen lassen sich auch nicht über die Quick
' Access Toolbar einbinden - Word hängt sich auf, wenn man den Button
' klickt. Wenn man dieses kleine Makro über die Quick Access Toolbar
' einbindet, dann kann man wieder aus SAP drucken!
'****************************************************************
Sub Drucken()
    Dialogs(wdDialogFilePrint).Show
End Sub


Public Sub onStartup()
    'no operation here
    Debug.Print "onStartup" '
End Sub

'****************************************************************
'Muster Makro für IS-H*MED WordContainer Version 4.63
'
Public Sub IshmedTest()
'****************************************************************
' Initialisierung - Never change
'****************************************************************
'
'
' Makro erstellt am 25.01.2002 von berndt
Set objWord.oApp = Word.Application  'das war ursprünglich in AutoOpen - aber das hat gecrasht dort
   InitCustomDocumentProperties
   SetPersistentDate "SAP.CREATIONDATE", "AUTODATE"
    Application.ScreenUpdating = False
        
    ActiveDocument.PageSetup.FirstPageTray = wdPrinterUpperBin
    ActiveDocument.PageSetup.OtherPagesTray = wdPrinterLowerBin

 On Error GoTo ErrorHandling
KeyWordFound = 0
'****************************************************************
'* Felder Aktualisieren für Anlegedatum und Uhrzeit
'* (Ausbauen falls in Ihrer Vorlage Felder sind die nicht
'* aktualisiert werden sollen)
'****************************************************************
    ActiveDocument.Fields.Update
'****************************************************************
'****************************************************************
'* Einfügen des Briefdatums bzw. Ersetzen von [Datum]
'* (eingefügt von HS am 10.09.02)
'* Prozedur "InsertDate" befindet sich im Modul "EditAddress"
'****************************************************************
    ' XT____TMP____IMPLEMENT_PLS    InsertDate
'****************************************************************

'*  Zugriff auf die R3/Tabelle mit den zu ändernden Schlüsselwörter
    GetR3Table
    
'*  Große Verarbeitungsschleife ersetzt die einzelnen Texte'
ISHins
    
    
    meldungen_anzeigen
ishmed_test_exit:
    Application.ScreenUpdating = True
    Initialize  'Steuerelemente anpassen
    ' Application.ActiveDocument.Sections(1).Range.Select

Exit Sub
'
ErrorHandling:
    MsgBox "Die Fehlerbehandlung hat zugeschlagen mit Error: " + str(Err.Number)
    Error Err.Number
    GoTo ishmed_test_exit:
'
End Sub
'****************************** End of Main
'Hilfsroutine wird aus dem ABAP aufgerufen
Public Sub issaved()
'Diese Funktion muß existieren.
'Default ist 0 = nicht gesichert, sonst 1
'Wenn die Eigenschaft IsSaved nicht existiert, wird auch nicht gesichert angenommen
'Die Übergabe Wert ist Character
On Error GoTo ErrorHandling
'make sure that this property exist
Dim bolIsValid As Boolean, prop
'
    bolIsValid = False
    For Each prop In ActiveDocument.CustomDocumentProperties
        If Trim(prop.name) = "IsSaved" Then
            bolIsValid = True
        End If
    Next
    If Not bolIsValid Then
        ActiveDocument.CustomDocumentProperties.Add _
            name:="IsSaved", LinkToContent:=False, Value:="0", _
            Type:=msoPropertyTypeString
    End If

    If ActiveDocument.Saved = True Then
        ActiveDocument.CustomDocumentProperties("IsSaved").Value = "1" 'True
    Else
        ActiveDocument.CustomDocumentProperties("IsSaved").Value = "0" 'False
    End If
Exit Sub
'
ErrorHandling:
    MsgBox "Die Fehlerbehandlung hat zugeschlagen mit Error: " + str(Err.Number)
    Error Err.Number
'
End Sub
Rem************************************************************
'Diese Funktion darf nicht geändert werden.
'Die Kommunikation erfolgt über "R3Table"
'Wenn man wuesse, wie man eine R3Table anlegt,
'koennte man hier Testdaten einlesen und muesste das
'Dokument nicht immer ueber SAP starten.
Private Sub GetR3Table()
'    readR3Table
    Set R3Table = ActiveDocument.Container.LinkServer.Items("R3Table").Table
'    writeR3Table
End Sub

' das kann oben in sub GetR3Table durchlaufen werden (Kommentarzeichen entfernen),
' um den Datensatz eines Arztbriefs mitzuloggen. Mit dem sub unten readR3Table kann
' dieser Datensatz wieder eingelesen werden, wenn man die Dokumente in SAP nicht
' immer neu zusammenstellen will. Der Start des Arztbriefs aus R3 ist trotzdem ratsam
' weil sonst nix angezeigt wird.
Private Sub writeR3Table()
    Const fn = "C:\users\soykas\Desktop\R3Table.txt"
    Dim f, zeile, spalte
    f = FreeFile
    Open fn For Output As f
    For zeile = 1 To R3Table.RowCount
        Dim txt As String
        Write #f, Replace(R3Table(zeile, 1), """", "&#34;"), Replace(R3Table(zeile, 2), """", "&#34;"), Replace(R3Table(zeile, 3), """", "&#34;")
        txt = ""
    Next zeile
    Close f
End Sub

' Einlesen eines vorher mit writeR3Table mitgeloggten Datensatzes fuer einen Arztbrief,
' kuerzt die Testphase ab
Private Sub readR3Table()
    Const fn = "C:\Users\soykas\Desktop\R3Table.txt"
    Dim f, zeile, spalte
    f = FreeFile
    Open fn For Input As f
    While Not EOF(f)
        Dim a, b, c, row
        Input #f, a, b, c
        Set row = R3Table.Rows.Add
        a = Replace(a, "&#34;", """")
        row(1) = a
        b = Replace(b, "&#34;", """")
        row(2) = b
        c = Replace(c, "&#34;", """")
        row(3) = c
        zeile = zeile + 1
    Wend
    Close f
End Sub

Public Sub LabordatenAuswahl(gcolData As Collection)
    Dim doktyp As Collection, instanz As Collection, zeile
    Dim tabelle As Collection, feld As Collection
    Dim labor As New Collection
    Dim lfn As String, neue_lfn As String
    Dim datum, zeit, rank As String
    Dim is_empty As Boolean
    
    If gcolData Is Nothing Then Exit Sub
    
    If getColl(gcolData, "N2_LABOR:N2LABOR", doktyp) Then
        For Each instanz In doktyp
            If getTabOfInst(instanz, gcolData, "N2_LABOR:N2LABOR001", tabelle) Then
                getDocVar instanz, "N2LADATUM", datum, is_empty, False
                getDocVar instanz, "N2LATIME", zeit, is_empty, False
                For Each zeile In tabelle
                    ' die laufenden Nummer der Elemente der multiplen Strukturen
                    ' setzen sich zusammen aus der laufenden Nummer des Dokuments
                    ' (immer ganzen Vielfaches von 100, z. B. 0100 fuer das erste Dokument
                    ' der laufenden Nummer des Elements (z. B. 0101 fuer das erste Element
                    ' des ersten Dokuments.
                    ' Datum ueberall mit reinkopieren
                    Dim row_key As String
                    Dim col As Collection
                    Set col = New Collection
                    col.Add datum, "var"
                    zeile.Add col, "N2LADATUM"
                    Set col = Nothing
                    Set col = New Collection
                    col.Add zeit, "var"
                    zeile.Add col, "N2LATIME"
                    Set col = Nothing
                    changeValueName zeile
                    ' Reihenfolge wird aus SAP uebergeben,
                    ' entspricht der Reihenfolge im Kumulativbefund
                    ' wird eingestellt in der Tabellenpflege der Tabelle tn2kum02
                    ' rank = getValueRank(zeile, "N2KATTEXT")
                    Set col = Nothing
                    Set col = New Collection
                    row_key = zeile.Item(1).Item("inst_key") & zeile.Item(1).Item("row_key")
                    col.Add row_key, "var"
                    zeile.Add col, "row_key"
                    labor.Add zeile, row_key
                Next
            End If
        Next
    End If
    
    If labor.Count = 0 Then Exit Sub
    Dim waehlen As New Auswahl
    Set waehlen.daten = labor
    
    ' Fenster modal anzeigen
    waehlen.Show 1
    ' Hier gehts erst weiter, wenn das Fenster geschlossen wurde
    ' Fenster abraeumen, Objekt entfernen
    Unload waehlen
    Set waehlen = Nothing
End Sub

    ' Langtextfelder zusammenbauen, damit der Code bei der Ausgabe damit
    ' nicht belastet wird. Der Aufwand fuer ein neues doktyp soll moeglichst
    ' ueberschaubar bleiben. Die Teile des Langtextes werden
    ' aneinandergehaengt, wobei nach bestimmten Regeln entweder
    ' Zeilenvorschuebe oder Leerzeichen verwendet werden (in getTextLine).
Private Sub langtextVerarbeiten(langtexte As Collection)
    Dim str, i
    Dim tab_key, inst_key, row_key
    tab_key = langtexte("tab_key")
    inst_key = langtexte("inst_key")
    row_key = langtexte("row_key")
    langtexte.Remove "tab_key"
    langtexte.Remove "inst_key"
    langtexte.Remove "row_key"
    
    str = getTextLine(langtexte)
    ' Irgendwie gibt es kein "Clear"
    For i = langtexte.Count To 1 Step -1
        langtexte.Remove i
    Next
    
    langtexte.Add str, "var"
    langtexte.Add tab_key, "tab_key"
    langtexte.Add inst_key, "inst_key"
    langtexte.Add row_key, "row_key"
End Sub

Sub DumpCollectionImmediate(col As Collection)
  
End Sub

Private Sub ISHins()
    Debug.Print "V71.ISHins()"
    Dim key     As String
    Dim content As String
    Dim lfn     As String
    Dim i As Long, dp As Integer
    Dim doktyp As Collection
    Dim instanz As Collection
    Dim gcolData As New Collection
    
    ' Nur zum Test!
    ' writeR3Table
    
    WC_Extended.initglobValues
    
    ' debug.print Now()
    
    Dim inst_key As String
    
    inst_key = String(25, "0")

    Selection.StartOf UNIT:=wdStory
    For i = 1 To R3Table.RowCount
        key = ""
        content = ""
        lfn = ""
        ' enthaelt der Key einen Doppelpunkt, dann folgt die Tabelle
        dp = InStr(R3Table(i, 1), ":") - 1
        key = R3Table(i, 1)
        If Right(R3Table(i, 1), 14) = ":NEW-DOCUMENT]" Then
            inst_key = R3Table(i, 2)
        Else
            content = R3Table(i, 2)
        End If
        lfn = R3Table(i, 3)
        DoCase gcolData, key, inst_key, content, lfn
    Next i
    
    Dim tabrow As Collection, var As Collection
    ' debug.print "Vor Langtextverarbeitung " & Now()
    For Each doktyp In gcolData
        For Each instanz In doktyp
            For Each tabrow In instanz
                For Each var In tabrow
                    langtextVerarbeiten var
                Next
            Next
        Next
    Next
    
    ' debug.print "Vor Labordatenauswahl " & Now()
    
    LabordatenAuswahl gcolData
    
    DoAerzte
    
    Dim erreger_vorhanden, medikation_vorhanden As Boolean
    Dim labor_mit_ueberschrift As Boolean
    Dim inst As Collection
    ' Untergeordnete Tabellen zum jeweiligen Dokument
    Dim tb As Collection
    Dim tb1 As Collection
    
    labor_mit_ueberschrift = True
    erreger_vorhanden = False
    medikation_vorhanden = False
    
    ' debug.print "Vor Dokumentausgabe " & Now()
DumpCollectionImmediate gcolData
DumpCollectionImmediate gcolErreger
    If getColl(gcolData, "USER:DIATAB", doktyp) Then
        For Each inst In doktyp
            DoDiagnosen inst
        Next
    End If
    
    ' eingefügt von x-tention im Oktober 2022
    'If getColl(gcolData, "USER:PROCTAB", doktyp) Then
    '    For Each inst In doktyp
    '        DoProzeduren inst
    '    Next
    'End If
    
    If getColl(gcolData, "KOERPUNT:ZKOERPU", doktyp) Then
        For Each inst In doktyp
            DoKoerper inst
        Next
    End If
    If getColl(gcolData, "ZBLUTPRO:Y0000009", doktyp) Then
        For Each inst In doktyp
            Dim antikoerper As Collection, transfusionen As Collection, summen As Collection
            getTabOfInst inst, gcolData, "ZBLUTPRO:Y000000900", antikoerper
            getTabOfInst inst, gcolData, "ZBLUTPRO:Y000000903", summen
            DoBlutprod inst, antikoerper, summen
        Next
    End If
    If getColl(gcolData, "ZERREGER:Y000000G00", doktyp) Then
        For Each inst In doktyp
            ' Wenn das Flag erreger_vorhanden nicht gesetzt ist, wird die Erreger-Tabelle
            ' am Ende des Dokuments geloescht.
            erreger_vorhanden = True
            DoErregerDetails inst
        Next
    End If
    ' Hier werden nur die Tabellen des Dokuments ausgewertet,
    ' das Dokument enthaelt weiter keine interessanten Daten
    If getColl(gcolData, "ZBLUTKOM:Y000000A00", doktyp) Then
        For Each inst In doktyp
            DoBlutkomp inst
        Next
    End If
    If getColl(gcolData, "N2_LABOR:N2LABOR", doktyp) Then
        For Each inst In doktyp
            If getTabOfInst(inst, gcolData, "N2_LABOR:N2LABOR001", tb) Then
                DoLabor inst, tb, labor_mit_ueberschrift
                ' Nur das erste Labor mit Ueberschrift, Hummel
                labor_mit_ueberschrift = False
            End If
        Next
    End If
    
    
    
    ' XT: OP-Bericht: Daten einfuegen
    If getColl(gcolData, "ZOPBERIP:ZOPBERIPKH000000", doktyp) Then
        Dim counter As Integer
        counter = 0
        For Each inst In doktyp
            If getTabOfInst(inst, gcolData, "ZOPBERIP:ZOPBERIPKH000000", tb) Then
                ' DoLabor inst, tb, labor_mit_ueberschrift
                DoOPBericht tb, counter
                counter = counter + 1
            End If
        Next
    End If
    
    
    If getColl(gcolData, "ZRECHTHE:Y000001O", doktyp) Then
    For Each inst In doktyp
        DoRechtsherz inst
    Next
    End If
    If getColl(gcolData, "ZDIALYSE:Y000001J", doktyp) Then
        For Each inst In doktyp
            DoDialyse inst
        Next
    End If
    If getColl(gcolData, "ANAMN:ZANAMN", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Anamnese:", "Anamnese"
        Next
    End If
    If getColl(gcolData, "ZTXVISIT:Y000002L", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "TX-Visite vom ", "Opbericht"
        Next
    End If
    
    If getColl(gcolData, "ZOPBERIC:Y0000018", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "OP-Bericht vom ", "Opbericht"
        Next
    End If
    
    If getColl(gcolData, "ZROENTGB:Y0000015", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Röntgenbefund vom ", "ROE"
        Next
    End If
    
    If getColl(gcolData, "ZROENTGBEF:Y0000015", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Röntgenbefund vom ", "ROE"
        Next
    End If
    
    If getColl(gcolData, "ZEKGBELA:Y0000016", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Belastungs-EKG vom ", "EKGbelast"
        Next
    End If
    
    If getColl(gcolData, "ZEKGLANG:Y0000017", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Langzeit-EKG vom ", "EKGlangzeit"
        Next
    End If
    
    If getColl(gcolData, "ZLANGZRR:Y000001P", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Langzeit-Blutdruckmessung vom ", "LANGZRR"
        Next
    End If
    
    If getColl(gcolData, "Z6MINUTE:Y000000C", doktyp) Then
        For Each inst In doktyp
            getTabOfInst inst, gcolData, "Z6MINUTE:Y000000C00", tb
            DoGehtest inst, tb
        Next
    End If
    
    If getColl(gcolData, "ZENDOSKO:Y000001D", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Endoskopie/Gastroskopie vom ", "Endo"
        Next
    End If
    
    If getColl(gcolData, "ZKAROTID:Y0000029", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Karotidendoppler vom ", "Karotidendoppler"
        Next
    End If
    If getColl(gcolData, "ZKNOCHEN:Y0000007", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Knochenmarkbiopsie vom ", "Knochenmarkbiopsie"
        Next
    End If
    If getColl(gcolData, "ZSONOBAU:Y0000008", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Sonographie Oberbauch vom ", "Sonobauch"
        Next
    End If
    If getColl(gcolData, "ZSONOGEF:Y0000010", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Sonographie Gefäße vom ", "Sonogef"
        Next
    End If
    If getColl(gcolData, "ZSONO_SD:Y000000Z", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Sonographie Schilddrüse vom ", "Sonosd"
        Next
    End If
    If getColl(gcolData, "ZTSONOGR:Z00000000000000L", doktyp) Then '2026 Umstellung auf Tomtec
        For Each inst In doktyp
            DoBefund inst, "Sonographie vom ", "Sonograf"
        Next
    End If
    If getColl(gcolData, "ZSCHRITT:Y000001F", doktyp) Then
        For Each inst In doktyp
            getTabOfInst inst, gcolData, "ZSCHRITT:Y000001F00", tb
            getTabOfInst inst, gcolData, "ZSCHRITT:Z000000000000007", tb1
            DoSchrittmacher inst, tb, tb1
        Next
    End If
    If getColl(gcolData, "ZAUGKONS:Y000000S", doktyp) Then
    For Each inst In doktyp
        DoBefund inst, "Konsil Augenarzt vom ", "Augkonsil"
    Next
    End If
    If getColl(gcolData, "ZCHIKONS:Y000000O", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Konsil Chirurgie vom ", "Chikonsil"
        Next
    End If
    If getColl(gcolData, "ZDERKONS:Y000000R", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Konsil Dermatologie vom ", "Derkonsil"
        Next
    End If
    If getColl(gcolData, "ZGYNKONS:Y000000M", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Konsil Gynäkologie vom ", "Gynkonsil"
        Next
    End If
    If getColl(gcolData, "ZHNOKONS:Y000000V", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Konsil HNO vom ", "Hnokonsil"
        Next
    End If
    If getColl(gcolData, "ZNEUKONS:Y000000P", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Konsil Neurologie vom ", "Neukonsil"
        Next
    End If
    If getColl(gcolData, "ZSONKONS:Y000000Y", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Konsil vom ", "Sonstkonsil"
        Next
    End If
    If getColl(gcolData, "ZUROKONS:Y000000Q", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Konsil Urologie vom ", "Urokonsil"
        Next
    End If
    If getColl(gcolData, "ZTXVISIT:Y000002C", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "TX-Visite vom ", "Sonstkonsil"
        Next
    End If
    If getColl(gcolData, "ZZAHKONS:Y000000T", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Konsil Zahnarzt vom ", "Zahkonsil"
        Next
    End If
    If getColl(gcolData, "ZTECHOKA:Z00000000000000J", doktyp) Then 'Head
        Debug.Print "ZECHOKA:Z00000000000000J FOUND"
        For Each inst In doktyp
            getTabOfInst inst, gcolData, "ZTECHOKA:Z00000000000000K", tb
            DoEcho inst, tb
            'Neu, 25.04.06 von KD, Ausgabe der Daten aus dem doktyp 6-Minuten-Gehtest:
            '(auf Anforderung direkt hinter Echo ausgeben)
        Next
    Else
        Debug.Print "ZECHOKA:00000000000000J NOT FOUND"
    End If
    If getColl(gcolData, "ZEPIKRIS:Y000001A", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Epikrise: ", "Epikrise"
        Next
    End If
    If getColl(gcolData, "ZEKGBEF:Y000000I", doktyp) Then
        For Each inst In doktyp
            getTabOfInst inst, gcolData, "ZEKGBEF:Y000000I02", tb
            DoEKG inst, tb
        Next
    End If
    

    ' EM_20250701: Tomtec Echokardio kommt hinzu --------------------------------------
    '(auf Anforderung direkt vor Mikrobiologie (Erreger) ausgeben)
    Debug.Print "V71 xt 2026: VOR Tomtec Daten " & Now()
    If getColl(gcolData, "ZTECHOKA:Z00000000000000J", doktyp) Then
        Dim cnt As Integer
        cnt = 0
        For Each inst In doktyp
            ' EM_20250828: PrimTabelle und Messwerte werden nun benötigt !
            getTabOfInst inst, gcolData, "ZTECHOKA:Z00000000000000K", tb
            DoTEchokard inst, tb

        Next
    Else
        Debug.Print "ZTECHOKA:Z00000000000000J NOT FOUND"
    End If
    
    If getColl(gcolData, "ZTSONOGR:Z00000000000000L", doktyp) Then
        Dim cnt1 As Integer
        cnt1 = 0
        For Each inst In doktyp
            If getTabOfInst(inst, gcolData, "ZTSONOGR:Z00000000000000L", tb) Then
                ' DoBefund inst, "Tomtec Sono-Befund vom ", "Sonographie"
                DoTSonograf tb, cnt1
                cnt1 = cnt1 + 1
            End If
        Next
    Else
        Debug.Print "ZTSONOGR:Z00000000000000L NOT FOUND"
    End If
    Debug.Print "V71 xt 2026: NACH Tomtec Daten " & Now()
    ' --------------------------------------------------------------------------------
    
    If getColl(gcolData, "KOERPUNT:ZKOERPUNT", doktyp) Then
        For Each inst In doktyp
            DoKoerper inst
        Next
    End If
    If getColl(gcolData, "ZLIQUORP:Y0000007", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Liquorpunktion vom ", "Liquorp"
        Next
    End If
    If getColl(gcolData, "ZDIAETBE:ZDIAET0000000000", doktyp) Then
        For Each inst In doktyp
            DoBefund inst, "Diätempfehlung", "Diaet"
        Next
    End If
    If getColl(gcolData, "ZLNGFUNK:Y000000X", doktyp) Then
        For Each inst In doktyp
            DoLufu inst
        Next
    End If
    If getColl(gcolData, "ZMEDIKAT:Y0000011", doktyp) Then
        'Neu, 02.09.2015, KD: Dokumenttyp ZMEDIKATIO ersetzt durch ZENTLMEDIK
        For Each inst In doktyp
            getTabOfInst inst, gcolData, "ZMEDIKAT:Y000001100", tb
            DoMedikation inst, tb
        Next
    End If
    If getColl(gcolData, "ZENTLMED:ZEMED00000000002", doktyp) Then
        For Each inst In doktyp
            Dim medis As Collection
            getTabOfInst inst, gcolData, "ZENTLMED:ZEMED00000000003", medis
            DoMedikationNeu inst, medis
        Next
    End If
    If getColl(gcolData, "ZTHERAPI:ZTZERAP000000000", doktyp) Then
        For Each inst In doktyp
            Dim am As Collection, ab As Collection, mre As Collection
            getTabOfInst inst, gcolData, "ZTHERAPI:ZTZERAP000000001", ab
            getTabOfInst inst, gcolData, "ZTHERAPI:ZTZERAP000000002", am
            getTabOfInst inst, gcolData, "ZTHERAPI:ZTZERAP000000003", mre
            DoTherapie inst, ab, am, mre
        Next
    End If
    
    ' Keine Erreger? Dann die letzte Seite mit der Tabellenvorlage loeschen
    ' Achtung: Alles ab der letzten Unterschrift wird geloescht!
    If erreger_vorhanden = False Then
        Dim rg, rg1 As Range
        ' Dieser Range umfasst das ganze doktyp
        Set rg = ActiveDocument.Range(wdMainTextStory)
        ' Dieser Range startet und endet hinter den Unterschriften
        Set rg1 = ActiveDocument.Bookmarks("Unterschrift").Range
        ' ... bis er bis zum Ende des Dokuments ausgedehnt wird
        rg1.End = rg.End
        ' Der Bereich wird gelöscht
        rg1.Delete
    End If

    
Erstschriftlich
    
    ' debug.print "Vor Platzhaltersertzung " & Now()
    
'Datei bis zum Ende gelesen
    Selection.StartOf UNIT:=wdStory
'Nächster Befehl
'Gehe zu allen Platzhalter-Feldern (z.B. [NPAT....] oder [U:...])und lösche diese
'Ergebnis ist ein doktyp mit den Daten ohne Schlüsselfelder
'
    ' ersetze alle Felder mit [DATUM]-Platzhaltern
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\[DATUM\]"
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ' ersetze alle Felder mit [NPAT-...]-Platzhaltern
    Dim oStory As Range
    For Each oStory In ActiveDocument.StoryRanges
        oStory.Find.ClearFormatting
        oStory.Find.Replacement.ClearFormatting
        With oStory.Find
    '    With Selection.Find
            .text = "\[NPAT-*\]"
            .Replacement.text = " "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        oStory.Find.Execute Replace:=wdReplaceAll
    Next
    ' ersetze alle Felder mit [U*:*]-Platzhaltern
    With Selection.Find
        .text = "\[U*:*\]"
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' ersetze alle Felder mit [NBEW-...]-Platzhaltern
    With Selection.Find
        .text = "\[NBEW-*\]"
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Wrap = wdFindContinue
        .text = " {2;}"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
    End With
    
    ' zum Anfang des Dokuments springen:
    Selection.StartOf UNIT:=wdStory
    ' bestimmte Textmarken entfernen:
    DelBookmark
    ' Anhang MiBi (Erreger) löschen, wenn nicht benötigt
IshIns_Ende:
    ' debug.print "Ende von ISH Ins " & Now()
    ' Stop
    Exit Sub
'
ErrorHandling:
    MsgBox "Die Fehlerbehandlung hat zugeschlagen mit Error: " + str(Err.Number)
    Error Err.Number
    GoTo IshIns_Ende:
'
End Sub         'ISHins

Function VariableExists(v, key) As Boolean
    Dim it, exists
    exists = False
    For Each it In v
        If it.name = key Then
            exists = True
            Exit For
        End If
    Next
    VariableExists = exists
End Function

Function getColl(c As Collection, key, ret As Collection) As Boolean
    Dim OK As Boolean
    OK = False
    Set ret = Nothing
    On Error GoTo getColl_Error
    Set ret = c.Item(key)
    OK = True
getColl_Error:
    getColl = OK
End Function

' Liefert die Tabelle mit key der gleichen Instanz wie das uebergebene Dokument
' key sollte der Schluessel einer Tabelle zum gleichen Dokumenttyp sein!
' Die Dokument-Instanz ist in der Collection jeder Variablen unter dem Schluessel "inst_key"  gespeichert.
Function getTabOfInst(c As Collection, dok As Collection, key, ret As Collection) As Boolean
    Dim OK As Boolean
    Dim inst As String
    OK = False
    Set ret = Nothing
    On Error GoTo getTabOfInst_Error
    inst = c.Item(1).Item(1).Item("inst_key")
    Set ret = dok.Item(key).Item(inst)
    OK = True
getTabOfInst_Error:
    getTabOfInst = OK
End Function

Function getFromCol(col As Collection, key, ByRef ret) As Boolean
    Dim OK As Boolean
    OK = False
    ret = ""
    On Error GoTo getFromCol_Error:
    ret = col.Item(key)
    OK = True
getFromCol_Error:
    getFromCol = OK
End Function

Function getTabVar(tabrow As Collection, key As String, ByRef ret, ByRef is_empty As Boolean, numerisch As Boolean) As Boolean
    Dim wert
    Dim s As String
    Dim OK As Boolean
    Dim zahl As Double
    OK = False
    is_empty = True
    On Error GoTo getTabVar_Error
    wert = Trim(tabrow.Item(key).Item("var"))
    If numerisch Then
        s = getDecimalFormat(wert)
        If IsNumeric(s) Then
            zahl = CDbl(s)
            ret = zahl
            is_empty = zahl = 0
            OK = True
        Else
            OK = False
        End If
    Else
        is_empty = wert = ""
        ret = wert
        OK = True
    End If
getTabVar_Error:
    On Error GoTo 0
    getTabVar = OK
End Function

Function getDocVar(tabelle As Collection, key As String, ByRef ret, ByRef is_empty As Boolean, numerisch As Boolean) As Boolean
    Dim wert
    Dim inhalt As Collection
    Dim OK As Boolean
    Dim zahl As Double
    OK = False
    is_empty = False
    On Error GoTo getDocVar_Error
    OK = getTabVar(tabelle.Item(dokvar), key, ret, is_empty, numerisch)
    Set inhalt = tabelle.Item(dokvar).Item(key)
getDocVar_Error:
    On Error GoTo 0
    getDocVar = OK
End Function

Function getAnyDocVar(c As Collection, keys, ByRef ret, ByRef is_empty As Boolean, numerisch As Boolean) As Boolean
    Dim wert
    Dim i As Integer
    Dim leer As Boolean
    Dim inhalt As Collection
    Dim OK As Boolean
    Dim zahl As Double
    OK = False
    is_empty = True
    On Error GoTo getAnyDocVar_Error:
    Set inhalt = c.Item(dokvar)
    OK = getAnyTabVar(inhalt, keys, ret, is_empty, numerisch)
getAnyDocVar_Error:
    getAnyDocVar = OK
End Function

Function getAnyTabVar(c As Collection, keys, ByRef ret, ByRef is_empty As Boolean, numerisch As Boolean) As Boolean
    Dim wert
    Dim i As Integer
    Dim leer As Boolean
    Dim inhalt As Collection
    Dim OK As Boolean
    Dim zahl As Double
    OK = False
    is_empty = True
    For i = LBound(keys) To UBound(keys)
        On Error GoTo getAnyTabVar_Error:
        Set inhalt = c.Item(keys(i))
        wert = inhalt.Item("var")
        wert = Trim(wert)
        If numerisch Then
            If IsNumeric(getDecimalFormat(wert)) Then
                zahl = CDbl(getDecimalFormat(wert))
                ret = zahl
                is_empty = zahl = 0
                OK = True
            Else
                OK = False
            End If
        Else
            is_empty = wert = ""
            ret = wert
            OK = True
        End If
getAnyTabVar_Error:
        Resume getAnyTabVar_Next:
getAnyTabVar_Next:
        If OK Then Exit For
    Next
    getAnyTabVar = OK
End Function

Private Sub DoCase(gcolData As Collection, key As String, inst_key As String, content As String, lfn As String)
    '****************************************************************
    ' Haupt-Routine: User - Change
    '
    ' Hier werden die aus R/3 kommenden Standart Daten im doktyp
    ' eingesetzt. Je nachdem, wie das
    ' Feld zu behandeln ist stehen drei Standard - Verfahren zur
    ' verfügung:
    '
    '   eigenes Macro(Sub/Function) und/oder
    '   Call ISHReplaceAll(key, content)
    '   Call ISHReplace(key, content)
    '   Call ISHInsert(key, content)
    '   Call ISHReplaceAllKopfFuss(key, content)
    '
    ' Falls Felder eine spezielle Behandlung benötigen sind
    ' geeignete Macros selbst zu programmieren.
    '****************************************************************
    Dim Fallart As String
    Dim MyNr As Long
    Dim doktyp As Collection ' Dokumente
    Dim instanz As Collection ' Instanzen, es koennen mehrere Kopien uebergeben werden
    Dim t As Collection ' Entweder Instanz oder multiple Struktur
    
    Dim exists As Boolean
    Dim itm, keys, parts

' On Error GoTo ErrorHandling

'
    ' Einlesen der Patientendaten, die aus SAP uebergeben werden.
    ' Diese Daten werden als Dokumentvariablen gespeichert. Beim
    ' Exportieren des Dokuments als MDM-Nachricht wird auf die
    ' Patientendaten zugegriffen.
    If Left(key, 5) = "[SYS-" Then
        ISHReplaceAll key, content
    End If
    If Left(key, 6) = "[NPAT-" Or Left(key, 6) = "[NFAL-" Or _
        Left(key, 5) = "[SYS-" Or key = "[NBEW-BWIDT]" Or _
        Left(key, 3) = "[U:" Or key = "[ENTL-BWIDT]" Then
        If Not VariableExists(ActiveDocument.Variables, key) Then
            ActiveDocument.Variables.Add name:=key, Value:=content
        End If
    Else
        ' Die Eingabedaten werden durchlaufen und in einem dreistufigen
        ' Baum abgelegt. Auf der obersten Ebene wird in einer Collection
        ' nach Dokumenttypen unterschieden, zu jedem Dokumenttyp gibt es
        ' auf der zweiten Ebene eine oder mehrere Instanzen (z. B. mehrere Epikrisen). Auf der
        ' untersten Ebene gibt es zu jeder Dokumentinstanz die Felder mit
        ' den Dateninhalten. Als Ausnahmefall koennen darunter noch einmal
        ' multiple Strukturen (Tabellen) angelegt, die Datenfelder enthalten.
        ' Wenn alles eingesammelt ist, wird der Inhalt
        ' der Datenfelder ausgegeben. Die Ausgabe richtet sich nach dem
        ' Dokumenttyp (z. B. DoEKG), ausser den den sehr aehnlich aufgebauten
        ' Konsildokumenten, die in der gemeinsamen Funktion DoBefund
        ' ausgegeben werden (Unterschiede werden ueber Parameter uebergeben).
        ' Weil es nicht moeglich ist, den key eines Elements in einer Collection
        ' auszulesen, wird jedes Element als Array aus key und Dateninhalt erzeugt
        ' und gespeichert. Der key existiert also bei jedem Element zweimal:
        ' einmal als element(0) und dann als key in der Collection.
        
        ' Vorne Dokumenttyp ggf. mit Tabelle als Key, hinten Feldname
        ' Content ist der Dateninhalt
        ' Stop
        Dim tab_key, var_key, row_key
        Dim tabelle As Collection, tabrow As Collection, var As Collection
        
        keys = Split(Mid(key, 2, Len(key) - 2), "-")
        tab_key = keys(0) & ""
        var_key = keys(1) & ""
        row_key = Right(dokvar & lfn, Len(dokvar)) ' Innerhalb einer Tabelle die Zeilenummer
        
        If Not getColl(gcolData, tab_key, tabelle) Then
            Set tabelle = New Collection
            gcolData.Add tabelle, tab_key
        End If

        If Not getColl(tabelle, inst_key, instanz) Then
            Set instanz = New Collection
            tabelle.Add instanz, inst_key
        End If
        
        If Not getColl(instanz, row_key, tabrow) Then
            Set tabrow = New Collection
            instanz.Add tabrow, row_key
        End If
        
        If Not getColl(tabrow, var_key, var) Then
            Set var = New Collection
            tabrow.Add var, var_key
            var.Add tab_key, "tab_key"
            var.Add inst_key, "inst_key"
            var.Add row_key, "row_key"
        End If
        
        var.Add content

    End If

    Select Case key
    Case "[NPAT-ANRED]"
        Select Case content
        Case "01"
            ISHReplace key, "Herrn"
            ISHReplaceAll "[U:ARTIKEL]", "der"
        Case "02"
            ISHReplace key, "Frau"
            ISHReplaceAll "[U:ARTIKEL]", "die"
            ISHReplaceAll "des Patienten", "der Patientin"
        Case Else
            ISHReplace key, content
            ISHReplaceAll "[U:ARTIKEL]", "der"
        End Select
    Case "[NPAT-TITEL]"
        Select Case content
        Case "01"
            ISHReplace key, "Dr."
        Case "02"
            ISHReplace key, "Prof."
        Case Else
            ISHReplace key, content
        End Select
    Case "[NPAT-GBDAT]"
        ISHReplaceAll key, ISHMED_Datum(content)
        ISHReplaceAllKopfFuss key, ISHMED_Datum(content)
    Case "[NPAT-NAMZU]", "[NPAT-VNAME]", "[NPAT-NNAME]"
        ISHReplaceAll key, content
        ISHReplaceAllKopfFuss key, content
    Case "[NPAT-GBNAM]", "[NPAT-STRAS]", "[NPAT-LAND]", "[NPAT-PSTLZ]", "[NPAT-ORT]", "[NPAT-ORT2]", "[NPAT-FAMST]", "[NPAT-AGNAM]", "[NPAT-AGNUM]", "[NPAT-GSCHL]"
        ISHReplace key, content
'Besuchsdaten / abhängig von der Fallart
    Case "[NBEW-BWIDT]", "[U:ENTL-BWIDT]"
        If content <> "??" And content <> "" Then
          ISHReplaceAll key, ISHMED_Datum(content)
        Else
          ISHReplaceAll key, content
        End If
' Diagnosedaten
    Case "[NDIA-DITXT]", "[NDIA-DKAT1]", "[NDIA-DKEY1]", "[NKDI-DTEXT1]", "[NKDI-DTEXT2]", "[NKDI-DTEXT3]", "[NDIA-DTEXT1]"
        ISHInsert key, content
'
'****************************************************************
' Ab hier können für die aus dem Fremddatenbaustein des Doktyps
' kommenden Daten Bearbeitet werden.
' Der prinzipielle Aufbau ist:
'
'   Case "Schlüssel"
'       eigenes Macro(Sub/Function) und/oder
'       Call ISHReplaceAll(key, content) oder
'       Call ISHReplace(key, content) oder
'       CallISHInsert(key, content) oder
'       Call ISHReplaceAllKopfFuss(key, content)
'
'
' Der Schlüssel setzt sich wie folgt zusammen:
'           USER:Tabellenname-Feld
'****************************************************************
'einfügen ender
' auskommentiert soyka 2020-10-30, entsprechend den Aenderungen im Funktionsbaustein Z_PARAMDOKU_WORD
' Diagnosen werden spaeter wie ein Dokument ausgegeben
Case "[USER:DIATAB-DIA_DKEY]"
'        ISHInsert key, content
Case "[USER:DIATAB-DIA_LANGTEXT]"
'        ISHInsert key, content
Case "[USER:DIATAB-DIA_LANGTEXT]<L>"

' eingefügt durch x-tention im Oktober 2022 - ANFANG
Case "[USER:PROCTAB-PROC_ICPML]"
        ISHInsert key, content
Case "[USER:PROCTAB-PROC_LANGTEXT]"
        ISHInsert key, content
' eingefügt durch x-tention im Oktober 2022 - ENDE

'einfügen ender
'****************************************************************
' Ende der User-Abhängigen Programmierung
'****************************************************************
    Case Else
       If Right(key, 4) <> "<L>]" Then
         ISHReplaceAll key, content
       End If
'Sinnvoll wärend der Test/Einführungsphase zu aktivieren,
'damit Fehler in der Kommunikation gefunden werden
    'EndDokument
    'MsgBox "Schlüsselwort unbekannt: " + Mid(key, 2, Len(key) - 1) + "," + content + "," + lfn$
    End Select
    
Exit Sub
'
ErrorHandling:
    MsgBox "Die Fehlerbehandlung hat zugeschlagen mit Error: " + str(Err.Number)
    Error Err.Number
'
End Sub

Private Sub ListKeyWords(key As String, content As String, lfn As String)
    Selection.EndOf UNIT:=wdStory
    Selection.TypeParagraph
    Selection.TypeText (key + "," + content + "," + lfn)
End Sub
'
Private Sub ISHReplace(key As String, content As String)
'****************************************************************
' Ersetzen von Platzhaltern  - Never Change
' Es wird nur der erste Platzhalter ersetzt
'****************************************************************
    content = Trim(content)
    
    '22.02.07, KD: da <content> mehr als 255 Zeichen enthalten kann, das <Replacement>-Merkmal
    'des Selection.Find-Attributs jedoch max. 255 Zeichen aufnimmt, wird im Folgenden mit einer
    'Hilfsvariablen <rest> der <content> ggf. abgeschnitten und der <rest> nach dem Erstzen weiter
    'unten einfach mit ausgegeben:
    Dim rest As String
    
    If Len(content) > 100 Then
        rest = Mid$(content, 101)
        content = Left$(content, 100)
    Else
        rest = ""
    End If
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = key
        content = Replace(content, "^", "^^")
        .Replacement.text = content
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    If Selection.Find.Found = True Then
        KeyWordFound = 1
        
        If rest <> "" Then
            Selection.MoveRight
            Selection.TypeText rest
        End If
    End If
End Sub         ' ISHReplace

Private Sub ISHReplaceAll(key As String, content As String)
'****************************************************************
' Ersetzen von Platzhaltern  - Never Change
' Es werden alle Platzhalter ersetzt
'****************************************************************
    content = Trim(content)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = key
        content = Replace(content, "^", "^^")
        .Replacement.text = content
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    If Selection.Find.Found = True Then
        KeyWordFound = 1
    End If
End Sub         ' ISHReplace

Private Sub ISHReplaceAllKopfFuss(key As String, content As String)
Static KopfFirstTime As Boolean
' Problem 2. Seite Kopf einmalig neue Seite erzeugen
    If KopfFirstTime Then
        Selection.EndKey UNIT:=wdStory
        Selection.InsertBreak Type:=wdPageBreak
        KopfFirstTime = False
    End If
    
    content = Trim(content)
    
'    Selection.HomeKey UNIT:=wdStory
'     If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
'        Dim p As Word.Pane, firstpane As Boolean
'        firstpane = True
'        ' Alle Panes ausser der ersten schliessen
'        For Each p In Application.ActiveDocument.ActiveWindow.Panes
'            If firstpane Then
'                firstpane = False
'            Else
'                p.Close
'            End If
'        Next
'     End If
'     ' Viewtype auf PageView umstellen
'     If ActiveWindow.ActivePane.View.Type <> wdPageView Then
'        ActiveWindow.ActivePane.View.Type = wdPageView
'     End If
'     ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
'' Replace einfügen Anfang
'     Selection.Find.ClearFormatting
'     Selection.Find.Replacement.ClearFormatting
'     With Selection.Find
'        .text = key
'        content = Replace(content, "^", "^^")
'        .Replacement.text = content
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'     End With
'     Selection.Find.Execute Replace:=wdReplaceAll
'     If Selection.Find.Found = True Then
'        KeyWordFound = 1
'     End If
'' Replace einfügen Ende
'     If Selection.headerfooter.IsHeader = True Then
'        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
'     Else
'        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
'     End If
' 'Replace einfügen Anfang
'     Selection.Find.ClearFormatting
'     Selection.Find.Replacement.ClearFormatting
'     With Selection.Find
'        .text = key
'        content = Replace(content, "^", "^^")
'        .Replacement.text = content
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'     End With
'     Selection.Find.Execute Replace:=wdReplaceAll
'     If Selection.Find.Found = True Then
'        KeyWordFound = 1
'     End If
'    ' Replace einfügen Ende
    
    Dim section As Word.section, hf As Word.HeaderFooter
    For Each section In Application.ActiveDocument.Sections
        For Each hf In section.Headers
            'Replace einfügen Anfang
            hf.Range.Select
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = key
                content = Replace(content, "^", "^^")
                .Replacement.text = content
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            If Selection.Find.Found = True Then
                KeyWordFound = 1
            End If
            'Replace einfügen Ende
        Next
        For Each hf In section.Footers
            'Replace einfügen Anfang
            hf.Range.Select
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = key
                content = Replace(content, "^", "^^")
                .Replacement.text = content
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            If Selection.Find.Found = True Then
                KeyWordFound = 1
            End If
            'Replace einfügen Ende
        Next
        'Replace einfügen Anfang
        section.Range.Select
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = key
            content = Replace(content, "^", "^^")
            .Replacement.text = content
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        If Selection.Find.Found = True Then
            KeyWordFound = 1
        End If
        'Replace einfügen Ende
    Next
    
    ' Application.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    ' Selection.HomeKey UNIT:=wdStory
End Sub         'ISHReplaceAllKopfFuss
'
Public Sub ISHInsert(key As String, content As String)

    Selection.StartOf UNIT:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = key
        content = Replace(content, "^", "^^")
        .Replacement.text = Left(content, 100)
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
        KeyWordFound = 1
        If Selection.Bookmarks.Count > 0 Then
            If Selection.Bookmarks.Item(1).name = "ISHTABELLE" Then
                Selection.MoveUp UNIT:=wdParagraph, Count:=1
                ' hier wird der Wert zum o.a. Schlüssel wieder zusammengesetzt,
                ' da ISHReplace wieder den kompletten Wert benötigt, der oben
                ' noch auseinander genommen wurde:
                'If rest <> "" Then
                '    content = content & rest
                ISHReplace key, content
            End If
        Else
            Selection.MoveDown UNIT:=wdParagraph, Count:=1
            Selection.MoveUp UNIT:=wdParagraph, Count:=1, Extend:=wdExtend
            With ActiveDocument.Bookmarks
                .Add Range:=Selection.Range, name:="ISHTABELLE"
                .DefaultSorting = wdSortByName
                .ShowHidden = False
            End With
            Selection.Copy
            Selection.MoveDown UNIT:=wdParagraph, Count:=1
            Selection.Paste
            Selection.MoveUp UNIT:=wdParagraph, Count:=2
            ' hier wird der Wert zum o.a. Schlüssel wieder zusammengesetzt,
            ' da ISHReplace wieder den kompletten Wert benötigt, der oben
            ' noch auseinander genommen wurde:
            'If rest <> "" Then
            '    content = content & rest
            ISHReplace key, content
        End If
    End If
End Sub         ' ISHInsert
'
Public Function ISHMED_Datum(content As String) As String
Dim Jahr As String
Dim Monat As String
Dim Tag As String
    Jahr = Left(content, 4)
    Monat = Mid(content, 5, 2)
    Tag = Right(content, 2)
    ISHMED_Datum = Tag + "." + Monat + "." + Jahr
End Function        ' ISHMED_Datum
'
Public Function ISHMED_Uhrzeit(content As String) As String
Dim Stunde As String
Dim Minute As String
    Stunde = Left(content, 2)
    Minute = Mid(content, 3, 2)
    ISHMED_Uhrzeit = Stunde + ":" + Minute
End Function        ' ISHMED_Uhrzeit

' Disable an entry in the command bar.
'
Public Sub DisableCommandBarEntry(CommandBarIconName As String)
  Dim cb        As CommandBar
  Dim cbc       As CommandBarControl
'
  On Error GoTo ExitLine
'
  For Each cb In ThisDocument.CommandBars
    For Each cbc In cb.Controls
      ' Disable the icon.
      '
      If InStr(1, cbc.Caption, CommandBarIconName, vbTextCompare) > 0 Then
        cbc.Enabled = False
      End If
    Next cbc
  Next cb
'
ExitLine:
End Sub

' Enable an entry in the comand bar.
'
Public Sub EnableCommandBarEntry(CommandBarIconName As String)
  Dim cb As CommandBar
  Dim cbc As CommandBarControl
'
  On Error GoTo ExitLine
'
  For Each cb In ThisDocument.CommandBars
    For Each cbc In cb.Controls
      ' Enable the icon.
      '
      If InStr(1, cbc.Caption, CommandBarIconName, vbTextCompare) > 0 Then
        cbc.Enabled = True
      End If
    Next cbc
  Next cb
'
ExitLine:
End Sub

' Disable an entry in the menu.
'
Public Sub DisableMenuEntry(Menu As String, MenuItem As String)
  Dim cb As CommandBar
  Dim cbc As CommandBarControl
'
  On Error GoTo ExitLine
'
  Set cb = CommandBars(Menu)
  For Each cbc In cb.Controls
    ' Enable menu entrys
    '
    If InStr(1, cbc.Caption, MenuItem, vbTextCompare) > 0 Then
        cbc.Enabled = False
    End If
  Next cbc
'
ExitLine:
End Sub

' Enable an entry in the menu.
'
Public Sub EnableMenuEntry(Menu As String, MenuItem As String)
  Dim cb As CommandBar
  Dim cbc As CommandBarControl
'
  On Error GoTo ExitLine
'
  Set cb = CommandBars(Menu)
  For Each cbc In cb.Controls
    ' Enable menu entrys
    '
    If InStr(1, cbc.Caption, MenuItem, vbTextCompare) > 0 Then
        cbc.Enabled = True
    End If
  Next cbc
'
ExitLine:
End Sub
'
Private Sub DelBookmark()
On Error Resume Next
    ActiveDocument.Bookmarks("Test").Delete
End Sub
'
Sub ishmed_show_r3table()
'*
On Error GoTo ErrorHandling
Dim i As Long
'
    If R3Table Is Nothing Then
       GetR3Table
    End If
    Selection.EndOf UNIT:=wdStory
    Selection.TypeParagraph
    Selection.TypeText ("KEY            " + "," + "Content    " + "," + " LFN ")
    Selection.TypeParagraph
'
    For i = 1 To R3Table.RowCount
        Selection.TypeParagraph
        Selection.TypeText (R3Table(i, 1) + "," + R3Table(i, 2) + "," + R3Table(i, 3))
    Next i
'
Exit Sub
'
ErrorHandling:
    MsgBox "Die Fehlerbehandlung hat zugeschlagen mit Error: " + str(Err.Number)
    Error Err.Number
'
End Sub


Sub check_adress()
 Dim F_name(11) As String, i As Long
 
 F_name(1) = "ANREDEERST"
 F_name(2) = "NAMEERST"
 F_name(3) = "STRASSERST"
 F_name(4) = "PLZERST"
 F_name(5) = "ORTERST"
 F_name(6) = "NAMEINFO1"
 F_name(7) = "NAMEINFO2"
 F_name(8) = "NAMEINFO3"
 F_name(9) = "NAMEINFO4"
 F_name(10) = "NAMEINFO5"
 F_name(11) = "ANSPRACHE"

 If (ActiveDocument.FormFields.Count < 11) Then
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst
    Selection.MoveDown UNIT:=wdLine, Count:=27, Extend:=wdExtend
    Selection.Cut
    For i = 1 To 10
      Selection.TypeParagraph
    Next i
    For i = 1 To 3
     ActiveDocument.FormFields.Add Range:=Selection.Range, Type:=wdFieldFormTextInput
     ActiveDocument.FormFields(i).name = F_name(i)
     Selection.TypeParagraph
    Next i
    ActiveDocument.FormFields.Add Range:=Selection.Range, Type:=wdFieldFormTextInput
    ActiveDocument.FormFields(4).name = F_name(4)
    Selection.TypeText " "
    ActiveDocument.FormFields.Add Range:=Selection.Range, Type:=wdFieldFormTextInput
    ActiveDocument.FormFields(5).name = F_name(5)
    For i = 1 To 4
      Selection.TypeParagraph
    Next i
    For i = 6 To 10
     ActiveDocument.FormFields.Add Range:=Selection.Range, Type:=wdFieldFormTextInput
     ActiveDocument.FormFields(i).name = F_name(i)
     Selection.TypeParagraph
    Next i
    For i = 1 To 3
      Selection.TypeParagraph
    Next i
    ActiveDocument.FormFields.Add Range:=Selection.Range, Type:=wdFieldFormTextInput
    ActiveDocument.FormFields(11).name = F_name(11)
    Selection.TypeText ","
    Selection.TypeParagraph
      
  End If
 
 


End Sub



Sub Erstschriftlich()

    Dim doc As Word.Document
    Dim propName1 As Office.DocumentProperty

    Set doc = ActiveDocument

    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "Im Anzeigemodus steht die Funktion nicht zur Verfügung", vbInformation
        Exit Sub
    End If

    On Error Resume Next
    Set propName1 = doc.CustomDocumentProperties("name1")
    On Error GoTo 0

    If propName1 Is Nothing Or Len(propName1.Value) = 0 Then

        With doc.FormFields
            .Item("NAMEERST").Result = _
                "An die weiterbehandelnde Kollegin/" & vbCr & _
                "den weiterbehandelnden Kollegen"

            .Item("ANSPRACHE").Result = _
                "Sehr geehrte Kollegin/sehr geehrter Kollege"

            .Item("ANREDEERST").Result = ""
            .Item("STRASSERST").Result = ""
            .Item("PLZERST").Result = ""
            .Item("ORTERST").Result = ""
        End With

    Else

        fill_adress _
            doc.CustomDocumentProperties("anrede1"), _
            doc.CustomDocumentProperties("name1"), _
            doc.CustomDocumentProperties("Firma1"), _
            doc.CustomDocumentProperties("strass1"), _
            doc.CustomDocumentProperties("plz1"), _
            doc.CustomDocumentProperties("ort1"), _
            doc.CustomDocumentProperties("ansprache1")

    End If

    fill_info 1

End Sub


Public Sub Erstschriftlich_prev()
'
   'check_adress
  
   If ActiveDocument.ProtectionType <> wdNoProtection Then
      MsgBox "Im Anzeigemodus steht die Funktion nicht zur Verfügung"
      Exit Sub
   End If
      
   'If (ActiveDocument.FormFields(1).name <> ("NAMEERST").) Then
   '  ActiveDocument.FormFields.Add Range:=Selection.Range, Type:=wdFieldFormTextInput
   'End If
   'ActiveDocument.GoTo What:=wdGoToPage, Which:=wdGoToFirst
    
    If Trim(ActiveDocument.CustomDocumentProperties("name1").Value) <> "" Then
    
    fill_adress ActiveDocument.CustomDocumentProperties("anrede1").Value, _
               ActiveDocument.CustomDocumentProperties("name1").Value, _
               ActiveDocument.CustomDocumentProperties("Firma1").Value, _
               ActiveDocument.CustomDocumentProperties("strass1").Value, _
               ActiveDocument.CustomDocumentProperties("plz1").Value, _
               ActiveDocument.CustomDocumentProperties("ort1").Value, _
               ActiveDocument.CustomDocumentProperties("ansprache1").Value
   Else
   
    ActiveDocument.FormFields("NAMEERST").Result = "An die weiterbehandelnde Kollegin/" & Chr$(13) & "den weiterbehandelnden Kollegen"
    ActiveDocument.FormFields("ANSPRACHE").Result = "Sehr geehrte Kollegin/sehr geehrter Kollege"
    ActiveDocument.FormFields("ANREDEERST").Result = ""
    ActiveDocument.FormFields("STRASSERST").Result = ""
    ActiveDocument.FormFields("PLZERST").Result = ""
    ActiveDocument.FormFields("ORTERST").Result = ""
    
   End If
   fill_info 1
 
End Sub
Private Sub fill_adress(anrede As Variant, _
                        name As Variant, _
                        firma As Variant, _
                        strasse As Variant, _
                        plz As Variant, _
                        ort As Variant, _
                        ansprache As Variant)
Dim nocheinstring As String

   ActiveDocument.FormFields("ANREDEERST").Result = ""
   ActiveDocument.FormFields("NAMEERST").Result = ""
   ActiveDocument.FormFields("STRASSERST").Result = ""
   ActiveDocument.FormFields("PLZERST").Result = ""
   ActiveDocument.FormFields("ORTERST").Result = ""
   ActiveDocument.FormFields("ANSPRACHE").Result = ""
   nocheinstring = ""
   
   ActiveDocument.FormFields("ANREDEERST").Result = Trim(anrede)
   If Trim(firma) <> "" Then
       nocheinstring = Chr$(13) & firma
   End If
   ActiveDocument.FormFields("NAMEERST").Result = name & nocheinstring
   ActiveDocument.FormFields("STRASSERST").Result = strasse
   ActiveDocument.FormFields("PLZERST").Result = plz
   ActiveDocument.FormFields("ORTERST").Result = ort
   'If (Trim(ansprache) <> "") Then
   If Trim(ansprache) = "Sehr geehrter Herr Kollege" Or Trim(ansprache) = "Sehr geehrte Frau Kollegin" Then
         ActiveDocument.FormFields("ANSPRACHE").Result = Trim(ansprache) & " " & name
   Else
         ActiveDocument.FormFields("ANSPRACHE").Result = Trim(ansprache)
   End If
   
End Sub
Public Sub Nachrichtlich()
'
   If ActiveDocument.ProtectionType <> wdNoProtection Then
      MsgBox "Im Anzeigemodus steht die Funktion nicht zur Verfügung"
      Exit Sub
   End If
   
   fill_adress ActiveDocument.CustomDocumentProperties("anrede2").Value, _
               ActiveDocument.CustomDocumentProperties("name2").Value, _
               ActiveDocument.CustomDocumentProperties("Firma2").Value, _
               ActiveDocument.CustomDocumentProperties("strass2").Value, _
               ActiveDocument.CustomDocumentProperties("plz2").Value, _
               ActiveDocument.CustomDocumentProperties("ort2").Value, _
               ActiveDocument.CustomDocumentProperties("ansprache2").Value
   
   fill_info 2

End Sub
Public Sub Nachrichtlich3()

   If ActiveDocument.ProtectionType <> wdNoProtection Then
      MsgBox "Im Anzeigemodus steht die Funktion nicht zur Verfügung"
      Exit Sub
   End If

   fill_adress ActiveDocument.CustomDocumentProperties("anrede3").Value, _
               ActiveDocument.CustomDocumentProperties("name3").Value, _
               ActiveDocument.CustomDocumentProperties("Firma3").Value, _
               ActiveDocument.CustomDocumentProperties("strass3").Value, _
               ActiveDocument.CustomDocumentProperties("plz3").Value, _
               ActiveDocument.CustomDocumentProperties("ort3").Value, _
               ActiveDocument.CustomDocumentProperties("ansprache3").Value
   fill_info 3

End Sub
Public Sub Nachrichtlich4()

   If ActiveDocument.ProtectionType <> wdNoProtection Then
      MsgBox "Im Anzeigemodus steht die Funktion nicht zur Verfügung"
      Exit Sub
   End If
   
   fill_adress ActiveDocument.CustomDocumentProperties("anrede4").Value, _
               ActiveDocument.CustomDocumentProperties("name4").Value, _
               ActiveDocument.CustomDocumentProperties("Firma4").Value, _
               ActiveDocument.CustomDocumentProperties("strass4").Value, _
               ActiveDocument.CustomDocumentProperties("plz4").Value, _
               ActiveDocument.CustomDocumentProperties("ort4").Value, _
               ActiveDocument.CustomDocumentProperties("ansprache4").Value
   fill_info 4

End Sub
Public Sub Nachrichtlich5()
   
   If ActiveDocument.ProtectionType <> wdNoProtection Then
      MsgBox "Im Anzeigemodus steht die Funktion nicht zur Verfügung"
      Exit Sub
   End If
   
   fill_adress ActiveDocument.CustomDocumentProperties("anrede5").Value, _
               ActiveDocument.CustomDocumentProperties("name5").Value, _
               ActiveDocument.CustomDocumentProperties("Firma5").Value, _
               ActiveDocument.CustomDocumentProperties("strass5").Value, _
               ActiveDocument.CustomDocumentProperties("plz5").Value, _
               ActiveDocument.CustomDocumentProperties("ort5").Value, _
               ActiveDocument.CustomDocumentProperties("ansprache5").Value
   fill_info 5

End Sub
Public Sub Nachrichtlich6()
   
   If ActiveDocument.ProtectionType <> wdNoProtection Then
      MsgBox "Im Anzeigemodus steht die Funktion nicht zur Verfügung"
      Exit Sub
   End If
    
   fill_adress ActiveDocument.CustomDocumentProperties("anrede6").Value, _
               ActiveDocument.CustomDocumentProperties("name6").Value, _
               ActiveDocument.CustomDocumentProperties("Firma6").Value, _
               ActiveDocument.CustomDocumentProperties("strass6").Value, _
               ActiveDocument.CustomDocumentProperties("plz6").Value, _
               ActiveDocument.CustomDocumentProperties("ort6").Value, _
               ActiveDocument.CustomDocumentProperties("ansprache6").Value
   
   fill_info 6

End Sub
Public Sub Nachrichtlich7()
 
   If ActiveDocument.ProtectionType <> wdNoProtection Then
      MsgBox "Im Anzeigemodus steht die Funktion nicht zur Verfügung"
      Exit Sub
   End If
   
   fill_adress ActiveDocument.CustomDocumentProperties("anrede7").Value, _
               ActiveDocument.CustomDocumentProperties("name7").Value, _
               ActiveDocument.CustomDocumentProperties("Firma7").Value, _
               ActiveDocument.CustomDocumentProperties("strass7").Value, _
               ActiveDocument.CustomDocumentProperties("plz7").Value, _
               ActiveDocument.CustomDocumentProperties("ort7").Value, _
               ActiveDocument.CustomDocumentProperties("ansprache7").Value
               
   fill_info 7
      
End Sub
Public Sub fill_info(adr_nr As Variant)
Dim strInfo, tst As String
   
   ' Felder leeren
   ActiveDocument.FormFields("NAMEINFO1").Result = ""
   ActiveDocument.FormFields("NAMEINFO2").Result = ""
   ActiveDocument.FormFields("NAMEINFO3").Result = ""
   ActiveDocument.FormFields("NAMEINFO4").Result = ""
   ActiveDocument.FormFields("NAMEINFO5").Result = ""
   
   
   ' Erstschriftlich ausfüllen
   If (adr_nr > 1) Then
    strInfo = "Erstschriftlich an : "
    ActiveDocument.FormFields("NAMEINFO1").Result = strInfo
    If Trim(ActiveDocument.CustomDocumentProperties("name1").Value) <> "" Then
      strInfo = Trim(ActiveDocument.CustomDocumentProperties("anrede1").Value)
      strInfo = LTrim(strInfo & " " & ActiveDocument.CustomDocumentProperties("name1").Value)
      strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("strass1").Value
      strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("plz1").Value
      strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("ort1").Value
    Else
      strInfo = "An die weiterbehandelnde Kollegin/den weiterbehandelnden Kollegen"
    End If
    ActiveDocument.FormFields("NAMEINFO2").Result = strInfo
    strInfo = ""
   End If
   
   ' Nachrichtlich ausfüllen
   ActiveDocument.FormFields("NAMEINFO3").Result = "Nachrichtlich an : "
   If (adr_nr <> 2) And (Trim(ActiveDocument.CustomDocumentProperties("name2").Value) <> "") Then
    strInfo = Trim(ActiveDocument.CustomDocumentProperties("anrede2").Value)
    strInfo = LTrim(strInfo & " " & ActiveDocument.CustomDocumentProperties("name2").Value)
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("strass2").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("plz2").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("ort2").Value & Chr(13)
   End If
   
   If (adr_nr <> 3) And (Trim(ActiveDocument.CustomDocumentProperties("name3").Value) <> "") Then
    strInfo = strInfo & Trim(ActiveDocument.CustomDocumentProperties("anrede3").Value)
    If (Trim(ActiveDocument.CustomDocumentProperties("anrede3").Value) <> "") Then
        strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("name3").Value
    Else
        strInfo = strInfo & ActiveDocument.CustomDocumentProperties("name3").Value
    End If
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("strass3").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("plz3").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("ort3").Value & Chr$(13)
   End If
   
   If (adr_nr <> 4) And (Trim(ActiveDocument.CustomDocumentProperties("name4").Value) <> "") Then
    strInfo = strInfo & Trim(ActiveDocument.CustomDocumentProperties("anrede4").Value)
    If (Trim(ActiveDocument.CustomDocumentProperties("anrede4").Value) <> "") Then
         strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("name4").Value
    Else
        strInfo = strInfo & ActiveDocument.CustomDocumentProperties("name4").Value
    End If
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("strass4").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("plz4").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("ort4").Value
   End If
   If (Right$(strInfo, 1) = Chr$(13)) Then
      tst = Left$(strInfo, Len(strInfo) - 1)
      strInfo = tst
   End If
   ActiveDocument.FormFields("NAMEINFO4").Result = strInfo
   strInfo = ""
   
   If (adr_nr <> 5) And (RTrim(ActiveDocument.CustomDocumentProperties("name5").Value) <> "") Then
    strInfo = ActiveDocument.CustomDocumentProperties("anrede5").Value
    If (Trim(ActiveDocument.CustomDocumentProperties("anrede5").Value) <> "") Then
        strInfo = LTrim(strInfo & " " & ActiveDocument.CustomDocumentProperties("name5").Value)
    Else
        strInfo = strInfo & ActiveDocument.CustomDocumentProperties("name5").Value
    End If
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("strass5").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("plz5").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("ort5").Value & Chr$(13)
   End If
   
   If (adr_nr <> 6) And (RTrim(ActiveDocument.CustomDocumentProperties("name6").Value) <> "") Then
    strInfo = strInfo & ActiveDocument.CustomDocumentProperties("anrede6").Value
    If (Trim(ActiveDocument.CustomDocumentProperties("anrede6").Value) <> "") Then
       strInfo = LTrim(strInfo & " " & ActiveDocument.CustomDocumentProperties("name6").Value)
    Else
       strInfo = strInfo & ActiveDocument.CustomDocumentProperties("name6").Value
    End If
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("strass6").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("plz6").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("ort6").Value & Chr$(13)
   End If
   
   If (adr_nr <> 7) And (RTrim(ActiveDocument.CustomDocumentProperties("name7").Value) <> "") Then
    strInfo = strInfo & Trim(ActiveDocument.CustomDocumentProperties("anrede7").Value)
    If (Trim(ActiveDocument.CustomDocumentProperties("anrede7").Value) <> "") Then
        strInfo = LTrim(strInfo & " " & ActiveDocument.CustomDocumentProperties("name7").Value)
    Else
       strInfo = strInfo & ActiveDocument.CustomDocumentProperties("name7").Value
    End If
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("strass7").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("plz7").Value
    strInfo = strInfo & " " & ActiveDocument.CustomDocumentProperties("ort7").Value
   End If
   ActiveDocument.FormFields("NAMEINFO5").Result = strInfo

End Sub

Public Sub PrintAll()
Dim pt
On Error GoTo Err_PrintAll
   
    'Dokumentenschutz aufheben, falls vorhanden:
    pt = UnprotectDoc
      
    Erstschriftlich
  
    'Setzen der Druckerschächte
    SetTrays
 
    ActiveDocument.PrintOut
    Nachrichtlich
    ActiveDocument.PrintOut
    If RTrim(ActiveDocument.CustomDocumentProperties("name3").Value) <> "" Then
        Nachrichtlich3
        ActiveDocument.PrintOut
    End If
    If RTrim(ActiveDocument.CustomDocumentProperties("name4").Value) <> "" Then
        Nachrichtlich4
        ActiveDocument.PrintOut
    End If
    If RTrim(ActiveDocument.CustomDocumentProperties("name5").Value) <> "" Then
        Nachrichtlich5
        ActiveDocument.PrintOut
    End If
    If RTrim(ActiveDocument.CustomDocumentProperties("name6").Value) <> "" Then
        Nachrichtlich6
        ActiveDocument.PrintOut
    End If
    If RTrim(ActiveDocument.CustomDocumentProperties("name7").Value) <> "" Then
        Nachrichtlich7
        ActiveDocument.PrintOut
    End If
    Erstschriftlich
    
Exit_PrintAll:

    ProtectDoc (pt)                     'Dokumentenschutz ggf. reaktivieren
    Exit Sub
    
Err_PrintAll:

    Resume Exit_PrintAll

End Sub

Public Sub PrintErstschriftlich()
Dim pt
On Error GoTo Err_PrintErstschriftlich
   
    'Dokumentenschutz aufheben, falls vorhanden:
    pt = UnprotectDoc
   
    'lokale Variablen der Adressierungsfelder setzen
    If Trim(ActiveDocument.CustomDocumentProperties("name1").Value) <> "" Then
    
        fill_adress ActiveDocument.CustomDocumentProperties("anrede1").Value, _
               ActiveDocument.CustomDocumentProperties("name1").Value, _
               ActiveDocument.CustomDocumentProperties("Firma1").Value, _
               ActiveDocument.CustomDocumentProperties("strass1").Value, _
               ActiveDocument.CustomDocumentProperties("plz1").Value, _
               ActiveDocument.CustomDocumentProperties("ort1").Value, _
               ActiveDocument.CustomDocumentProperties("ansprache1").Value
    Else
        ActiveDocument.FormFields("NAMEERST").Result = "An die weiterbehandelnde Kollegin/" & Chr$(13) & "den weiterbehandelnden Kollegen"
        ActiveDocument.FormFields("ANSPRACHE").Result = "Sehr geehrte Kollegin/sehr geehrter Kollege"
        ActiveDocument.FormFields("ANREDEERST").Result = ""
        ActiveDocument.FormFields("STRASSERST").Result = ""
        ActiveDocument.FormFields("PLZERST").Result = ""
        ActiveDocument.FormFields("ORTERST").Result = ""
   
    End If
    
    fill_info 1                         'Daten in das Dokument setzen
    PrintSingle                         'Einzeldruck durchführen

Exit_PrintErstschriftlich:

    ProtectDoc (pt)                     'Dokumentenschutz ggf. reaktivieren
    Exit Sub
    
Err_PrintErstschriftlich:

    Resume Exit_PrintErstschriftlich

End Sub

Public Sub PrintNachrichtlich2()
Dim pt
On Error GoTo Err_PrintNachrichtlich2
    
    'Dokumentenschutz aufheben, falls vorhanden:
    pt = UnprotectDoc
   
    'lokale Variablen der Adressierungsfelder setzen
    fill_adress ActiveDocument.CustomDocumentProperties("anrede2").Value, _
                ActiveDocument.CustomDocumentProperties("name2").Value, _
                ActiveDocument.CustomDocumentProperties("Firma2").Value, _
                ActiveDocument.CustomDocumentProperties("strass2").Value, _
                ActiveDocument.CustomDocumentProperties("plz2").Value, _
                ActiveDocument.CustomDocumentProperties("ort2").Value, _
                ActiveDocument.CustomDocumentProperties("ansprache2").Value

    fill_info 2                         'Daten in das Dokument setzen
    PrintSingle                         'Einzeldruck durchführen

Exit_PrintNachrichtlich2:

    ProtectDoc (pt)                     'Dokumentenschutz ggf. reaktivieren
    Exit Sub
    
Err_PrintNachrichtlich2:

    Resume Exit_PrintNachrichtlich2

End Sub
Public Sub PrintNachrichtlich3()
Dim pt
On Error GoTo Err_PrintNachrichtlich3

    'Dokumentenschutz aufheben, falls vorhanden:
    pt = UnprotectDoc
   
    'lokale Variablen der Adressierungsfelder setzen
    fill_adress ActiveDocument.CustomDocumentProperties("anrede3").Value, _
                ActiveDocument.CustomDocumentProperties("name3").Value, _
                ActiveDocument.CustomDocumentProperties("Firma3").Value, _
                ActiveDocument.CustomDocumentProperties("strass3").Value, _
                ActiveDocument.CustomDocumentProperties("plz3").Value, _
                ActiveDocument.CustomDocumentProperties("ort3").Value, _
                ActiveDocument.CustomDocumentProperties("ansprache3").Value

    fill_info 3                         'Daten in das Dokument setzen
    PrintSingle                         'Einzeldruck durchführen

Exit_PrintNachrichtlich3:

    ProtectDoc (pt)                     'Dokumentenschutz ggf. reaktivieren
    Exit Sub
    
Err_PrintNachrichtlich3:

    Resume Exit_PrintNachrichtlich3

End Sub
Public Sub PrintNachrichtlich4()
Dim pt
On Error GoTo Err_PrintNachrichtlich4

    'Dokumentenschutz aufheben, falls vorhanden:
    pt = UnprotectDoc
   
    'lokale Variablen der Adressierungsfelder setzen
    fill_adress ActiveDocument.CustomDocumentProperties("anrede4").Value, _
                ActiveDocument.CustomDocumentProperties("name4").Value, _
                ActiveDocument.CustomDocumentProperties("Firma4").Value, _
                ActiveDocument.CustomDocumentProperties("strass4").Value, _
                ActiveDocument.CustomDocumentProperties("plz4").Value, _
                ActiveDocument.CustomDocumentProperties("ort4").Value, _
                ActiveDocument.CustomDocumentProperties("ansprache4").Value

    fill_info 4                         'Daten in das Dokument setzen
    PrintSingle                         'Einzeldruck durchführen

Exit_PrintNachrichtlich4:

    ProtectDoc (pt)                     'Dokumentenschutz ggf. reaktivieren
    Exit Sub
    
Err_PrintNachrichtlich4:

    Resume Exit_PrintNachrichtlich4

End Sub
Public Sub PrintNachrichtlich5()
Dim pt
On Error GoTo Err_PrintNachrichtlich5
   
    'Dokumentenschutz aufheben, falls vorhanden:
    pt = UnprotectDoc
   
    'lokale Variablen der Adressierungsfelder setzen
    fill_adress ActiveDocument.CustomDocumentProperties("anrede5").Value, _
                ActiveDocument.CustomDocumentProperties("name5").Value, _
                ActiveDocument.CustomDocumentProperties("Firma5").Value, _
                ActiveDocument.CustomDocumentProperties("strass5").Value, _
                ActiveDocument.CustomDocumentProperties("plz5").Value, _
                ActiveDocument.CustomDocumentProperties("ort5").Value, _
                ActiveDocument.CustomDocumentProperties("ansprache5").Value

    fill_info 5                         'Füllen der Dokumentenfelder
    PrintSingle                         'Einzeldruck durchführen

Exit_PrintNachrichtlich5:

    ProtectDoc (pt)                     'Dokumentenschutz ggf. reaktivieren
    Exit Sub
    
Err_PrintNachrichtlich5:

    Resume Exit_PrintNachrichtlich5

End Sub
Public Sub PrintNachrichtlich6()
Dim pt
On Error GoTo Err_PrintNachrichtlich6

    'Dokumentenschutz aufheben, falls vorhanden:
    pt = UnprotectDoc
   
    'lokale Variablen der Adressierungsfelder setzen
    fill_adress ActiveDocument.CustomDocumentProperties("anrede6").Value, _
                ActiveDocument.CustomDocumentProperties("name6").Value, _
                ActiveDocument.CustomDocumentProperties("Firma6").Value, _
                ActiveDocument.CustomDocumentProperties("strass6").Value, _
                ActiveDocument.CustomDocumentProperties("plz6").Value, _
                ActiveDocument.CustomDocumentProperties("ort6").Value, _
                ActiveDocument.CustomDocumentProperties("ansprache6").Value

    fill_info 6                         'Füllen der Dokumentenfelder
    PrintSingle                         'Einzeldruck durchführen

Exit_PrintNachrichtlich6:

    ProtectDoc (pt)                     'Dokumentenschutz ggf. reaktivieren
    Exit Sub
    
Err_PrintNachrichtlich6:

    Resume Exit_PrintNachrichtlich6

End Sub
Public Sub PrintNachrichtlich7()
Dim pt
On Error GoTo Err_PrintNachrichtlich7

    'Dokumentenschutz aufheben, falls vorhanden:
    pt = UnprotectDoc
   
    'lokale Variablen der Adressierungsfelder setzen
    fill_adress ActiveDocument.CustomDocumentProperties("anrede7").Value, _
                ActiveDocument.CustomDocumentProperties("name7").Value, _
                ActiveDocument.CustomDocumentProperties("Firma7").Value, _
                ActiveDocument.CustomDocumentProperties("strass7").Value, _
                ActiveDocument.CustomDocumentProperties("plz7").Value, _
                ActiveDocument.CustomDocumentProperties("ort7").Value, _
                ActiveDocument.CustomDocumentProperties("ansprache7").Value

    fill_info 7                         'Füllen der Dokumentenfelder
    PrintSingle                         'Einzeldruck durchführen

Exit_PrintNachrichtlich7:

    ProtectDoc (pt)                     'Dokumentenschutz ggf. reaktivieren
    Exit Sub
    
Err_PrintNachrichtlich7:

    Resume Exit_PrintNachrichtlich7

End Sub

Public Sub SetTrays()
On Error GoTo Err_SetTrays
Dim sString As String
' Druckerschächte:
' FS-1800   Kassette1 (First):  258
'           Kassette2 (Other):  259
' FS-2000   Kassette1 (First):  257
'           Kassette2 (Other):  258
'           Univeralz(ufuhr):   259
    Dim druckerName As String
    Dim druckerPort As String
    Dim hPrinter As Long
    Dim bins As Long
    Dim binList As String
    Dim binNum() As Integer
    Dim ManBinNum As Integer
    Dim binString As String

    Dim X As Integer
    Dim strListe() As String
    Dim i As Long
    Dim ManText, Fach1, Fach1Text, Fach2, Fach2Text
    Const DC_BINS = 6
    Const DC_BINNAMES = 12
    Const suchText As String = " on "
    sString = ActivePrinter
'   aktiven Drucker auslesen und anschließend die möglichen Trays und Papiersorten -> bins
    On Error Resume Next
    'Stop
    If InStr(sString, suchText) > 3 Then
        druckerName = Left(sString, InStr(1, sString, suchText) - 1)
        druckerPort = Right(sString, Len(sString) - Len(druckerName) - Len(suchText))
    Else
        ' Office 14 liefert nur noch den UNC Namen
        druckerPort = sString
        druckerName = sString
    End If
    bins = DeviceCapabilities(druckerName, druckerPort, DC_BINS, ByVal vbNullString, 0)
    ReDim binNum(1 To bins)
    bins = DeviceCapabilities(druckerName, druckerPort, DC_BINS, binNum(1), 0)
    binList = String$(24 * bins, 0)
    bins = DeviceCapabilities(druckerName, druckerPort, DC_BINNAMES, ByVal binList, 0)
    ReDim strListe(1 To bins, 2)
    For X = 1 To bins
        binString = Mid(binList, 24 * (X - 1) + 1, 24)
        binString = Left(binString, InStr(1, binString, Chr(0)) - 1)
        strListe(X, 0) = binString
        strListe(X, 1) = (binNum(X))
'   die eingetragene Tray-Nummer auslesen für Manuelle Zufuhr
        If binString = "Massenzufuhr" Or _
            Left(binString, 12) = " Man. Zufuhr" Or _
            Left(binString, 10) = "Multi Purp" Or _
            Left(binString, 15) = "Universalzufuhr" Then
            ManBinNum = (binNum(X))
            ManText = binString
'   die eingetragene Tray-Nummer auslesen für Fach 1
        ElseIf binString = " Fach 1" Or Left(binString, 10) = "Kassette 1" Or binString = "Cassette 1" Then
            Fach1 = (binNum(X))
            Fach1Text = binString
'   die eingetragene Tray-Nummer auslesen für Fach 2
        ElseIf binString = " Fach 2" Or Left(binString, 10) = "Kassette 2" Or binString = "Cassette 2" Then
            Fach2 = (binNum(X))
            Fach2Text = binString
        End If
    Next X
   
'   Falls kein Fach 2 vorhanden auf manuelle Zufuhr setzen für Logopapier
'   Obsolet: jetzt immer Fach 1 verwenden, wenn kein zweites Fach vorhanden ist,
'   z. B. wegen Arztdrucker ITS
    If Fach2 = "" Then
        Fach2 = Fach1
        ' alle Abschnitte durchlaufen
'        For i = 1 To ActiveDocument.Sections.count
'            ' nur im ersten Abschnitt die Schächte unterscheiden
'            If i = 1 Then
'                ActiveDocument.Sections(i).PageSetup.FirstPageTray = Fach1
'            ' ab zweitem Abschnitt wird immer der 2. Schacht angesteuert
'            Else
'                ActiveDocument.Sections(i).PageSetup.FirstPageTray = ManBinNum
'            End If
'            ActiveDocument.Sections(i).PageSetup.OtherPagesTray = ManBinNum
'        Next i
'        y = MsgBox("Bitte Logopapier in " & ManText & " legen und " & vbCrLf & "Standardpapier in " & Fach1Text & vbCrLf & "an Drucker " & druckerName, vbOKOnly + vbExclamation)
    Else
'   Sonst Logopapier in Fach 1 und Standard in Fach 2
        ' alle Abschnitte durchlaufen
        For i = 1 To ActiveDocument.Sections.Count
            ' nur im ersten Abschnitt die Schächte unterscheiden
            If i = 1 Then
                ActiveDocument.Sections(i).PageSetup.FirstPageTray = Fach1
            ' ab zweitem Abschnitt wird immer der 2. Schacht angesteuert
            Else
                ActiveDocument.Sections(i).PageSetup.FirstPageTray = Fach2
            End If
            ActiveDocument.Sections(i).PageSetup.OtherPagesTray = Fach2
        Next i
 '       y = MsgBox("Bitte Logopapier in " & Fach2Text & " legen und " & vbCrLf & "Standardpapier in " & Fach1Text & vbCrLf & "an Drucker " & druckerName, vbOKOnly + vbExclamation)
    End If
Exit_SetTrays:
    Exit Sub
    
Err_SetTrays:

    MsgBox "Fehler beim Initialisieren der Druckerschächte!"
    Resume Exit_SetTrays
End Sub

Public Sub PrintSingle()
On Error GoTo Err_PrintSingle

    'Setzen der Druckerschächte
    SetTrays
    'Dokument "wie es ist" ausdrucken
    ActiveDocument.PrintOut
    
Exit_PrintSingle:

    Exit Sub
    
Err_PrintSingle:

    MsgBox "Fehler beim Drucken eines Einzeldokuments!"
    Resume Exit_PrintSingle
    
End Sub

Public Function UnprotectDoc() As Long
On Error GoTo Err_UnprotectDoc

    Dim pt As Long

    pt = ActiveDocument.ProtectionType  'akt. Dokumentenschutzstatus merken
    'Dokumentenschutz aufheben, falls vorhanden:
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect
    End If
    UnprotectDoc = pt
    
Exit_UnprotectDoc:

    Exit Function
    
Err_UnprotectDoc:

    MsgBox "Fehler beim Einzeldruck!"
    Resume Exit_UnprotectDoc

End Function

Public Sub ProtectDoc(pt As Long)
On Error GoTo Err_ProtectDoc
    
    If pt <> wdNoProtection Then
        ActiveDocument.Protect Type:=pt 'Dokumentenschutz setzen
    End If

Exit_ProtectDoc:

    Exit Sub
    
Err_ProtectDoc:

    MsgBox "Fehler beim Einzeldruck!"
    Resume Exit_ProtectDoc

End Sub
Sub Initialize()
'
' Initialize Makro
' Makro erstellt am 30.09.2004 von KD
'
     If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        Dim p As Word.Pane, firstpane As Boolean
        firstpane = True
        ' Alle Panes ausser der ersten schliessen
        For Each p In Application.ActiveDocument.ActiveWindow.Panes
            If firstpane Then
                firstpane = False
            Else
                p.Close
            End If
        Next
     End If
     ' Viewtype auf PageView umstellen
     If ActiveWindow.ActivePane.View.Type <> wdPageView Then
        ActiveWindow.ActivePane.View.Type = wdPageView
     End If
     If ActiveDocument.ProtectionType <> wdNoProtection Then
        DisableCommandBarEntry ("Drucken")
    End If
    'ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitBestFit
    
End Sub







