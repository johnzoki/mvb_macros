Attribute VB_Name = "NewMacros"
Function GetUserFolder() As String
    ' Gibt den Pfad zum Benutzerordner zurück
    
    GetUserFolder = Environ("USERPROFILE")
    
End Function


Function Error_FileNotFound()
    ' Gibt Error-Nachricht zurück, wenn eine Datei nicht existiert.
    MsgBox "Die Datei existiert nicht. Überprüfe den angegebenen Datei-Pfad der Funktion in Visual Basic oder kontaktiere deine EDV.  LG John", vbExclamation, "Dateipfad Überprüfung"

End Function


Function GetPathSharePointFile(ByVal sharepointPath As String) As String
    ' Gibt den Pfad zur File auf dem SharePoint zurück
    
    userFolder = GetUserFolder()
    filePath = userFolder & sharepointPath
    
    GetPathSharePointFile = filePath

End Function


Function OpenDocument(ByVal sharepointPath As String)
    ' Testet, ob die Datei geöffnet werden kann. Je nachdem wird die Datei dann geöffnet, sonst gibt es eine Error-Nachricht
    
    filePath = GetPathSharePointFile(sharepointPath)
    
    FileName = Dir(filePath)

    If FileName <> "" Then
        ' Datei existiert
        Documents.Add filePath
    Else
        ' Datei existiert nicht
        Error_FileNotFound
    End If

End Function


Sub Vorlage_A4_hoch_leer()

    Dim sharepointPath As String
    
    ' Dateipfad Sharepoint hier eintippen:
    sharepointPath = "\Axess Architekten AG\100_Büro Sharepoint - Dokumente\02 Vorlagen\3_Excel_Word\Word\Vorlage_A4_hoch_leer.dotm"
    
    OpenDocument sharepointPath
    
End Sub

Sub Vorlage_Kostenschaetzung()

    Dim sharepointPath As String
    
    ' Dateipfad Sharepoint hier eintippen:
    sharepointPath = "\Axess Architekten AG\100_Büro Sharepoint - Dokumente\02 Vorlagen\2_Verträge_Bauleitung_Bauadministration\Kostenschätzung\Vorlage_Kostenschätzung-mit-Inhalt.dotm"
    
    OpenDocument sharepointPath

End Sub

Sub Vorlage_Brief()

    Dim sharepointPath As String
    
    ' Dateipfad Sharepoint hier eintippen:
    sharepointPath = "\Axess Architekten AG\100_Büro Sharepoint - Dokumente\02 Vorlagen\3_Excel_Word\Word\Vorlage_Brief.dotm"
    
    OpenDocument sharepointPath

End Sub


Sub Vorlage_Lieferschein()

    Dim sharepointPath As String
    
    ' Dateipfad Sharepoint hier eintippen:
    sharepointPath = "\Axess Architekten AG\100_Büro Sharepoint - Dokumente\02 Vorlagen\3_Excel_Word\Word\Vorlage_Lieferschein.dotm"
    
    OpenDocument sharepointPath

End Sub


Sub Vorlage_Baubeschrieb()

    Dim sharepointPath As String
    
    ' Dateipfad Sharepoint hier eintippen:
    sharepointPath = "\Axess Architekten AG\100_Büro Sharepoint - Dokumente\02 Vorlagen\2_Verträge_Bauleitung_Bauadministration\Baubeschrieb\Vorlage_Baubeschrieb.dotm"
    
    OpenDocument sharepointPath

End Sub
