Attribute VB_Name = "SharepointMacros"
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


Sub Vorlage_example()

    Dim sharepointPath As String
    
    ' Dateipfad Sharepoint hier eintippen:
    sharepointPath = "\Sharepoint\your\word\preset\file\path\example.dotm"
    
    OpenDocument sharepointPath
    
End Sub