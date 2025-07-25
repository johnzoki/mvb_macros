Attribute VB_Name = "SharepointMacros"
Function GetUserFolder() As String
    ' Returns Windows User path
    
    GetUserFolder = Environ("USERPROFILE")
    
End Function

Function GetPresetFolder() As String
    ' Define Value "havePresetFolder" as True or False
    Dim havePresetFolder As Boolean

    ' Define "havePresetFolder" decides if your the given path in "presetPath" starts from the path of the given Preset-Folder or the Sharepoint Home-Directory
    havePresetFolder = True
    ' If "havePresetFolder" is True set "presetFolder" to the path of the Preset-Folder starting from the Sharepoint Home Directory
    presetFolder = "\sharepoint\preset\folder\path\"

    If havePresetFolder Then
        GetPresetFolder = presetFolder
    Else
        GetPresetFolder = "\"
    End If
    
End Function


Function Error_FileNotFound()
    ' Returns Error-Message if the file doesnt exists.
    MsgBox "Die Datei existiert nicht. Ueberpruefe den angegebenen Datei-Pfad der Funktion in Visual Basic oder kontaktiere deine EDV.  LG John", vbExclamation, "Dateipfad Ueberpruefung"

End Function


Function GetPathSharePointFile(ByVal presetPath As String) As String
    ' Returns complete path to preset-file by adding the user and preset folder to the given path in function "OpenDocument".

    userFolder = GetUserFolder()
    presetFolder = GetPresetFolder()
    filePath = userFolder & presetFolder & presetPath
    
    GetPathSharePointFile = filePath

End Function


Function OpenDocument(ByVal presetPath As String)
    ' Tests if the file can be opened.
    ' If not, it will run a function that opens a Error-Message.
    
    filePath = GetPathSharePointFile(presetPath)
    
    FileName = Dir(filePath)

    If FileName <> "" Then
        ' File is valid
        Documents.Add filePath
    Else
        ' File is not valid
        Error_FileNotFound
    End If

End Function


Sub Vorlage_example()

    Dim presetPath As String
    
    ' If in Function "GetPresetFolder" "havePresetFolder" is True then only use the Path starting from Preset-Folder. Otherwise start from Sharepoint directory.
    ' Insert filepath from Sharepoint/Preset-Folder here:
    presetPath = "PresetFolderOnSharepoint\your\word\preset\file\path\example.dotm"
    
    OpenDocument presetPath
    
End Sub