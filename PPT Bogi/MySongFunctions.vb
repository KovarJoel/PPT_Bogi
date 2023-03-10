Option Explicit On

Imports Microsoft.Office.Interop.PowerPoint
Imports PPT_Bogi.MyFunctions
Imports Scripting

Public Class MySongFunctions

    Private Shared Sub SetSongsPath(ByVal path As String)

        Dim errorMsg As String

        If Not path Like "?:*" Then
            errorMsg = "Dateipfad muss mit Laufwerksbuchstaben anfangen z.B. ""C:\"""
            MsgBox(errorMsg)
            Exit Sub
        End If

        If Not path Like "*\" Then
            path &= "\"
        End If

        Dim filePath As String
        filePath = Environ("AppData") & "\pptMacro"

        If Not DirectoryExists(filePath) Then
            MkDir(filePath)
        End If

        filePath &= "\songLocationFile.txt"
        If Not FileExists(filePath) Then
            CreateTextFile(filePath)
        End If

        Dim fso As FileSystemObject
        fso = New FileSystemObject
        Dim fileStream As TextStream

        fileStream = fso.OpenTextFile(filePath, IOMode.ForWriting)
        fileStream.Write(path)
        fileStream.Close()

    End Sub

    Private Shared Function GetSongsPath() As String

        Dim appDataPath As String
        appDataPath = Environ("AppData")

        Dim filePath As String
        filePath = appDataPath & "\pptMacro\songLocationFile.txt"

        If Not FileExists(filePath) Then
            Return Nothing
        End If

        Dim fso As FileSystemObject
        fso = New FileSystemObject

        Dim file As TextStream
        file = fso.OpenTextFile(filePath, IOMode.ForReading)

        Dim data As String
        data = file.ReadLine()

        file.Close()

        GetSongsPath = data

    End Function

    Public Shared Sub ChangeSongsDirectory()

        Dim inputString As String

        inputString = InputBox("Lieder Verzeichnis: " & vbNewLine & "(z.B. C:\Users\Max\Documents\Lieder)", "Verzeichnis ändern")

        If inputString = Nothing Then
            Exit Sub
        ElseIf Not DirectoryExists(inputString) Then
            MsgBox("Verzeichnis konnte nicht gefunden werden")
        Else
            SetSongsPath(inputString)
        End If

    End Sub

    Public Shared Sub InsertSong()

        Dim songsPath As String
        songsPath = GetSongsPath()

        If songsPath = Nothing Or Not DirectoryExists(songsPath) Then
            ChangeSongsDirectory()
            songsPath = GetSongsPath()
        End If
        If songsPath = Nothing Or Not DirectoryExists(songsPath) Then
            Exit Sub
        End If

        ' set filePath
        Dim songName As String
        Dim srcString As String
        Dim inputString As String

        ' get song nr as input
        inputString = InputBox("Lied Nummer:" & vbNewLine & "(z.B. 187)")

        If Not inputString = Nothing And Not songsPath = Nothing Then
            ' concatenate strings
            songName = GetSongName(inputString, songsPath)
            If songName = Nothing Then
                MsgBox("Lied konnte nicht gefunden werden")
                Exit Sub
            Else
                srcString = songsPath & songName
            End If
            ' run insert macro
            InsertSlides(srcString)
        End If

    End Sub

    Private Shared Function GetSongName(number As String, folderPath As String) As String

        Dim fileNames() As String
        fileNames = GetFileNames(folderPath)

        Dim songName As String
        songName = Nothing
        For Each songName In fileNames
            If songName Like "*[!1-9][!1-9]" & number & "[!0-9]*" And songName Like "*.ppt*" Then
                Exit For
            End If
        Next

        If songName Like "*.ppt*" Then
            GetSongName = songName
        Else
            GetSongName = Nothing
        End If

    End Function

End Class
