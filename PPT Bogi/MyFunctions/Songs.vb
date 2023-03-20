Option Explicit On

Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Scripting

Partial Public Class MyFunctions
    Public Class Songs

        Public Shared announce As Boolean = False
        Public Shared directoryPath As String

        Private Shared Sub SetSongsPath(ByVal path As String)

            If Not path Like "*\" Then
                path &= "\"
            End If

            Dim filePath As String
            filePath = appDataPath

            If Not Files.DirectoryExists(filePath) Then
                MkDir(filePath)
            End If

            filePath &= "songLocationFile.txt"
            If Not Files.FileExists(filePath) Then
                Files.CreateTextFile(filePath)
            End If

            Dim fso As FileSystemObject
            fso = New FileSystemObject
            Dim fileStream As TextStream

            fileStream = fso.OpenTextFile(filePath, IOMode.ForWriting)
            fileStream.Write(path)
            fileStream.Close()

            directoryPath = path

        End Sub

        Public Shared Function GetSongsPath() As String

            Dim filePath As String
            filePath = appDataPath & "songLocationFile.txt"

            If Not Files.FileExists(filePath) Then
                Return Nothing
            End If

            Dim fso As FileSystemObject
            fso = New FileSystemObject

            Dim file As TextStream
            file = fso.OpenTextFile(filePath, IOMode.ForReading)

            Dim data As String
            data = file.ReadLine()

            file.Close()
            directoryPath = data
            GetSongsPath = data

        End Function

        Public Shared Sub ChangeSongsDirectory()

            Dim result As DialogResult
            Dim dialog As New OpenFileDialog
            Dim path As String

            With dialog

                If Files.DirectoryExists(directoryPath) Then
                    .InitialDirectory = directoryPath
                Else
                    .InitialDirectory = Environ("UserProfile") & "\Documents"
                End If

                .Title = "Lied auswählen (Verzeichnis wird automatisch gewählt)"

                .ValidateNames = True
                .CheckFileExists = True
                .CheckPathExists = True
                .Multiselect = False
            End With

            result = dialog.ShowDialog()

            If result = Windows.Forms.DialogResult.OK Then
                path = dialog.FileName
                path = path.Substring(0, path.LastIndexOf("\"))
                SetSongsPath(path)
            End If

        End Sub

        Public Shared Sub InsertSong()

            If directoryPath = Nothing Or Not Files.DirectoryExists(directoryPath) Then
                ChangeSongsDirectory()
                directoryPath = GetSongsPath()
            End If
            If directoryPath = Nothing Or Not Files.DirectoryExists(directoryPath) Then
                Exit Sub
            End If

            ' set filePath
            Dim songName As String
            Dim srcString As String
            Dim inputString As String

            ' get song nr as input
            inputString = InputBox("Lied Nummer:" & vbNewLine & "(z.B. 187)")

            If Not inputString = Nothing And Not directoryPath = Nothing Then
                ' concatenate strings
                songName = GetSongName(inputString, directoryPath)
                If songName = Nothing Then
                    MsgBox("Lied konnte nicht gefunden werden")
                    Exit Sub
                Else
                    srcString = directoryPath & songName
                End If
                ' run insert macro

                If announce = True Then
                    Slides.InsertAnnouncement(inputString)
                End If

                Slides.InsertSlides(srcString)
            End If

        End Sub

        Public Shared Sub InsertSongManually()

            Dim srcPath As String
            Dim number As String
            Dim dialog As New OpenFileDialog
            Dim result As DialogResult

            With dialog
                .CheckFileExists = True
                .CheckPathExists = True
                .Multiselect = False
                .Title = "Lied einfügen"

                If Files.DirectoryExists(directoryPath) Then
                    .InitialDirectory = directoryPath
                Else
                    .InitialDirectory = Environ("UserProfile") & "\Documents"
                End If

            End With

            result = dialog.ShowDialog()

            If result = DialogResult.OK Then
                srcPath = dialog.FileName

                If announce = True Then

                    Dim indexFirst As Integer = 0
                    Dim indexLast As Integer = 0
                    Dim letter As Char

                    Dim i As Integer = 0
                    For Each letter In srcPath
                        If Char.IsDigit(letter) = True Then
                            If i > srcPath.LastIndexOf("\") And indexFirst = 0 Then
                                indexFirst = i
                            End If
                            indexLast = i
                        End If
                        i += 1
                    Next

                    number = srcPath.Substring(indexFirst, indexLast - indexFirst + 1)

                    If Regex.IsMatch(number, "^[0-9 ]+$") Then

                        While number Like "0*"
                            number = number.Substring(1)
                        End While

                        Slides.InsertAnnouncement(number)

                    End If

                End If

                Slides.InsertSlides(srcPath)

            Else
                Exit Sub
            End If

        End Sub

        Private Shared Function GetSongName(number As String, folderPath As String) As String

            Dim fileNames() As String
            fileNames = Files.GetFileNames(folderPath)

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

End Class