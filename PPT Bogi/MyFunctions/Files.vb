Option Explicit On

Imports System.Collections
Imports System.IO
Imports System.Net
Imports System.Net.Cache
Imports System.Runtime.Remoting.Channels
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Scripting

Partial Public Class MyFunctions

    Public Shared appDataPath As String = Environ("AppData") & "\pptMacro\"
    Public Shared tempDataPath As String = Environ("AppData") & "\pptMacro\" & "temp\"

    Public Class Files
        Public Shared Function DirectoryExists(ByVal DirectorySpec As String) As Boolean

            Dim Attr As Long
            On Error Resume Next
            Attr = GetAttr(DirectorySpec)

            If Err.Number = 0 Then
                DirectoryExists = ((Attr And vbDirectory) = vbDirectory)
            End If

        End Function

        Public Shared Function FileExists(ByVal DirectorySpec As String) As Boolean

            Dim Attr As Long
            On Error Resume Next
            Attr = GetAttr(DirectorySpec)

            If Err.Number = 0 Then
                FileExists = Not ((Attr And vbDirectory) = vbDirectory)
            End If

        End Function

        Public Shared Sub CreateTextFile(ByVal filePath As String)

            Dim fso As FileSystemObject
            fso = New FileSystemObject
            Dim fileStream As TextStream

            fileStream = fso.CreateTextFile(filePath)
            fileStream.Close()

        End Sub

        Public Shared Function GetFileNames(folderPath As String) As String()

            Dim i As Integer
            Dim fileNames() As String

            Dim fso As FileSystemObject
            fso = New FileSystemObject

            Dim folder As Object
            folder = fso.GetFolder(folderPath)

            Dim files As Object
            files = folder.files

            If files.Count = 0 Then
                Return Nothing
            End If

            Dim file As Object
            ReDim fileNames(0 To files.Count)
            i = 0
            For Each file In files
                fileNames(i) = file.name
                i += 1
            Next

            GetFileNames = fileNames

        End Function

        Public Shared Sub LoadAllResources()

            Dim path As String = appDataPath
            IO.File.WriteAllBytes(path & "EGW.pptx", My.Resources.EGW)
            IO.File.WriteAllBytes(path & "Main.pptx", My.Resources.Main)
            IO.File.WriteAllBytes(path & "Missionskarte.pptx", My.Resources.Missionskarte)
            IO.File.WriteAllBytes(path & "Sabbatschulgruppen.pptx", My.Resources.Sabbatschulgruppen)

        End Sub

        Public Shared Sub LoadFileFromLink(URL As String, filePath As String)

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            If DirectoryExists(tempDataPath) = False Then
                MkDir(tempDataPath)
            End If

            Try
                My.Computer.Network.DownloadFile(URL, filePath, "", "", True, 500, True)

            Catch ex As Exception

                MsgBox("Download fehlgeschlagen")

                If IO.File.Exists(filePath) Then
                    IO.File.Delete(filePath)
                End If

                Exit Sub

            End Try

        End Sub

    End Class

End Class