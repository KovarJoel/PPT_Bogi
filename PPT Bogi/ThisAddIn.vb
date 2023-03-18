Imports System.IO
Imports Scripting

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

        Dim tempFile As String

        For Each tempFile In Directory.GetFiles(MyFunctions.tempDataPath)
            If MyFunctions.Files.FileExists(tempFile) Then
                System.IO.File.Delete(tempFile)
            End If
        Next

    End Sub

End Class
