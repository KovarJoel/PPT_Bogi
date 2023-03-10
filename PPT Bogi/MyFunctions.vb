Option Explicit On

Imports System.IO
Imports Microsoft.Office.Interop.PowerPoint
Imports Scripting

Public Class MyFunctions

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

    Public Shared Sub ChangeSlidesFormat(ByVal filePath As String)

        Dim pptApp As Application
        pptApp = New Application

        Dim ppt As Presentation
        ppt = pptApp.Presentations.Open(filePath)

        ppt.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen16x9

        ppt.Save()
        ppt.Close()

    End Sub

    Public Shared Sub ChangeFolderSlidesFormat(ByVal folderPath As String)

        Dim fileNames() As String
        fileNames = GetFileNames(folderPath)

        Dim fileName As String

        If folderPath Like "*[!/\]" Then
            folderPath &= "\"
        End If

        For Each fileName In fileNames
            If fileName Like "*.ppt*" Then
                ChangeSlidesFormat(folderPath & fileName)
            End If
        Next

    End Sub

    Public Shared Sub InsertSlides(srcPath As String)

        Dim errorMsg As String
        errorMsg = Nothing

        If (Not srcPath Like "*.pptx") Or (Not srcPath Like "?:\*") Then
            errorMsg = "Dateeinde muss "".pptx"" sein" & vbNewLine & "Dateipfad muss mit Laufwerksbuchstaben anfangen: z.B. ""C:\"""
        End If

        If Not srcPath Like "*.pptx" Then
            errorMsg = "Dateiende muss "".pptx"" sein"
        End If
        If Not srcPath Like "?:\*" Then
            errorMsg = "Dateipfad muss mit Laufwerksbuchstaben anfangen z.B. ""C:\"""
        End If

        If FileExists(srcPath) = True And Not errorMsg Then

            ' Set PPT object variables
            Dim pptApp As Application
            Dim destPPT As Presentation
            Dim srcPPT As Presentation
            pptApp = GetObject([Class]:="PowerPoint.Application")
            'pptApp = New PowerPoint.Application
            destPPT = pptApp.ActivePresentation
            srcPPT = pptApp.Presentations.Open(srcPath)

            'copy Slides
            srcPPT.Slides.Range().Copy()
            destPPT.Windows.Item(1).Activate()
            destPPT.Application.CommandBars.ExecuteMso("PasteSourceFormatting")

            'close opened PPT
            srcPPT.Close()

        Else

            MsgBox("Datei " & srcPath & " konnte nicht geöffnet werden" & vbNewLine & vbNewLine & errorMsg)

        End If

    End Sub

End Class
