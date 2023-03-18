Imports System.Collections
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint


Partial Public Class MyFunctions

    Public Class Slides

        Public Shared Sub ChangeSlidesFormat(ByVal filePath As String)

            Dim pptApp As PowerPoint.Application
            pptApp = New PowerPoint.Application

            Dim ppt As Presentation
            ppt = pptApp.Presentations.Open(filePath)

            ppt.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen16x9

            ppt.Save()
            ppt.Close()

        End Sub

        Public Shared Sub ChangeFolderSlidesFormat(ByVal folderPath As String)

            Dim fileNames() As String
            fileNames = Files.GetFileNames(folderPath)

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

            'If (Not srcPath Like "*.pptx") Or (Not srcPath Like "?:\*") Then
            'errorMsg = "Dateeinde muss "".pptx"" sein" & vbNewLine & "Dateipfad muss mit Laufwerksbuchstaben anfangen: z.B. ""C:\"""
            'End If

            'If Not srcPath Like "*.pptx" Then
            'errorMsg = "Dateiende muss "".pptx"" sein"
            'End If
            'If Not srcPath Like "?:\*" Then
            'errorMsg = "Dateipfad muss mit Laufwerksbuchstaben anfangen z.B. ""C:\"""
            'End If

            If Files.FileExists(srcPath) = True Then

                ' Set PPT object variables
                Dim pptApp As PowerPoint.Application
                Dim destPPT As Presentation
                Dim srcPPT As Presentation

                pptApp = Globals.ThisAddIn.Application

                destPPT = pptApp.ActivePresentation
                srcPPT = pptApp.Presentations.Open(srcPath, MsoTriState.msoTrue,, MsoTriState.msoFalse)

                'copy Slides
                srcPPT.Slides.Range().Copy()
                pptApp.CommandBars.ExecuteMso("PasteSourceFormatting")

                System.Windows.Forms.Application.DoEvents()

                'https://stackoverflow.com/a/20450573/20785822
                'https://stackoverflow.com/a/69303122/20785822
                'destPPT.Save()

                'close opened PPT
                srcPPT.Close()

            Else

                MsgBox("Datei " & srcPath & " konnte nicht geöffnet werden")

            End If

        End Sub

        Public Shared Sub InsertAnnouncement(songNumber As String)

            Dim path As String = appDataPath & "Main.pptx"
            InsertSlides(path)

            Dim slide As Slide
            Dim ogFont As PowerPoint.Font
            Dim textbox As PowerPoint.Shape

            Dim width As Single
            Dim height As Single
            Dim x As Single
            Dim y As Single

            With Globals.ThisAddIn.Application
                slide = .ActiveWindow.View.Slide

                width = 200
                height = 50

                With .ActivePresentation.PageSetup
                    x = .SlideWidth - width * 1.5
                    y = height * 1
                End With
            End With

            textbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, x, y, width, height)

            With textbox.TextFrame.TextRange

                .Text = songNumber

                Dim tempShape As PowerPoint.Shape
                Dim i As Integer = 1
                For Each tempShape In slide.Shapes

                    If tempShape.TextFrame.HasText = MsoTriState.msoTrue Then
                        If tempShape.TextFrame.TextRange.Text.Length > 0 Then
                            Exit For
                        End If
                    End If

                    i += 1
                Next

                ogFont = slide.Shapes.Item(i).TextFrame.TextRange.Font
                'MsgBox(slide.Shapes.Item(i).TextFrame.TextRange.Text)

                With .Font

                    .Size = ogFont.Size
                    .Name = ogFont.Name
                    .Italic = False
                    .Bold = True
                    .Color.RGB = ogFont.Color.RGB

                End With

            End With

        End Sub

        Public Shared Sub InsertVideo(filePath As String)

            Dim ppt As PowerPoint.Presentation
            Dim current As PowerPoint.Slide
            Dim slide As PowerPoint.Slide
            Dim videoShape As PowerPoint.Shape

            Dim width As Single
            Dim height As Single

            ppt = Globals.ThisAddIn.Application.ActivePresentation
            current = ppt.Application.ActiveWindow.View.Slide
            slide = ppt.Slides.AddSlide(current.SlideNumber + 1, ppt.SlideMaster.CustomLayouts.Item(1))

            width = ppt.PageSetup.SlideWidth
            height = ppt.PageSetup.SlideHeight

            videoShape = slide.Shapes.AddMediaObject2(filePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, width, height)
            videoShape.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue

        End Sub

    End Class

End Class