Imports System.Collections
Imports System.Drawing
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Windows.Forms.LinkLabel
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
            Dim textBox As PowerPoint.Shape

            Dim width As Single
            Dim height As Single
            Dim x As Single
            Dim y As Single

            With Globals.ThisAddIn.Application
                slide = .ActiveWindow.View.Slide

                width = 300
                height = 60

                With .ActivePresentation.PageSetup
                    x = .SlideWidth - width - 25
                    y = height - 10
                End With
            End With

            textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, x, y, width, height)

            With textBox.TextFrame.TextRange

                Dim ogFont As PowerPoint.Font = Nothing

                For Each tempShape As PowerPoint.Shape In slide.Shapes
                    If tempShape.HasTextFrame Then
                        If tempShape.TextFrame.HasText = MsoTriState.msoTrue Then
                            If tempShape.TextFrame.TextRange.Text.Length > 0 Then
                                ogFont = tempShape.TextFrame.TextRange.Lines(0, 1).Font
                                Exit For
                            End If
                        End If
                    End If
                Next

                .Text = songNumber
                .Font.Bold = True
                .Font.Italic = False
                .Font.Size = 36

                If ogFont IsNot Nothing Then
                    .Font.Color.RGB = ogFont.Color.RGB
                Else
                    .Font.Color.RGB = Color.Black.ToArgb
                End If

                .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight

            End With

        End Sub

        Public Shared Sub InsertClosing()

            Dim slide As PowerPoint.Slide
            Dim textBox As PowerPoint.Shape
            Dim ogFont As PowerPoint.Font = Nothing

            Dim x, y, width, height As Single

            InsertSlides(appDataPath & "Main.pptx")

            slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide

            With slide.Application.ActivePresentation.PageSetup

                width = .SlideWidth * 0.9
                height = 50
                x = (.SlideWidth - width) / 2
                y = height / 2

            End With

            textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, x, y, width, height)

            textBox.TextFrame.TextRange.Text = "Es wird um andächtige Stille auch nach dem Ende des Gottesdienstes " _
                & "gebeten. Herzlichen Dank und gesegneten Sabbat!"

            For Each shape As PowerPoint.Shape In slide.Shapes
                If shape.HasTextFrame Then
                    If shape.TextFrame.HasText Then
                        If shape.TextFrame.TextRange.Text.Length > 0 Then
                            ogFont = shape.TextFrame.TextRange.Lines(0, 1).Font
                            Exit For
                        End If
                    End If
                End If
            Next

            With textBox.TextFrame.TextRange.Font

                If ogFont IsNot Nothing Then
                    .Color.RGB = ogFont.Color.RGB
                Else
                    .Color.RGB = Color.Black.ToArgb()
                End If

                .Bold = True
                .Italic = False
                .Shadow = True
                .Size = 28

            End With

            textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter

        End Sub

        Public Shared Sub InsertVideo()

            Dim ppt As PowerPoint.Presentation
            Dim current As PowerPoint.Slide
            Dim slide As PowerPoint.Slide
            Dim videoShape As PowerPoint.Shape

            Dim width As Single
            Dim height As Single

            Dim URL As String
            Dim filePath As String = tempDataPath & "video.mp4"

            URL = InputBox("Downloadlink zu Video einfügen: ", "Missionsvideo")
            If URL = Nothing Then
                Exit Sub
            End If
            Files.LoadFileFromLink(URL, filePath)

            If Not IO.File.Exists(filePath) Then
                MsgBox("Datei konnte nicht geöffnet werden")
                Exit Sub
            End If

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