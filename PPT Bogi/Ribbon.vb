Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon

    Private Sub Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        MyFunctions.Files.LoadAllResources()
        MyFunctions.Songs.GetSongsPath()

    End Sub

    Private Sub ButtonInsertSongManually_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonInsertSongManually.Click

        MyFunctions.Songs.InsertSongManually()

    End Sub

    Private Sub ButtonInsertSong_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonInsertSong.Click

        MyFunctions.Songs.InsertSong()

    End Sub

    Private Sub ButtonChangeDirectory_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonChangeDirectory.Click

        MyFunctions.Songs.ChangeSongsDirectory()

    End Sub

    Private Sub CheckBoxAnnounce_Click(sender As Object, e As RibbonControlEventArgs) Handles CheckBoxAnnounce.Click

        MyFunctions.Songs.announce = Not MyFunctions.Songs.announce

    End Sub

    Private Sub ButtonInsertMap_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonInsertMap.Click

        Dim path As String = MyFunctions.appDataPath & "Missionskarte.pptx"
        MyFunctions.Slides.InsertSlides(path)

    End Sub

    Private Sub ButtonInsertLocations_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonInsertLocations.Click

        Dim path As String = MyFunctions.appDataPath & "Sabbatschulgruppen.pptx"
        MyFunctions.Slides.InsertSlides(path)

    End Sub

    Private Sub ButtonInsertEGW_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonInsertEGW.Click

        Dim path As String = MyFunctions.appDataPath & "EGW.pptx"
        MyFunctions.Slides.InsertSlides(path)

    End Sub

    Private Sub ButtonInsertClosing_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonInsertClosing.Click

        MyFunctions.Slides.InsertClosing()

    End Sub

    Private Sub ButtonInsertVideo_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonInsertVideo.Click

        ' https://cloud.eud.adventist.org/index.php/s/i9nTwt55bHEmpLJ?path=%2F2023_1.%20Quartal
        ' https://cloud.eud.adventist.org/index.php/s/ZiLmkiKZgN262Fq/download/01_Mehr%20Wohnungen%20auf%20dem%20Campus.mp4

        MyFunctions.Slides.InsertVideo()

    End Sub

    Private Sub ButtonSetBackground_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSetBackground.Click

        MyFunctions.Slides.InsertSlides(MyFunctions.appDataPath & "Main.pptx")

    End Sub

End Class
