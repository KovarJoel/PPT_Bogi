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

    Private Sub GroupLieder_DialogLauncherClick(sender As Object, e As RibbonControlEventArgs) Handles GroupSongs.DialogLauncherClick

        Dim helpInfo As String

        helpInfo = "Mit diesem Add-In können Sie Lieder schneller in Ihre Präsentation einfügen." _
            & vbNewLine & vbNewLine & vbNewLine

        helpInfo &= " ******************" & vbNewLine
        helpInfo &= " * Lied einfügen: *" & vbNewLine
        helpInfo &= " ******************" & vbNewLine
        helpInfo &= "Mit diesem Button können Sie direkt ein Lied nach der aktuellen Position einfügen." & vbNewLine
        helpInfo &= "Als Eingabe geben Sie bitte die gewünschte Lied Nummer ein (z.B. 187)." & vbNewLine
        helpInfo &= "Falls Sie ein Lied, welches als 'a' und 'b' Teil existiert, geben Sie die Nummer mit " _
            & "anschließend dem Buchstaben ein (z.B. 85b)." & vbNewLine
        helpInfo &= "Falls Sie Lieder mit verschiedenen Sprachen, nur bestimmten Strophen, oder alternativem Text" _
                & "einzufügen wollen, müssen Sie dies leider manuell machen." _
                & vbNewLine & vbNewLine

        helpInfo &= " ***********************" & vbNewLine
        helpInfo &= " * Verzeichnis ändern: *" & vbNewLine
        helpInfo &= " ***********************" & vbNewLine
        helpInfo &= "Mit diesem Button legen Sie fest, aus welchem Verzeichnis die Lieder geöffnet werden sollen." & vbNewLine
        helpInfo &= "Geben Sie dazu den absoluten Ordnerpfad an (z.B. ""C:\Users\Max\Documents\Lieder"")." _
            & vbNewLine & vbNewLine & vbNewLine

        helpInfo &= "Sollten Sie Probleme oder weitere Fragen haben, wenden Sie sich bitte an Joel Kovar." & vbNewLine
        helpInfo &= "joel.m.kovar@gmail.com"


        MsgBox(helpInfo, MsgBoxStyle.OkOnly, "Hilfe 'PPT Bogi Add-In v1.0'")

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

    End Sub

    Private Sub ButtonInsertVideo_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonInsertVideo.Click

        Dim link As String = "http://raw.githubusercontent.com/KovarJoel/PPT_Bogi/master/PPT%20Bogi/MyFunctions.vb"
        Dim path As String = "C:\Users\joelm\Documents\PPT\temp.vb"

        MyFunctions.Files.LoadFileFromLink(link, path)

    End Sub

    Private Sub ButtonSetBackground_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSetBackground.Click

        MyFunctions.Slides.InsertSlides(MyFunctions.appDataPath & "Main.pptx")

    End Sub

End Class
