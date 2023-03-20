Partial Class Ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Me.TabBogi = Me.Factory.CreateRibbonTab
        Me.GroupSongs = Me.Factory.CreateRibbonGroup
        Me.CheckBoxAnnounce = Me.Factory.CreateRibbonCheckBox
        Me.GroupSlides = Me.Factory.CreateRibbonGroup
        Me.ButtonInsertSong = Me.Factory.CreateRibbonButton
        Me.ButtonChangeDirectory = Me.Factory.CreateRibbonButton
        Me.ButtonInsertSongManually = Me.Factory.CreateRibbonButton
        Me.ButtonSetBackground = Me.Factory.CreateRibbonButton
        Me.ButtonInsertMap = Me.Factory.CreateRibbonButton
        Me.ButtonInsertLocations = Me.Factory.CreateRibbonButton
        Me.ButtonInsertEGW = Me.Factory.CreateRibbonButton
        Me.ButtonInsertClosing = Me.Factory.CreateRibbonButton
        Me.ButtonInsertVideo = Me.Factory.CreateRibbonButton
        Me.TabBogi.SuspendLayout()
        Me.GroupSongs.SuspendLayout()
        Me.GroupSlides.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabBogi
        '
        Me.TabBogi.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabBogi.Groups.Add(Me.GroupSongs)
        Me.TabBogi.Groups.Add(Me.GroupSlides)
        Me.TabBogi.Label = "PPT Bogi"
        Me.TabBogi.Name = "TabBogi"
        '
        'GroupSongs
        '
        RibbonDialogLauncherImpl1.ScreenTip = "Hilfe"
        RibbonDialogLauncherImpl1.SuperTip = "Bei Problemen oder Fragen."
        Me.GroupSongs.DialogLauncher = RibbonDialogLauncherImpl1
        Me.GroupSongs.Items.Add(Me.ButtonInsertSong)
        Me.GroupSongs.Items.Add(Me.ButtonChangeDirectory)
        Me.GroupSongs.Items.Add(Me.ButtonInsertSongManually)
        Me.GroupSongs.Items.Add(Me.CheckBoxAnnounce)
        Me.GroupSongs.Label = "Lieder"
        Me.GroupSongs.Name = "GroupSongs"
        '
        'CheckBoxAnnounce
        '
        Me.CheckBoxAnnounce.Label = "Ankündigung"
        Me.CheckBoxAnnounce.Name = "CheckBoxAnnounce"
        Me.CheckBoxAnnounce.ScreenTip = "Ankündigung"
        Me.CheckBoxAnnounce.SuperTip = "Wenn aktiviert, wird vor jedem eingefügten Lied automatisch eine Titelfolie mit L" &
    "ied-Nummer erstellt."
        '
        'GroupSlides
        '
        Me.GroupSlides.Items.Add(Me.ButtonSetBackground)
        Me.GroupSlides.Items.Add(Me.ButtonInsertMap)
        Me.GroupSlides.Items.Add(Me.ButtonInsertLocations)
        Me.GroupSlides.Items.Add(Me.ButtonInsertEGW)
        Me.GroupSlides.Items.Add(Me.ButtonInsertClosing)
        Me.GroupSlides.Items.Add(Me.ButtonInsertVideo)
        Me.GroupSlides.Label = "Folien"
        Me.GroupSlides.Name = "GroupSlides"
        '
        'ButtonInsertSong
        '
        Me.ButtonInsertSong.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonInsertSong.Image = Global.PPT_Bogi.My.Resources.Resources.Song
        Me.ButtonInsertSong.Label = "Lied einfügen"
        Me.ButtonInsertSong.Name = "ButtonInsertSong"
        Me.ButtonInsertSong.ScreenTip = "Lied einfügen"
        Me.ButtonInsertSong.ShowImage = True
        Me.ButtonInsertSong.SuperTip = "Fügt ein Lied ein."
        '
        'ButtonChangeDirectory
        '
        Me.ButtonChangeDirectory.Label = "Verzeichnis ändern"
        Me.ButtonChangeDirectory.Name = "ButtonChangeDirectory"
        Me.ButtonChangeDirectory.ScreenTip = "Verzeichnis ändern"
        Me.ButtonChangeDirectory.SuperTip = "Ändert das Verzeichnis, aus welchem die Lieder geöffnet werden sollen. Dabei muss" &
    " eine Datei aus dem gewählten Verzeichniss angeklickt werden, welche allerdings " &
    "nicht geöffnet wird."
        '
        'ButtonInsertSongManually
        '
        Me.ButtonInsertSongManually.Label = "Manuell einfügen"
        Me.ButtonInsertSongManually.Name = "ButtonInsertSongManually"
        Me.ButtonInsertSongManually.ScreenTip = "Lied manuell einfügen"
        Me.ButtonInsertSongManually.SuperTip = "Lässt einen selber das aktuelle Verzeichnis von Liedern brwosen."
        '
        'ButtonSetBackground
        '
        Me.ButtonSetBackground.Label = "Titelfolie"
        Me.ButtonSetBackground.Name = "ButtonSetBackground"
        Me.ButtonSetBackground.ScreenTip = "Titelfolie einfügen"
        Me.ButtonSetBackground.SuperTip = "Fügt eine leere Titelfolie ein."
        '
        'ButtonInsertMap
        '
        Me.ButtonInsertMap.Label = "Missionskarte"
        Me.ButtonInsertMap.Name = "ButtonInsertMap"
        Me.ButtonInsertMap.ScreenTip = "Missionskarte"
        Me.ButtonInsertMap.SuperTip = "Fügt die Karte vom Missionsbericht ein."
        '
        'ButtonInsertLocations
        '
        Me.ButtonInsertLocations.Label = "Sabbatschule"
        Me.ButtonInsertLocations.Name = "ButtonInsertLocations"
        Me.ButtonInsertLocations.ScreenTip = "Sabbatschule"
        Me.ButtonInsertLocations.SuperTip = "Fügt eine Folie mit den vorhanden Sabbatschulgruppen und deren Treffpunkten ein."
        '
        'ButtonInsertEGW
        '
        Me.ButtonInsertEGW.Label = "EGW Folie"
        Me.ButtonInsertEGW.Name = "ButtonInsertEGW"
        Me.ButtonInsertEGW.ScreenTip = "EGW Folie einfügen."
        Me.ButtonInsertEGW.SuperTip = "Fügt eine EGW Folie ein."
        '
        'ButtonInsertClosing
        '
        Me.ButtonInsertClosing.Label = "Postludium"
        Me.ButtonInsertClosing.Name = "ButtonInsertClosing"
        Me.ButtonInsertClosing.ScreenTip = "Postludium"
        Me.ButtonInsertClosing.SuperTip = "Fügt die letzte Folie ein, welche während des Postludiums gezeigt wird."
        '
        'ButtonInsertVideo
        '
        Me.ButtonInsertVideo.Label = "Missionsvideo"
        Me.ButtonInsertVideo.Name = "ButtonInsertVideo"
        Me.ButtonInsertVideo.ScreenTip = "Missionsvideo"
        Me.ButtonInsertVideo.SuperTip = "Downloaded nach eingabe vom Link das Missionsvideo und fügt dieses ein."
        '
        'Ribbon
        '
        Me.Name = "Ribbon"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.TabBogi)
        Me.TabBogi.ResumeLayout(False)
        Me.TabBogi.PerformLayout()
        Me.GroupSongs.ResumeLayout(False)
        Me.GroupSongs.PerformLayout()
        Me.GroupSlides.ResumeLayout(False)
        Me.GroupSlides.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabBogi As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GroupSongs As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonInsertSong As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonChangeDirectory As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GroupSlides As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonInsertClosing As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonInsertMap As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonInsertLocations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonInsertEGW As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CheckBoxAnnounce As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ButtonInsertVideo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSetBackground As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonInsertSongManually As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon() As Ribbon
        Get
            Return Me.GetRibbon(Of Ribbon)()
        End Get
    End Property
End Class
