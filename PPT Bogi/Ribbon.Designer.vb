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
        Me.ButtonInsertSong = Me.Factory.CreateRibbonButton
        Me.ButtonOpenDirectory = Me.Factory.CreateRibbonButton
        Me.ButtonChangeDirectory = Me.Factory.CreateRibbonButton
        Me.CheckBoxAnnounce = Me.Factory.CreateRibbonCheckBox
        Me.GroupSlides = Me.Factory.CreateRibbonGroup
        Me.ButtonInsertMap = Me.Factory.CreateRibbonButton
        Me.ButtonLocations = Me.Factory.CreateRibbonButton
        Me.ButtonInsertEGW = Me.Factory.CreateRibbonButton
        Me.ButtonInsertClosing = Me.Factory.CreateRibbonButton
        Me.ButtonInsertVideo = Me.Factory.CreateRibbonButton
        Me.ButtonSetBackground = Me.Factory.CreateRibbonButton
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
        Me.GroupSongs.Items.Add(Me.ButtonOpenDirectory)
        Me.GroupSongs.Items.Add(Me.ButtonChangeDirectory)
        Me.GroupSongs.Items.Add(Me.CheckBoxAnnounce)
        Me.GroupSongs.Label = "Lieder"
        Me.GroupSongs.Name = "GroupSongs"
        '
        'ButtonInsertSong
        '
        Me.ButtonInsertSong.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonInsertSong.Image = Global.PPT_Bogi.My.Resources.Resources.Song
        Me.ButtonInsertSong.Label = "Lied einfügen"
        Me.ButtonInsertSong.Name = "ButtonInsertSong"
        Me.ButtonInsertSong.ScreenTip = "Lied einfügen"
        Me.ButtonInsertSong.ShowImage = True
        Me.ButtonInsertSong.SuperTip = "Fügt ein Lied an der aktuellen Position ein."
        '
        'ButtonOpenDirectory
        '
        Me.ButtonOpenDirectory.Image = Global.PPT_Bogi.My.Resources.Resources.Folder
        Me.ButtonOpenDirectory.Label = "Verzeichnis öffnen"
        Me.ButtonOpenDirectory.Name = "ButtonOpenDirectory"
        Me.ButtonOpenDirectory.ShowImage = True
        '
        'ButtonChangeDirectory
        '
        Me.ButtonChangeDirectory.Image = Global.PPT_Bogi.My.Resources.Resources.Folder
        Me.ButtonChangeDirectory.Label = "Verzeichnis ändern"
        Me.ButtonChangeDirectory.Name = "ButtonChangeDirectory"
        Me.ButtonChangeDirectory.ScreenTip = "Verzeichnis ändern"
        Me.ButtonChangeDirectory.ShowImage = True
        Me.ButtonChangeDirectory.SuperTip = "Ändert das Verzeichnis, aus welchem die Lieder geöffnet werden sollen."
        '
        'CheckBoxAnnounce
        '
        Me.CheckBoxAnnounce.Label = "Ankündigung"
        Me.CheckBoxAnnounce.Name = "CheckBoxAnnounce"
        '
        'GroupSlides
        '
        Me.GroupSlides.Items.Add(Me.ButtonSetBackground)
        Me.GroupSlides.Items.Add(Me.ButtonInsertMap)
        Me.GroupSlides.Items.Add(Me.ButtonLocations)
        Me.GroupSlides.Items.Add(Me.ButtonInsertEGW)
        Me.GroupSlides.Items.Add(Me.ButtonInsertClosing)
        Me.GroupSlides.Items.Add(Me.ButtonInsertVideo)
        Me.GroupSlides.Label = "Folien"
        Me.GroupSlides.Name = "GroupSlides"
        '
        'ButtonInsertMap
        '
        Me.ButtonInsertMap.Label = "Missionskarte"
        Me.ButtonInsertMap.Name = "ButtonInsertMap"
        Me.ButtonInsertMap.ScreenTip = "Missionskarte"
        Me.ButtonInsertMap.SuperTip = "Fügt die Karte vom Missionsbericht ein."
        '
        'ButtonLocations
        '
        Me.ButtonLocations.Label = "Sabbatschule"
        Me.ButtonLocations.Name = "ButtonLocations"
        '
        'ButtonInsertEGW
        '
        Me.ButtonInsertEGW.Label = "EGW Folie"
        Me.ButtonInsertEGW.Name = "ButtonInsertEGW"
        '
        'ButtonInsertClosing
        '
        Me.ButtonInsertClosing.Label = "Schlusslied"
        Me.ButtonInsertClosing.Name = "ButtonInsertClosing"
        '
        'ButtonInsertVideo
        '
        Me.ButtonInsertVideo.Label = "Missionsvideo"
        Me.ButtonInsertVideo.Name = "ButtonInsertVideo"
        '
        'ButtonSetBackground
        '
        Me.ButtonSetBackground.Label = "Hintergrund"
        Me.ButtonSetBackground.Name = "ButtonSetBackground"
        Me.ButtonSetBackground.ScreenTip = "Hintergrund"
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
    Friend WithEvents ButtonLocations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonInsertEGW As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonOpenDirectory As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CheckBoxAnnounce As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ButtonInsertVideo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSetBackground As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon() As Ribbon
        Get
            Return Me.GetRibbon(Of Ribbon)()
        End Get
    End Property
End Class
