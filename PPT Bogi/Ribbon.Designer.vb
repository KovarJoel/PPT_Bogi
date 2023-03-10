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
        Me.GroupLieder = Me.Factory.CreateRibbonGroup
        Me.ButtonInsertSong = Me.Factory.CreateRibbonButton
        Me.ButtonChangeDirectory = Me.Factory.CreateRibbonButton
        Me.TabBogi.SuspendLayout()
        Me.GroupLieder.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabBogi
        '
        Me.TabBogi.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabBogi.Groups.Add(Me.GroupLieder)
        Me.TabBogi.Label = "PPT Bogi"
        Me.TabBogi.Name = "TabBogi"
        '
        'GroupLieder
        '
        RibbonDialogLauncherImpl1.ScreenTip = "Hilfe"
        RibbonDialogLauncherImpl1.SuperTip = "Bei Problemen oder Fragen."
        Me.GroupLieder.DialogLauncher = RibbonDialogLauncherImpl1
        Me.GroupLieder.Items.Add(Me.ButtonInsertSong)
        Me.GroupLieder.Items.Add(Me.ButtonChangeDirectory)
        Me.GroupLieder.Label = "Lieder"
        Me.GroupLieder.Name = "GroupLieder"
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
        'ButtonChangeDirectory
        '
        Me.ButtonChangeDirectory.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonChangeDirectory.Image = Global.PPT_Bogi.My.Resources.Resources.Folder
        Me.ButtonChangeDirectory.Label = "Verzeichnis ändern"
        Me.ButtonChangeDirectory.Name = "ButtonChangeDirectory"
        Me.ButtonChangeDirectory.ScreenTip = "Verzeichnis ändern"
        Me.ButtonChangeDirectory.ShowImage = True
        Me.ButtonChangeDirectory.SuperTip = "Ändert das Verzeichnis, aus welchem die Lieder geöffnet werden sollen."
        '
        'Ribbon
        '
        Me.Name = "Ribbon"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.TabBogi)
        Me.TabBogi.ResumeLayout(False)
        Me.TabBogi.PerformLayout()
        Me.GroupLieder.ResumeLayout(False)
        Me.GroupLieder.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabBogi As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GroupLieder As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonInsertSong As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonChangeDirectory As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon() As Ribbon
        Get
            Return Me.GetRibbon(Of Ribbon)()
        End Get
    End Property
End Class
