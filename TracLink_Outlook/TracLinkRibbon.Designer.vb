Partial Class TracLinkRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms クラス作成デザイナーのサポートに必要です
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'この呼び出しは、コンポーネント デザイナーで必要です。
        InitializeComponent()

    End Sub

    'Component は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
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

    'コンポーネント デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャはコンポーネント デザイナーで必要です。
    'コンポーネント デザイナーを使って変更できます。
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Me.TracLinkTab = Me.Factory.CreateRibbonTab
        Me.TracLinkUpdate = Me.Factory.CreateRibbonGroup
        Me.Sync = Me.Factory.CreateRibbonButton
        Me.TracLinkSetting = Me.Factory.CreateRibbonGroup
        Me.QuickProjectURLSetting = Me.Factory.CreateRibbonEditBox
        Me.TracLinkTab.SuspendLayout()
        Me.TracLinkUpdate.SuspendLayout()
        Me.TracLinkSetting.SuspendLayout()
        Me.SuspendLayout()
        '
        'TracLinkTab
        '
        Me.TracLinkTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TracLinkTab.Groups.Add(Me.TracLinkUpdate)
        Me.TracLinkTab.Groups.Add(Me.TracLinkSetting)
        Me.TracLinkTab.Label = "TracLink"
        Me.TracLinkTab.Name = "TracLinkTab"
        '
        'TracLinkUpdate
        '
        Me.TracLinkUpdate.Items.Add(Me.Sync)
        Me.TracLinkUpdate.Label = "更新"
        Me.TracLinkUpdate.Name = "TracLinkUpdate"
        '
        'Sync
        '
        Me.Sync.Label = "同期"
        Me.Sync.Name = "Sync"
        '
        'TracLinkSetting
        '
        Me.TracLinkSetting.DialogLauncher = RibbonDialogLauncherImpl1
        Me.TracLinkSetting.Items.Add(Me.QuickProjectURLSetting)
        Me.TracLinkSetting.Label = "設定"
        Me.TracLinkSetting.Name = "TracLinkSetting"
        '
        'QuickProjectURLSetting
        '
        Me.QuickProjectURLSetting.Label = "プロジェクトURL"
        Me.QuickProjectURLSetting.Name = "QuickProjectURLSetting"
        Me.QuickProjectURLSetting.SizeString = "111111111122222222223333333333444444444455555555556666666666"
        Me.QuickProjectURLSetting.Text = Nothing
        '
        'TracLinkRibbon
        '
        Me.Name = "TracLinkRibbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.TracLinkTab)
        Me.TracLinkTab.ResumeLayout(False)
        Me.TracLinkTab.PerformLayout()
        Me.TracLinkUpdate.ResumeLayout(False)
        Me.TracLinkUpdate.PerformLayout()
        Me.TracLinkSetting.ResumeLayout(False)
        Me.TracLinkSetting.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TracLinkTab As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents TracLinkUpdate As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Sync As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TracLinkSetting As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents QuickProjectURLSetting As Microsoft.Office.Tools.Ribbon.RibbonEditBox
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As TracLinkRibbon
        Get
            Return Me.GetRibbon(Of TracLinkRibbon)()
        End Get
    End Property
End Class
