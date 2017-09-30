Imports Microsoft.Office.Tools.Ribbon

Public Class TracLinkRibbon

    Private YYUTIL As New yyutil

    Private Sub ButtonProjectSettings_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    ''' <summary>
    ''' ダイアログボックス起動ツール押下時のフォーム表示処理(未実装)
    ''' </summary>
    ''' <param name="sender">未使用</param>
    ''' <param name="e">未使用</param>
    Private Sub TracLinkSetting_DialogLauncherClick(sender As Object, e As RibbonControlEventArgs) Handles TracLinkSetting.DialogLauncherClick
        'Dim projectForm As New ProjectForm
        'ProjectForm.Show()
    End Sub

    ''' <summary>
    ''' TracLink用リボン読み込み時の初期設定
    ''' </summary>
    ''' <param name="sender">未使用</param>
    ''' <param name="e">未使用</param>
    Private Sub TracLinkRibbon_Load(sender As Object, e As RibbonUIEventArgs) Handles Me.Load
        QuickProjectURLSetting.Text = YYUTIL.GetUserSetting("ProjectURL")
    End Sub

    ''' <summary>
    ''' ProjectURL簡易設定結果反映
    ''' </summary>
    ''' <param name="sender">未使用</param>
    ''' <param name="e">未使用</param>
    Private Sub QuickProjectURLSetting_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles QuickProjectURLSetting.TextChanged
        YYUTIL.SetUserSetting("ProjectURL", QuickProjectURLSetting.Text)
    End Sub

    Private Sub Sync_Click(sender As Object, e As RibbonControlEventArgs) Handles Sync.Click
        Dim traclink As TracLink = New TracLink

        Call traclink.updateAllProject("Trac")
    End Sub
End Class
