Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports ShareRibbon
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class ThisAddIn

    Private Const ThemePptTaskPaneInitialWidth As Integer = 780
    Private Const ThemePptTaskPaneMinUsableWidth As Integer = 680

    Private themePptTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private themePptControl As ThemePptTaskPane
    Private _translateService As PowerPointTranslateService

    Private ReadOnly _lazyWebView2 As New Lazy(Of Boolean)(Function()
                                                               WebView2Loader.EnsureWebView2Loader()
                                                               Return True
                                                           End Function)

    Private ReadOnly _lazySqlite As New Lazy(Of Boolean)(Function()
                                                             SqliteNativeLoader.EnsureLoaded()
                                                             Return True
                                                         End Function)

    Private Sub PowerPointAi_Startup() Handles Me.Startup
        PhaseStartupManager.Instance.RunCriticalPhase(Me.Application)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        themePptTaskPane = Nothing
        themePptControl = Nothing
    End Sub

    Private Sub EnsureCoreServicesLoaded()
        If PhaseStartupManager.Instance.IsBackgroundReady Then Return

        Try
            Dim webView2Init = _lazyWebView2.Value
        Catch ex As Exception
            MessageBox.Show("WebView2 初始化失败: " & ex.Message)
        End Try

        Try
            Dim sqliteInit = _lazySqlite.Value
        Catch ex As Exception
            MessageBox.Show("SQLite 原生库加载失败: " & ex.Message)
        End Try
    End Sub

    Public Sub ShowThemePptTaskPane()
        EnsureCoreServicesLoaded()

        If themePptTaskPane Is Nothing Then
            themePptControl = New ThemePptTaskPane(Me.Application)
            themePptTaskPane = Me.CustomTaskPanes.Add(themePptControl, "AI生成PPT")
            themePptTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            themePptTaskPane.Width = ThemePptTaskPaneInitialWidth
        ElseIf themePptTaskPane.Width < ThemePptTaskPaneMinUsableWidth Then
            themePptTaskPane.Width = ThemePptTaskPaneInitialWidth
        End If

        themePptTaskPane.Visible = True
    End Sub

    Public Sub HideThemePptTaskPane()
        If themePptTaskPane IsNot Nothing Then
            themePptTaskPane.Visible = False
        End If
    End Sub

End Class
