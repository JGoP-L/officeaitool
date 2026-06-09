' PowerPointAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon  ' 添加此引用
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Private Const DocmeePptxIdTagName As String = "wenduoduoAI_DocmeePptxId"

    Private Class PptTextTarget
        Public Property TextRange As PowerPoint.TextRange
        Public Property Shape As PowerPoint.Shape
        Public Property OriginalText As String
        Public Property SlideContextText As String
    End Class

    Private Class AiCreationOptions
        Public Property ModeName As String
        Public Property TargetLanguageName As String
        Public Property TargetLanguageCode As String
    End Class

    Private Class ReplaceSlideOptions
        Public Property Content As String
        Public Property PptxId As String
    End Class

    Private Shared _lastReplaceSlidePptxId As String = ""

    Protected Overrides Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
    End Sub

    Protected Overrides Sub WebResearchButton_Click(sender As Object, e As RibbonControlEventArgs)
    End Sub

    Protected Overrides Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs)
    End Sub
    Protected Overrides Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' Word 特定的数据分析逻辑
        MessageBox.Show("Word数据分析功能正在开发中...")
    End Sub

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint)
    End Function

    Protected Overrides Sub BatchDataGenButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show("批量数据生成功能仅适用于 Excel。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Protected Overrides Sub MCPButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' 创建并显示MCP配置表单
        Dim mcpConfigForm As New MCPConfigForm()
        If mcpConfigForm.ShowDialog() = DialogResult.OK Then
        End If
    End Sub

    Protected Overrides Async Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            ShareRibbon.GlobalStatusStripAll.ShowProgress("AI内容提效：正在启动AI生成单页...")
            LogInfo("[ReplaceSlide] Button clicked.")
            Dim request = ShowReplaceSlideDialog()
            LogInfo("[ReplaceSlide] Dialog closed. hasRequest=" & (request IsNot Nothing).ToString())
            If request Is Nothing OrElse
               String.IsNullOrWhiteSpace(request.Content) Then
                LogInfo("[ReplaceSlide] Canceled or empty content.")
                Return
            End If

            Await ReplaceCurrentSlideWithDocmeeAsync(request)
        Catch ex As Exception
            LogError("[ReplaceSlide] Failed in button handler.", ex)
            MessageBox.Show("AI生成单页出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 美化当前页 - 使用 Docmee newPageWithAiV2 生成结果页并替换当前页。
    Protected Overrides Async Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Await BeautifyCurrentSlideWithDocmeeTemplateAsync()
        Catch ex As Exception
            LogError("[BeautifyTemplate] Failed.", ex)
            MessageBox.Show("美化模板出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 判断形状类型（标题/副标题/正文）
    ''' </summary>
    Private Function GetShapeType(shp As Microsoft.Office.Interop.PowerPoint.Shape) As String
        Try
            If shp.PlaceholderFormat IsNot Nothing Then
                Select Case shp.PlaceholderFormat.Type
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                         Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                        Return "标题"
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle
                        Return "副标题"
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody
                        Return "正文"
                End Select
            End If
        Catch
        End Try
        Return "文本框"
    End Function

    Private Function ShowReplaceSlideDialog() As ReplaceSlideOptions
        Using dialog As New NewPageRequestForm()
            If dialog.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return New ReplaceSlideOptions() With {
                .Content = dialog.RequestContent
            }
        End Using
    End Function

    Private Async Function ReplaceCurrentSlideWithDocmeeAsync(request As ReplaceSlideOptions) As Task
        LogInfo("[ReplaceSlide] Start Docmee replace. contentLength=" & If(request?.Content, "").Length.ToString())
        Dim originalSlide = GetCurrentSlide()
        If originalSlide Is Nothing Then
            LogInfo("[ReplaceSlide] No current slide.")
            MessageBox.Show("请先选中要生成单页的幻灯片。", "AI生成单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        LogInfo("[ReplaceSlide] Current slide index=" & originalSlide.SlideIndex.ToString())

        request.PptxId = ResolveDocmeePptxId()
        LogInfo("[ReplaceSlide] Resolved Docmee PPT ID=" & If(request.PptxId, ""))
        If String.IsNullOrWhiteSpace(request.PptxId) Then
            request.PptxId = DocmeePptClient.GetRandomNewPageTemplateId()
            LogInfo("[ReplaceSlide] No presentation Docmee PPT ID, using random ID=" & request.PptxId)
        End If

        Dim client As New DocmeePptClient()
        ShareRibbon.GlobalStatusStripAll.ShowProgress("正在用 Docmee 生成新单页...")
        LogInfo("[ReplaceSlide] Calling Docmee newPageWithAiV2.")
        Dim progressWindow = CreateReplaceProgressWindow()
        progressWindow.Item1.Show()
        UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在匹配页面类型 ...")

        Dim result As DocmeeNewPageResult = Nothing
        Try
            result = Await client.NewPageWithAiV2Async(
                request.Content,
                request.PptxId,
                Sub(message)
                    If Not String.IsNullOrWhiteSpace(message) Then
                        ShareRibbon.GlobalStatusStripAll.ShowProgress("Docmee 单页生成中...")
                        UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在生成候选页面...")
                    End If
                End Sub)

            Dim pptxId = If(String.IsNullOrWhiteSpace(result.PptxId), request.PptxId, result.PptxId)
            Dim fileUrl = result.FileUrl
            LogInfo("[ReplaceSlide] Docmee result. pptxId=" & pptxId & ", hasFileUrl=" & Not String.IsNullOrWhiteSpace(fileUrl))
            If String.IsNullOrWhiteSpace(fileUrl) Then
                ShareRibbon.GlobalStatusStripAll.ShowProgress("正在获取AI生成单页 PPTX...")
                UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在整理页面结构...")
                LogInfo("[ReplaceSlide] Download URL missing, calling downloadPptx.")
                fileUrl = Await client.DownloadPptxAsync(pptxId, True, 6, 1200)
            End If

            Dim localPath = Path.Combine(Path.GetTempPath(), $"wenduoduoAI_newpage_{pptxId}_{Guid.NewGuid():N}.pptx")
            ShareRibbon.GlobalStatusStripAll.ShowProgress("正在下载AI生成单页 PPTX...")
            UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在下载候选页面...")
            LogInfo("[ReplaceSlide] Downloading PPTX to " & localPath)
            Await client.DownloadPptxFileAsync(fileUrl, localPath)

            ShareRibbon.GlobalStatusStripAll.ShowProgress("请选择要应用的AI生成单页...")
            UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在准备选择界面...")
            LogInfo("[ReplaceSlide] Selecting replacement slide from downloaded PPTX.")
            CloseReplaceProgressWindow(progressWindow.Item1)
            If Not ReplaceCurrentSlideWithSelectedSlideFromPptx(originalSlide, localPath) Then Return

            _lastReplaceSlidePptxId = pptxId
            SaveDocmeePptxId(pptxId)
            ShareRibbon.GlobalStatusStripAll.ShowProgress("AI生成单页完成")
            LogInfo("[ReplaceSlide] Completed.")
        Finally
            CloseReplaceProgressWindow(progressWindow.Item1)
        End Try
    End Function

    Private Function CreateReplaceProgressWindow(Optional sectionTitle As String = "新页面",
                                                 Optional initialStatus As String = "正在匹配页面类型 ...") As Tuple(Of Form, Label)
        Dim dialog As New NewPageProgressForm(sectionTitle, initialStatus)
        Return Tuple.Create(CType(dialog, Form), dialog.StatusLabel)
    End Function

    Private Sub UpdateReplaceProgressWindow(dialog As Form, label As Label, message As String)
        If dialog Is Nothing OrElse dialog.IsDisposed OrElse label Is Nothing OrElse label.IsDisposed Then Return

        If dialog.InvokeRequired Then
            dialog.BeginInvoke(CType(Sub() UpdateReplaceProgressWindow(dialog, label, message), MethodInvoker))
            Return
        End If

        label.Text = message
        label.Refresh()
        dialog.Refresh()
    End Sub

    Private Sub CloseReplaceProgressWindow(dialog As Form)
        If dialog Is Nothing OrElse dialog.IsDisposed Then Return

        If dialog.InvokeRequired Then
            dialog.BeginInvoke(CType(Sub() CloseReplaceProgressWindow(dialog), MethodInvoker))
            Return
        End If

        dialog.Close()
        dialog.Dispose()
    End Sub

    Private NotInheritable Class NewPageUiPainter
        Private Sub New()
        End Sub

        Private Const WM_NCLBUTTONDOWN As Integer = &HA1
        Private Shared ReadOnly HTCAPTION As New IntPtr(2)

        <DllImport("user32.dll")>
        Private Shared Function ReleaseCapture() As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        Public Shared Function CreateRoundedRect(rect As Rectangle, radius As Integer) As GraphicsPath
            Dim path As New GraphicsPath()
            If rect.Width <= 0 OrElse rect.Height <= 0 Then
                path.AddRectangle(rect)
                Return path
            End If

            Dim diameter = Math.Max(1, Math.Min(radius * 2, Math.Min(rect.Width, rect.Height)))
            Dim arc As New Rectangle(rect.Location, New Size(diameter, diameter))
            path.AddArc(arc, 180, 90)
            arc.X = rect.Right - diameter
            path.AddArc(arc, 270, 90)
            arc.Y = rect.Bottom - diameter
            path.AddArc(arc, 0, 90)
            arc.X = rect.Left
            path.AddArc(arc, 90, 90)
            path.CloseFigure()
            Return path
        End Function

        Public Shared Sub StyleCloseButton(btn As Button)
            btn.Text = "×"
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.BackColor = Color.FromArgb(246, 247, 251)
            btn.ForeColor = Color.FromArgb(82, 82, 91)
            btn.Font = New Font("Microsoft YaHei UI", 15.0!, FontStyle.Regular)
            btn.Cursor = Cursors.Hand
            btn.TabStop = False
            btn.UseVisualStyleBackColor = False
            AddHandler btn.MouseEnter, Sub()
                                           btn.BackColor = Color.FromArgb(238, 242, 255)
                                           btn.ForeColor = OfficeAIStyleHelper.BrandPrimary
                                       End Sub
            AddHandler btn.MouseLeave, Sub()
                                           btn.BackColor = Color.FromArgb(246, 247, 251)
                                           btn.ForeColor = Color.FromArgb(82, 82, 91)
                                       End Sub
        End Sub

        Public Shared Sub EnableFormDrag(frm As Form, ParamArray dragSurfaces As Control())
            If frm Is Nothing OrElse dragSurfaces Is Nothing Then Return

            For Each surface In dragSurfaces
                If surface Is Nothing Then Continue For
                AddHandler surface.MouseDown,
                    Sub(sender As Object, e As MouseEventArgs)
                        If e.Button <> MouseButtons.Left OrElse frm.IsDisposed Then Return

                        ReleaseCapture()
                        SendMessage(frm.Handle, WM_NCLBUTTONDOWN, HTCAPTION, IntPtr.Zero)
                    End Sub
            Next
        End Sub
    End Class

    Private Class NewPageCardPanel
        Inherits Panel

        Public Sub New()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.ResizeRedraw Or
                     ControlStyles.UserPaint, True)
            BackColor = Color.FromArgb(245, 246, 250)
        End Sub

        Protected Overrides Sub OnResize(e As EventArgs)
            MyBase.OnResize(e)
            If Width <= 0 OrElse Height <= 0 Then Return
            Using path = NewPageUiPainter.CreateRoundedRect(New Rectangle(0, 0, Width, Height), 10)
                Region = New Region(path)
            End Using
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim rect = New Rectangle(2, 2, Width - 5, Height - 5)
            Using path = NewPageUiPainter.CreateRoundedRect(rect, 10)
                Using brush As New SolidBrush(Color.White)
                    g.FillPath(brush, path)
                End Using
                Using pen As New Pen(Color.FromArgb(100, 55, 229), 3.0F)
                    g.DrawPath(pen, path)
                End Using
            End Using
        End Sub
    End Class

    Private Class NewPageInputPanel
        Inherits Panel

        Public Sub New()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.ResizeRedraw Or
                     ControlStyles.UserPaint, True)
            BackColor = Color.White
        End Sub

        Protected Overrides Sub OnResize(e As EventArgs)
            MyBase.OnResize(e)
            If Width <= 0 OrElse Height <= 0 Then Return
            Using path = NewPageUiPainter.CreateRoundedRect(New Rectangle(0, 0, Width, Height), 12)
                Region = New Region(path)
            End Using
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim rect = New Rectangle(1, 1, Width - 3, Height - 3)
            Using path = NewPageUiPainter.CreateRoundedRect(rect, 12)
                Using brush As New SolidBrush(Color.White)
                    g.FillPath(brush, path)
                End Using
                Using pen As New Pen(Color.FromArgb(236, 238, 245), 2.0F)
                    g.DrawPath(pen, path)
                End Using
            End Using
        End Sub
    End Class

    Private Class SlidePreviewFramePanel
        Inherits Panel

        Public Sub New()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.ResizeRedraw Or
                     ControlStyles.UserPaint, True)
            BackColor = Color.White
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim shadowRect = New Rectangle(8, 10, Width - 16, Height - 18)
            Using shadowPath = NewPageUiPainter.CreateRoundedRect(shadowRect, 2)
                Using shadowBrush As New SolidBrush(Color.FromArgb(22, 15, 23, 42))
                    g.FillPath(shadowBrush, shadowPath)
                End Using
            End Using

            Dim slideRect = New Rectangle(0, 0, Width - 18, Height - 18)
            Using slidePath = NewPageUiPainter.CreateRoundedRect(slideRect, 2)
                Using brush As New SolidBrush(Color.White)
                    g.FillPath(brush, slidePath)
                End Using
            End Using
        End Sub
    End Class

    Private Class NewPageGradientButton
        Inherits Button

        Private _hover As Boolean

        Public Sub New()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.ResizeRedraw Or
                     ControlStyles.UserPaint, True)
            FlatStyle = FlatStyle.Flat
            FlatAppearance.BorderSize = 0
            Cursor = Cursors.Hand
            Font = New Font("Microsoft YaHei UI", 11.0!, FontStyle.Bold)
            ForeColor = Color.White
            Height = 42
            TabStop = False
            UseVisualStyleBackColor = False
        End Sub

        Protected Overrides Sub OnMouseEnter(e As EventArgs)
            _hover = True
            Invalidate()
            MyBase.OnMouseEnter(e)
        End Sub

        Protected Overrides Sub OnMouseLeave(e As EventArgs)
            _hover = False
            Invalidate()
            MyBase.OnMouseLeave(e)
        End Sub

        Protected Overrides Sub OnEnabledChanged(e As EventArgs)
            Invalidate()
            MyBase.OnEnabledChanged(e)
        End Sub

        Protected Overrides Sub OnPaint(pevent As PaintEventArgs)
            Dim g = pevent.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim rect = New Rectangle(0, 0, Width - 1, Height - 1)
            Using path = NewPageUiPainter.CreateRoundedRect(rect, 8)
                If Enabled Then
                    Dim leftColor = If(_hover, Color.FromArgb(91, 61, 235), Color.FromArgb(109, 52, 229))
                    Dim rightColor = If(_hover, Color.FromArgb(78, 70, 229), Color.FromArgb(89, 45, 222))
                    Using brush As New LinearGradientBrush(rect, leftColor, rightColor, LinearGradientMode.Horizontal)
                        g.FillPath(brush, path)
                    End Using
                Else
                    Using brush As New SolidBrush(Color.FromArgb(203, 213, 225))
                        g.FillPath(brush, path)
                    End Using
                End If
            End Using

            TextRenderer.DrawText(g,
                                  Text,
                                  Font,
                                  rect,
                                  If(Enabled, Color.White, Color.FromArgb(241, 245, 249)),
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
        End Sub
    End Class

    Private Class NewPageSpinnerControl
        Inherits Control

        Private ReadOnly _timer As New System.Windows.Forms.Timer() With {.Interval = 30}
        Private _angle As Single

        Public Sub New()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.ResizeRedraw Or
                     ControlStyles.UserPaint, True)
            Size = New Size(72, 72)
            BackColor = Color.White
            AddHandler _timer.Tick, Sub()
                                        _angle = (_angle + 12.0F) Mod 360.0F
                                        Invalidate()
                                    End Sub
        End Sub

        Public Sub Start()
            _timer.Start()
        End Sub

        Public Sub [Stop]()
            _timer.Stop()
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _timer.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim rect = New Rectangle(10, 10, Width - 20, Height - 20)
            For i As Integer = 0 To 8
                Dim alpha As Integer = Math.Max(20, 230 - i * 22)
                Dim spinnerColor As Color = If(i < 3, OfficeAIStyleHelper.BrandPrimary, Color.FromArgb(69, 166, 244))
                Using pen As New Pen(Color.FromArgb(alpha, spinnerColor), 7.0F)
                    pen.StartCap = LineCap.Round
                    pen.EndCap = LineCap.Round
                    g.DrawArc(pen, rect, _angle - i * 18.0F, 18.0F)
                End Using
            Next
        End Sub
    End Class

    Private Class NewPageMarqueeControl
        Inherits Control

        Private ReadOnly _timer As New System.Windows.Forms.Timer() With {.Interval = 28}
        Private _offset As Integer

        Public Sub New()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.ResizeRedraw Or
                     ControlStyles.UserPaint, True)
            Size = New Size(160, 12)
            BackColor = Color.White
            AddHandler _timer.Tick, Sub()
                                        _offset = (_offset + 5) Mod Math.Max(1, Width + 70)
                                        Invalidate()
                                    End Sub
        End Sub

        Public Sub Start()
            _timer.Start()
        End Sub

        Public Sub [Stop]()
            _timer.Stop()
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _timer.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim trackRect = New Rectangle(0, 2, Width - 1, Height - 4)
            Using trackPath = NewPageUiPainter.CreateRoundedRect(trackRect, 6)
                Using brush As New SolidBrush(Color.FromArgb(240, 242, 247))
                    g.FillPath(brush, trackPath)
                End Using
            End Using

            Dim pillWidth = 62
            Dim pillX = _offset - pillWidth
            Dim pillRect = New Rectangle(pillX, 2, pillWidth, Height - 4)
            Using pillPath = NewPageUiPainter.CreateRoundedRect(pillRect, 6)
                Using brush As New LinearGradientBrush(pillRect, Color.FromArgb(189, 169, 242), Color.FromArgb(105, 82, 235), LinearGradientMode.Horizontal)
                    g.FillPath(brush, pillPath)
                End Using
            End Using
        End Sub
    End Class

    Private Class NewPageRequestForm
        Inherits Form

        Private Const MaxInputChars As Integer = 1000
        Private Const PlaceholderText As String = "请输入页面的标题或者具体内容，我会根据您输入的文本为您生成新的页面"

        Private ReadOnly _inputBox As New TextBox()
        Private ReadOnly _counterLabel As New Label()
        Private ReadOnly _submitButton As New NewPageGradientButton()
        Private ReadOnly _fadeTimer As New System.Windows.Forms.Timer() With {.Interval = 15}
        Private _placeholderActive As Boolean
        Private _requestContent As String = ""

        Public ReadOnly Property RequestContent As String
            Get
                Return _requestContent
            End Get
        End Property

        Public Sub New()
            BuildLayout()
            AddHandler _fadeTimer.Tick, AddressOf FadeTimer_Tick
        End Sub

        Private Sub BuildLayout()
            Text = "AI生成单页"
            Font = OfficeAIStyleHelper.FontUi
            BackColor = Color.FromArgb(245, 246, 250)
            ClientSize = New Size(960, 600)
            FormBorderStyle = FormBorderStyle.None
            StartPosition = FormStartPosition.CenterScreen
            ShowIcon = False
            ShowInTaskbar = False
            KeyPreview = True

            Dim sectionLabel As New Label() With {
                .Text = "新页面",
                .Location = New Point(34, 16),
                .AutoSize = True,
                .Font = New Font("Microsoft YaHei UI", 10.0!, FontStyle.Regular),
                .ForeColor = Color.FromArgb(107, 114, 128)
            }
            Controls.Add(sectionLabel)

            Dim card As New NewPageCardPanel() With {
                .Location = New Point(32, 42),
                .Size = New Size(896, 520)
            }
            Controls.Add(card)
            NewPageUiPainter.EnableFormDrag(Me, Me, sectionLabel, card)

            Dim closeButton As New Button() With {
                .Location = New Point(card.Width - 48, 14),
                .Size = New Size(34, 34),
                .DialogResult = DialogResult.Cancel
            }
            NewPageUiPainter.StyleCloseButton(closeButton)
            AddHandler closeButton.Click, Sub()
                                             DialogResult = DialogResult.Cancel
                                             Close()
                                         End Sub
            card.Controls.Add(closeButton)
            CancelButton = closeButton

            Dim titleLabel As New Label() With {
                .Text = "请输入页面标题或内容",
                .Location = New Point(30, 62),
                .AutoSize = True,
                .Font = New Font("Microsoft YaHei UI", 13.0!, FontStyle.Bold),
                .ForeColor = Color.FromArgb(55, 65, 81)
            }
            card.Controls.Add(titleLabel)
            NewPageUiPainter.EnableFormDrag(Me, titleLabel)

            Dim inputShell As New NewPageInputPanel() With {
                .Location = New Point(30, 108),
                .Size = New Size(card.Width - 60, 330)
            }
            card.Controls.Add(inputShell)

            _inputBox.BorderStyle = BorderStyle.None
            _inputBox.Multiline = True
            _inputBox.ScrollBars = ScrollBars.Vertical
            _inputBox.AcceptsReturn = True
            _inputBox.AcceptsTab = True
            _inputBox.MaxLength = MaxInputChars
            _inputBox.Font = New Font("Microsoft YaHei UI", 10.5!, FontStyle.Regular)
            _inputBox.Location = New Point(18, 18)
            _inputBox.Size = New Size(inputShell.Width - 36, inputShell.Height - 36)
            _inputBox.BackColor = Color.White
            AddHandler _inputBox.Enter, AddressOf InputBox_Enter
            AddHandler _inputBox.Leave, AddressOf InputBox_Leave
            AddHandler _inputBox.TextChanged, AddressOf InputBox_TextChanged
            inputShell.Controls.Add(_inputBox)

            _counterLabel.Location = New Point(inputShell.Right - 100, inputShell.Bottom + 6)
            _counterLabel.Size = New Size(100, 22)
            _counterLabel.TextAlign = ContentAlignment.MiddleRight
            _counterLabel.Font = New Font("Microsoft YaHei UI", 9.0!, FontStyle.Regular)
            _counterLabel.ForeColor = Color.FromArgb(156, 163, 175)
            card.Controls.Add(_counterLabel)

            _submitButton.Text = "⚡  AI智能生成新页面"
            _submitButton.Location = New Point(30, card.Height - 78)
            _submitButton.Size = New Size(card.Width - 60, 44)
            AddHandler _submitButton.Click, AddressOf SubmitButton_Click
            card.Controls.Add(_submitButton)

            SetPlaceholder()
            AddHandler KeyDown, Sub(sender, e)
                                    If e.KeyCode = Keys.Escape Then
                                        DialogResult = DialogResult.Cancel
                                        Close()
                                    End If
                                End Sub
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            Opacity = 0
            _fadeTimer.Start()
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _fadeTimer.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Sub FadeTimer_Tick(sender As Object, e As EventArgs)
            Opacity = Math.Min(1, Opacity + 0.08)
            If Opacity >= 1 Then _fadeTimer.Stop()
        End Sub

        Private Sub InputBox_Enter(sender As Object, e As EventArgs)
            If Not _placeholderActive Then Return
            _placeholderActive = False
            _inputBox.Text = ""
            _inputBox.ForeColor = OfficeAIStyleHelper.TextPrimary
            UpdateCounter()
        End Sub

        Private Sub InputBox_Leave(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(_inputBox.Text) Then
                SetPlaceholder()
            End If
        End Sub

        Private Sub InputBox_TextChanged(sender As Object, e As EventArgs)
            UpdateCounter()
        End Sub

        Private Sub SetPlaceholder()
            _placeholderActive = True
            _inputBox.ForeColor = Color.FromArgb(180, 184, 194)
            _inputBox.Text = PlaceholderText
            _inputBox.SelectionStart = 0
            _inputBox.SelectionLength = 0
            UpdateCounter()
        End Sub

        Private Function GetActualText() As String
            If _placeholderActive Then Return ""
            Return If(_inputBox.Text, "").Trim()
        End Function

        Private Sub UpdateCounter()
            Dim count = GetActualText().Length
            _counterLabel.Text = count.ToString() & " / " & MaxInputChars.ToString()
        End Sub

        Private Sub SubmitButton_Click(sender As Object, e As EventArgs)
            Dim content = GetActualText()
            If String.IsNullOrWhiteSpace(content) Then
                MessageBox.Show("请输入页面标题或内容。", "AI生成单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
                _inputBox.Focus()
                Return
            End If

            _requestContent = content
            DialogResult = DialogResult.OK
            Close()
        End Sub
    End Class

    Private Class NewPageProgressForm
        Inherits Form

        Private ReadOnly _spinner As New NewPageSpinnerControl()
        Private ReadOnly _marquee As New NewPageMarqueeControl()
        Private ReadOnly _statusLabel As New Label()
        Private ReadOnly _fadeTimer As New System.Windows.Forms.Timer() With {.Interval = 15}

        Public ReadOnly Property StatusLabel As Label
            Get
                Return _statusLabel
            End Get
        End Property

        Public Sub New(sectionTitle As String, initialStatus As String)
            BuildLayout(sectionTitle, initialStatus)
            AddHandler _fadeTimer.Tick, AddressOf FadeTimer_Tick
        End Sub

        Private Sub BuildLayout(sectionTitle As String, initialStatus As String)
            Text = sectionTitle
            Font = OfficeAIStyleHelper.FontUi
            BackColor = Color.FromArgb(245, 246, 250)
            ClientSize = New Size(960, 600)
            FormBorderStyle = FormBorderStyle.None
            StartPosition = FormStartPosition.CenterScreen
            ShowIcon = False
            ShowInTaskbar = False
            KeyPreview = True

            Dim sectionLabel As New Label() With {
                .Text = sectionTitle,
                .Location = New Point(34, 16),
                .AutoSize = True,
                .Font = New Font("Microsoft YaHei UI", 10.0!, FontStyle.Regular),
                .ForeColor = Color.FromArgb(107, 114, 128)
            }
            Controls.Add(sectionLabel)

            Dim card As New NewPageCardPanel() With {
                .Location = New Point(32, 42),
                .Size = New Size(896, 520)
            }
            Controls.Add(card)
            NewPageUiPainter.EnableFormDrag(Me, Me, sectionLabel, card)

            Dim closeButton As New Button() With {
                .Location = New Point(card.Width - 48, 14),
                .Size = New Size(34, 34)
            }
            NewPageUiPainter.StyleCloseButton(closeButton)
            AddHandler closeButton.Click, Sub()
                                             Hide()
                                         End Sub
            card.Controls.Add(closeButton)

            _spinner.Location = New Point((card.Width - _spinner.Width) \ 2, 196)
            card.Controls.Add(_spinner)

            _marquee.Location = New Point((card.Width - _marquee.Width) \ 2, _spinner.Bottom + 28)
            card.Controls.Add(_marquee)

            _statusLabel.Text = initialStatus
            _statusLabel.Location = New Point(0, _marquee.Bottom + 24)
            _statusLabel.Size = New Size(card.Width, 32)
            _statusLabel.TextAlign = ContentAlignment.MiddleCenter
            _statusLabel.AutoEllipsis = True
            _statusLabel.Font = New Font("Microsoft YaHei UI", 11.0!, FontStyle.Bold)
            _statusLabel.ForeColor = OfficeAIStyleHelper.BrandPrimary
            card.Controls.Add(_statusLabel)

            Dim noteLabel As New Label() With {
                .Text = "即便是AI也不能保证百分百正确，请注意甄别并核实。",
                .Location = New Point(30, card.Height - 34),
                .Size = New Size(card.Width - 60, 22),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Font = New Font("Microsoft YaHei UI", 9.0!, FontStyle.Regular),
                .ForeColor = Color.FromArgb(107, 114, 128)
            }
            card.Controls.Add(noteLabel)

            AddHandler KeyDown, Sub(sender, e)
                                    If e.KeyCode = Keys.Escape Then Hide()
                                End Sub
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            Opacity = 0
            _spinner.Start()
            _marquee.Start()
            _fadeTimer.Start()
        End Sub

        Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
            _spinner.Stop()
            _marquee.Stop()
            MyBase.OnFormClosed(e)
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _fadeTimer.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Sub FadeTimer_Tick(sender As Object, e As EventArgs)
            Opacity = Math.Min(1, Opacity + 0.08)
            If Opacity >= 1 Then _fadeTimer.Stop()
        End Sub
    End Class

    Private Function ResolveDocmeePptxId() As String
        Try
            Dim presentation = Globals.ThisAddIn.Application.ActivePresentation
            If presentation IsNot Nothing Then
                Dim taggedId = ""
                Try
                    taggedId = presentation.Tags.Item(DocmeePptxIdTagName)
                Catch
                End Try

                If Not String.IsNullOrWhiteSpace(taggedId) Then Return taggedId.Trim()
            End If
        Catch
        End Try

        Return If(_lastReplaceSlidePptxId, "").Trim()
    End Function

    Private Sub SaveDocmeePptxId(pptxId As String)
        If String.IsNullOrWhiteSpace(pptxId) Then Return

        _lastReplaceSlidePptxId = pptxId.Trim()
        Try
            Dim presentation = Globals.ThisAddIn.Application.ActivePresentation
            If presentation Is Nothing Then Return

            Try
                presentation.Tags.Delete(DocmeePptxIdTagName)
            Catch
            End Try

            presentation.Tags.Add(DocmeePptxIdTagName, pptxId.Trim())
        Catch ex As Exception
            Debug.WriteLine("保存 Docmee PPT ID 失败: " & ex.Message)
        End Try
    End Sub

    Private Class PptxSlideChoice
        Public Property SlideIndex As Integer
        Public Property Title As String
        Public Property PptxPath As String
        Public Property PreviewPath As String

        Public Overrides Function ToString() As String
            Dim displayTitle = If(String.IsNullOrWhiteSpace(Title), "未命名页面", Title.Trim())
            Return $"第 {SlideIndex} 页 - {displayTitle}"
        End Function
    End Class

    Private Class ReplacementThumbCard
        Inherits Panel

        Private ReadOnly _pictureBox As New PictureBox()
        Private ReadOnly _titleLabel As New Label()
        Private _selected As Boolean

        Public Event CardClicked(sender As Object, e As EventArgs)

        Public ReadOnly Property Choice As PptxSlideChoice

        Public Sub New(choice As PptxSlideChoice, index As Integer, total As Integer)
            Me.Choice = choice
            Size = New Size(235, 164)
            Margin = New Padding(0, 0, 0, 16)
            BackColor = Color.White
            Cursor = Cursors.Hand
            SetStyle(ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.ResizeRedraw Or
                     ControlStyles.UserPaint, True)

            _pictureBox.Location = New Point(10, 10)
            _pictureBox.Size = New Size(215, 121)
            _pictureBox.BackColor = Color.FromArgb(249, 250, 251)
            _pictureBox.SizeMode = PictureBoxSizeMode.Zoom
            _pictureBox.Cursor = Cursors.Hand
            Controls.Add(_pictureBox)

            _titleLabel.Location = New Point(10, 136)
            _titleLabel.Size = New Size(215, 22)
            _titleLabel.AutoEllipsis = True
            _titleLabel.TextAlign = ContentAlignment.MiddleLeft
            _titleLabel.Font = New Font("Microsoft YaHei UI", 7.5!, FontStyle.Regular)
            _titleLabel.ForeColor = Color.FromArgb(75, 85, 99)
            _titleLabel.Text = GetDisplayTitle(index, total)
            _titleLabel.Cursor = Cursors.Hand
            Controls.Add(_titleLabel)

            AddHandler Click, AddressOf RaiseCardClicked
            AddHandler _pictureBox.Click, AddressOf RaiseCardClicked
            AddHandler _titleLabel.Click, AddressOf RaiseCardClicked
        End Sub

        Public Sub SetSelected(value As Boolean)
            _selected = value
            Invalidate()
        End Sub

        Public Sub SetLoading()
            If _pictureBox.Image Is Nothing Then
                _pictureBox.BackColor = Color.FromArgb(248, 250, 252)
            End If
        End Sub

        Public Sub SetPreview(previewPath As String)
            DisposePreview()
            If String.IsNullOrWhiteSpace(previewPath) OrElse Not File.Exists(previewPath) Then
                _pictureBox.BackColor = Color.FromArgb(248, 250, 252)
                Return
            End If

            Using previewImage = Image.FromFile(previewPath)
                _pictureBox.Image = New Bitmap(previewImage)
            End Using
            _pictureBox.BackColor = Color.White
        End Sub

        Public Sub DisposePreview()
            If _pictureBox.Image Is Nothing Then Return
            Dim oldImage = _pictureBox.Image
            _pictureBox.Image = Nothing
            oldImage.Dispose()
        End Sub

        Private Function GetDisplayTitle(index As Integer, total As Integer) As String
            Dim title = If(Choice Is Nothing OrElse String.IsNullOrWhiteSpace(Choice.Title), "未命名页面", Choice.Title.Trim())
            Return $"第 {index}/{total} 页  {title}"
        End Function

        Private Sub RaiseCardClicked(sender As Object, e As EventArgs)
            RaiseEvent CardClicked(Me, EventArgs.Empty)
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias

            Using brush As New SolidBrush(Color.White)
                g.FillRectangle(brush, ClientRectangle)
            End Using

            Dim imageRect = New Rectangle(_pictureBox.Left - 3, _pictureBox.Top - 3, _pictureBox.Width + 6, _pictureBox.Height + 6)
            Using imagePath = NewPageUiPainter.CreateRoundedRect(imageRect, 12)
                If _selected Then
                    Using pen As New Pen(Color.FromArgb(100, 55, 229), 3.0F)
                        g.DrawPath(pen, imagePath)
                    End Using
                End If
            End Using
        End Sub
    End Class

    Private Class ReplacementSlideChoiceForm
        Inherits Form

        Private ReadOnly _choices As New List(Of PptxSlideChoice)()
        Private ReadOnly _previewLoader As Func(Of PptxSlideChoice, String)
        Private ReadOnly _cards As New Dictionary(Of PptxSlideChoice, ReplacementThumbCard)()
        Private ReadOnly _mainPreviewBox As New PictureBox()
        Private ReadOnly _mainTitleLabel As New Label()
        Private ReadOnly _thumbFlow As New FlowLayoutPanel()
        Private ReadOnly _insertButton As New NewPageGradientButton()
        Private ReadOnly _pageLabel As New Label()
        Private _selectedChoice As PptxSlideChoice
        Private _loadingPreviews As Boolean

        Public ReadOnly Property SelectedChoice As PptxSlideChoice
            Get
                Return _selectedChoice
            End Get
        End Property

        Public Sub New(choices As IEnumerable(Of PptxSlideChoice),
                       previewLoader As Func(Of PptxSlideChoice, String))
            If choices IsNot Nothing Then
                For Each choice In choices
                    If choice IsNot Nothing Then _choices.Add(choice)
                Next
            End If
            _previewLoader = previewLoader
            BuildLayout()
        End Sub

        Private Sub BuildLayout()
            Text = "选择AI生成单页"
            Font = OfficeAIStyleHelper.FontUi
            BackColor = Color.FromArgb(245, 246, 250)
            FormBorderStyle = FormBorderStyle.None
            StartPosition = FormStartPosition.CenterScreen
            ShowIcon = False
            ShowInTaskbar = False
            KeyPreview = True

            Dim workingArea = Screen.FromPoint(Cursor.Position).WorkingArea
            Dim formWidth = Math.Min(1500, Math.Max(1100, workingArea.Width - 90))
            Dim formHeight = Math.Min(860, Math.Max(720, workingArea.Height - 90))
            ClientSize = New Size(formWidth, formHeight)

            Dim sectionLabel As New Label() With {
                .Text = "新页面",
                .Location = New Point(18, 8),
                .AutoSize = True,
                .Font = New Font("Microsoft YaHei UI", 10.0!, FontStyle.Regular),
                .ForeColor = Color.FromArgb(107, 114, 128)
            }
            Controls.Add(sectionLabel)

            Dim card As New NewPageCardPanel() With {
                .Location = New Point(10, 32),
                .Size = New Size(ClientSize.Width - 20, ClientSize.Height - 54)
            }
            Controls.Add(card)
            NewPageUiPainter.EnableFormDrag(Me, Me, sectionLabel, card)

            Dim closeButton As New Button() With {
                .Location = New Point(card.Width - 48, 14),
                .Size = New Size(34, 34),
                .DialogResult = DialogResult.Cancel
            }
            NewPageUiPainter.StyleCloseButton(closeButton)
            AddHandler closeButton.Click, Sub()
                                             DialogResult = DialogResult.Cancel
                                             Close()
                                         End Sub
            card.Controls.Add(closeButton)
            CancelButton = closeButton

            Dim rightPanelWidth = 245
            Dim rightPanelLeft = card.Width - rightPanelWidth - 38
            Dim separatorX = rightPanelLeft - 50
            Dim previewLeft = 92
            Dim previewTop = 112
            Dim previewWidth = Math.Max(660, separatorX - previewLeft - 96)
            Dim previewHeight = CInt(previewWidth * 9 / 16)
            Dim maxPreviewHeight = Math.Max(380, card.Height - previewTop - 190)
            If previewHeight > maxPreviewHeight Then
                previewHeight = maxPreviewHeight
                previewWidth = CInt(previewHeight * 16 / 9)
            End If

            Dim previewHost As New SlidePreviewFramePanel() With {
                .Location = New Point(previewLeft, previewTop),
                .Size = New Size(previewWidth + 18, previewHeight + 18)
            }
            card.Controls.Add(previewHost)

            _mainPreviewBox.Location = New Point(0, 0)
            _mainPreviewBox.Size = New Size(previewWidth, previewHeight)
            _mainPreviewBox.BackColor = Color.White
            _mainPreviewBox.SizeMode = PictureBoxSizeMode.Zoom
            previewHost.Controls.Add(_mainPreviewBox)

            _mainTitleLabel.Location = New Point(previewLeft, previewHost.Bottom + 6)
            _mainTitleLabel.Size = New Size(previewWidth, 22)
            _mainTitleLabel.TextAlign = ContentAlignment.MiddleCenter
            _mainTitleLabel.AutoEllipsis = True
            _mainTitleLabel.Font = New Font("Microsoft YaHei UI", 9.5!, FontStyle.Regular)
            _mainTitleLabel.ForeColor = Color.FromArgb(100, 116, 139)
            _mainTitleLabel.Visible = False
            card.Controls.Add(_mainTitleLabel)

            _insertButton.Text = "✓  将此页面插入演示文档"
            Dim insertButtonY = Math.Min(previewHost.Bottom + 24, card.Height - 112)
            _insertButton.Location = New Point(previewLeft + (previewWidth - 280) \ 2, insertButtonY)
            _insertButton.Size = New Size(280, 48)
            AddHandler _insertButton.Click, AddressOf InsertButton_Click
            card.Controls.Add(_insertButton)
            AcceptButton = _insertButton

            Dim separator As New Label() With {
                .Location = New Point(separatorX, 88),
                .Size = New Size(1, card.Height - 174),
                .BackColor = Color.FromArgb(229, 231, 235),
                .AutoSize = False
            }
            card.Controls.Add(separator)

            Dim verticalHint As New Label() With {
                .Text = "来" & vbCrLf & "试" & vbCrLf & "试" & vbCrLf & "其" & vbCrLf & "他" & vbCrLf & "的" & vbCrLf & "页" & vbCrLf & "面",
                .Location = New Point(separatorX + 8, Math.Max(120, (card.Height - 180) \ 2)),
                .Size = New Size(24, 180),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Font = New Font("Microsoft YaHei UI", 9.0!, FontStyle.Regular),
                .ForeColor = Color.FromArgb(156, 163, 175)
            }
            card.Controls.Add(verticalHint)

            Dim candidateTitle As New Label() With {
                .Text = "候选页面",
                .Location = New Point(rightPanelLeft, 60),
                .Size = New Size(rightPanelWidth, 32),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Font = New Font("Microsoft YaHei UI", 15.0!, FontStyle.Regular),
                .ForeColor = Color.FromArgb(55, 65, 81)
            }
            card.Controls.Add(candidateTitle)
            NewPageUiPainter.EnableFormDrag(Me, candidateTitle)

            _thumbFlow.Location = New Point(rightPanelLeft, 106)
            _thumbFlow.Size = New Size(rightPanelWidth, card.Height - 192)
            _thumbFlow.BackColor = Color.White
            _thumbFlow.AutoScroll = True
            _thumbFlow.FlowDirection = FlowDirection.TopDown
            _thumbFlow.WrapContents = False
            _thumbFlow.Padding = New Padding(0, 0, 0, 8)
            card.Controls.Add(_thumbFlow)

            Dim noteLabel As New Label() With {
                .Text = "即便是AI也不能保证百分百正确，请注意甄别并核实。",
                .Location = New Point(30, card.Height - 34),
                .Size = New Size(card.Width - 60, 22),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Font = New Font("Microsoft YaHei UI", 9.0!, FontStyle.Regular),
                .ForeColor = Color.FromArgb(107, 114, 128)
            }
            card.Controls.Add(noteLabel)
            NewPageUiPainter.EnableFormDrag(Me, noteLabel)

            _pageLabel.Location = New Point(card.Width - 158, card.Height - 34)
            _pageLabel.Size = New Size(100, 22)
            _pageLabel.TextAlign = ContentAlignment.MiddleRight
            _pageLabel.Font = New Font("Microsoft YaHei UI", 9.0!, FontStyle.Regular)
            _pageLabel.ForeColor = Color.FromArgb(224, 226, 232)
            card.Controls.Add(_pageLabel)

            For index As Integer = 0 To _choices.Count - 1
                Dim thumb = New ReplacementThumbCard(_choices(index), index + 1, _choices.Count)
                AddHandler thumb.CardClicked, AddressOf ThumbCard_Click
                _cards(_choices(index)) = thumb
                _thumbFlow.Controls.Add(thumb)
            Next

            AddHandler KeyDown, Sub(sender, e)
                                    If e.KeyCode = Keys.Escape Then
                                        DialogResult = DialogResult.Cancel
                                        Close()
                                    End If
                                End Sub
            AddHandler FormClosed, AddressOf ReplacementSlideChoiceForm_FormClosed
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            Dim defaultIndex = If(_choices.Count > 0, 0, -1)
            If defaultIndex >= 0 Then
                SelectChoice(_choices(defaultIndex))
                If _cards.ContainsKey(_choices(defaultIndex)) Then
                    _thumbFlow.ScrollControlIntoView(_cards(_choices(defaultIndex)))
                End If
            End If

            BeginInvoke(CType(Sub() LoadAllThumbPreviews(), MethodInvoker))
        End Sub

        Private Sub ThumbCard_Click(sender As Object, e As EventArgs)
            Dim thumb = TryCast(sender, ReplacementThumbCard)
            If thumb Is Nothing Then Return
            SelectChoice(thumb.Choice)
        End Sub

        Private Sub SelectChoice(choice As PptxSlideChoice)
            If choice Is Nothing Then Return
            _selectedChoice = choice
            For Each pair In _cards
                pair.Value.SetSelected(Object.ReferenceEquals(pair.Key, choice))
            Next
            Dim pageIndex = _choices.IndexOf(choice)
            _pageLabel.Text = If(pageIndex >= 0, (pageIndex + 1).ToString() & " / " & _choices.Count.ToString(), "")
            LoadMainPreview(choice)
        End Sub

        Private Sub LoadMainPreview(choice As PptxSlideChoice)
            DisposeMainPreview()
            If choice Is Nothing Then
                _mainTitleLabel.Text = ""
                Return
            End If

            _mainTitleLabel.Text = choice.ToString() & "    正在生成预览..."
            Cursor = Cursors.WaitCursor
            _mainPreviewBox.Refresh()
            _mainTitleLabel.Refresh()
            Application.DoEvents()

            Try
                Dim previewPath = LoadPreviewPath(choice)
                If Not String.IsNullOrWhiteSpace(previewPath) AndAlso File.Exists(previewPath) Then
                    Using previewImage = Image.FromFile(previewPath)
                        _mainPreviewBox.Image = New Bitmap(previewImage)
                    End Using
                    If _cards.ContainsKey(choice) Then _cards(choice).SetPreview(previewPath)
                    _mainTitleLabel.Text = choice.ToString()
                Else
                    _mainTitleLabel.Text = choice.ToString() & "    预览图生成失败"
                End If
            Finally
                Cursor = Cursors.Default
            End Try
        End Sub

        Private Sub LoadAllThumbPreviews()
            If _loadingPreviews Then Return
            _loadingPreviews = True
            Try
                For Each choice In _choices
                    If IsDisposed Then Return
                    If Not _cards.ContainsKey(choice) Then Continue For

                    Dim thumb = _cards(choice)
                    thumb.SetLoading()
                    thumb.Refresh()
                    Application.DoEvents()

                    Dim previewPath = LoadPreviewPath(choice)
                    thumb.SetPreview(previewPath)
                    If Object.ReferenceEquals(choice, _selectedChoice) AndAlso _mainPreviewBox.Image Is Nothing Then
                        LoadMainPreview(choice)
                    End If
                Next
            Finally
                _loadingPreviews = False
            End Try
        End Sub

        Private Function LoadPreviewPath(choice As PptxSlideChoice) As String
            If choice Is Nothing Then Return ""
            If _previewLoader Is Nothing Then Return ""
            Return _previewLoader(choice)
        End Function

        Private Sub InsertButton_Click(sender As Object, e As EventArgs)
            If _selectedChoice Is Nothing Then
                MessageBox.Show("请先选择一个候选页面。", "选择AI生成单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            DialogResult = DialogResult.OK
            Close()
        End Sub

        Private Sub DisposeMainPreview()
            If _mainPreviewBox.Image Is Nothing Then Return
            Dim oldImage = _mainPreviewBox.Image
            _mainPreviewBox.Image = Nothing
            oldImage.Dispose()
        End Sub

        Private Sub ReplacementSlideChoiceForm_FormClosed(sender As Object, e As FormClosedEventArgs)
            DisposeMainPreview()
            For Each thumb In _cards.Values
                thumb.DisposePreview()
            Next
        End Sub
    End Class

    Private Function ReplaceCurrentSlideWithSelectedSlideFromPptx(originalSlide As PowerPoint.Slide, pptxPath As String) As Boolean
        If originalSlide Is Nothing Then Throw New ArgumentNullException(NameOf(originalSlide))
        If String.IsNullOrWhiteSpace(pptxPath) OrElse Not File.Exists(pptxPath) Then
            Throw New FileNotFoundException("未找到 Docmee 生成的 PPTX 文件。", pptxPath)
        End If

        Dim presentation = Globals.ThisAddIn.Application.ActivePresentation
        Dim originalIndex = originalSlide.SlideIndex
        Dim choices = GetPptxSlideChoices(pptxPath)
        If choices.Count = 0 Then Throw New InvalidOperationException("Docmee 生成的 PPTX 中没有可导入的幻灯片。")

        Dim selectedSlideIndex = choices(choices.Count - 1).SlideIndex
        If choices.Count > 1 Then
            Dim selectedChoice = ShowReplacementSlideChoiceDialog(choices)
            If selectedChoice Is Nothing Then
                LogInfo("[ReplaceSlide] User canceled replacement slide selection.")
                Return False
            End If
            selectedSlideIndex = selectedChoice.SlideIndex
        End If

        LogInfo("[ReplaceSlide] Import selected slide index=" & selectedSlideIndex.ToString())
        Dim insertedCount = presentation.Slides.InsertFromFile(pptxPath, originalIndex, selectedSlideIndex, selectedSlideIndex)
        If insertedCount <= 0 Then Throw New InvalidOperationException("导入 Docmee 单页失败。")

        originalSlide.Delete()

        Try
            Dim targetIndex = Math.Max(1, Math.Min(originalIndex, presentation.Slides.Count))
            presentation.Slides(targetIndex).Select()
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(targetIndex)
        Catch
        End Try

        Return True
    End Function

    Private Function GetPptxSlideChoices(pptxPath As String) As List(Of PptxSlideChoice)
        Dim sourcePresentation As PowerPoint.Presentation = Nothing
        Dim choices As New List(Of PptxSlideChoice)()
        Try
            sourcePresentation = Globals.ThisAddIn.Application.Presentations.Open(
                pptxPath,
                ReadOnly:=Microsoft.Office.Core.MsoTriState.msoTrue,
                Untitled:=Microsoft.Office.Core.MsoTriState.msoFalse,
                WithWindow:=Microsoft.Office.Core.MsoTriState.msoFalse)

            For slideIndex As Integer = 1 To sourcePresentation.Slides.Count
                choices.Add(New PptxSlideChoice() With {
                    .SlideIndex = slideIndex,
                    .Title = GetSlideTitleText(sourcePresentation.Slides(slideIndex)),
                    .PptxPath = pptxPath
                })
            Next
        Finally
            If sourcePresentation IsNot Nothing Then
                sourcePresentation.Close()
            End If
        End Try

        Return choices
    End Function

    Private Function GetSlideTitleText(slide As PowerPoint.Slide) As String
        If slide Is Nothing Then Return ""

        Try
            If slide.Shapes.HasTitle = Microsoft.Office.Core.MsoTriState.msoTrue Then
                Dim titleText = slide.Shapes.Title.TextFrame.TextRange.Text
                If Not String.IsNullOrWhiteSpace(titleText) Then Return titleText.Trim()
            End If
        Catch
        End Try

        Try
            For shapeIndex As Integer = 1 To slide.Shapes.Count
                Dim shape = slide.Shapes(shapeIndex)
                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                   shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    Dim text = shape.TextFrame.TextRange.Text
                    If Not String.IsNullOrWhiteSpace(text) Then Return text.Trim()
                End If
            Next
        Catch
        End Try

        Return ""
    End Function

    Private Function EnsureSlidePreview(choice As PptxSlideChoice) As String
        If choice Is Nothing Then Return ""
        If Not String.IsNullOrWhiteSpace(choice.PreviewPath) AndAlso File.Exists(choice.PreviewPath) Then Return choice.PreviewPath
        If String.IsNullOrWhiteSpace(choice.PptxPath) OrElse Not File.Exists(choice.PptxPath) Then Return ""

        Dim sourcePresentation As PowerPoint.Presentation = Nothing
        Try
            sourcePresentation = Globals.ThisAddIn.Application.Presentations.Open(
                choice.PptxPath,
                ReadOnly:=Microsoft.Office.Core.MsoTriState.msoTrue,
                Untitled:=Microsoft.Office.Core.MsoTriState.msoFalse,
                WithWindow:=Microsoft.Office.Core.MsoTriState.msoFalse)

            Dim previewPath = Path.Combine(Path.GetTempPath(), $"wenduoduoAI_newpage_preview_{Guid.NewGuid():N}_{choice.SlideIndex}.png")
            sourcePresentation.Slides(choice.SlideIndex).Export(previewPath, "PNG", 960, 540)
            choice.PreviewPath = previewPath
            Return previewPath
        Catch ex As Exception
            LogInfo("[ReplaceSlide] Export preview failed. slideIndex=" & choice.SlideIndex.ToString() & ", message=" & ex.Message)
            Return ""
        Finally
            If sourcePresentation IsNot Nothing Then
                sourcePresentation.Close()
            End If
        End Try
    End Function

    Private Function ShowReplacementSlideChoiceDialog(choices As List(Of PptxSlideChoice)) As PptxSlideChoice
        Using dialog As New ReplacementSlideChoiceForm(choices, AddressOf EnsureSlidePreview)
            If dialog.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return dialog.SelectedChoice
        End Using
    End Function

    Private Async Function BeautifyCurrentSlideWithDocmeeTemplateAsync() As Task
        LogInfo("[BeautifyTemplate] Button clicked.")
        Dim originalSlide = GetCurrentSlide()
        If originalSlide Is Nothing Then
            MessageBox.Show("请先选中要美化的幻灯片。", "美化模板", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim slideContent = GetSlidePlainText(originalSlide)
        If String.IsNullOrWhiteSpace(slideContent) Then
            MessageBox.Show("当前页没有可用于美化的文本内容。", "美化模板", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        LogInfo("[BeautifyTemplate] NewPageWithAi contentLength=" & slideContent.Length.ToString())

        Dim client As New DocmeePptClient()
        Dim progressWindow = CreateReplaceProgressWindow("美化单页", "正在生成美化页面...")
        progressWindow.Item1.Text = "正在美化当前页"
        progressWindow.Item1.Show()
        UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在生成美化页面...")

        Try
            ShareRibbon.GlobalStatusStripAll.ShowProgress("AI内容提效：正在启动美化当前页...")
            Dim pptxId = ResolveDocmeePptxId()
            If String.IsNullOrWhiteSpace(pptxId) Then
                pptxId = DocmeePptClient.GetRandomNewPageTemplateId()
                LogInfo("[BeautifyTemplate] No presentation Docmee PPT ID, using random ID=" & pptxId)
            End If

            ShareRibbon.GlobalStatusStripAll.ShowProgress("正在用 Docmee 美化当前页...")
            LogInfo("[BeautifyTemplate] Calling Docmee newPageWithAiV2.")
            Dim result = Await client.NewPageWithAiV2Async(
                slideContent,
                pptxId,
                Sub(message)
                    If Not String.IsNullOrWhiteSpace(message) Then
                        ShareRibbon.GlobalStatusStripAll.ShowProgress("Docmee 美化生成中...")
                        UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在生成美化页面...")
                    End If
                End Sub)

            Dim resultPptxId = If(String.IsNullOrWhiteSpace(result.PptxId), pptxId, result.PptxId)
            Dim fileUrl = result.FileUrl
            LogInfo("[BeautifyTemplate] Docmee result. pptxId=" & resultPptxId & ", hasFileUrl=" & Not String.IsNullOrWhiteSpace(fileUrl))
            If String.IsNullOrWhiteSpace(fileUrl) Then
                UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在获取美化 PPTX...")
                fileUrl = Await client.DownloadPptxAsync(resultPptxId, True, 6, 1200)
            End If

            Dim localPath = Path.Combine(Path.GetTempPath(), $"wenduoduoAI_beautify_{resultPptxId}_{Guid.NewGuid():N}.pptx")
            UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在下载美化结果...")
            Await client.DownloadPptxFileAsync(fileUrl, localPath)

            UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在准备替换页面...")
            CloseReplaceProgressWindow(progressWindow.Item1)

            If Not ReplaceCurrentSlideWithSelectedSlideFromPptx(originalSlide, localPath) Then Return

            _lastReplaceSlidePptxId = resultPptxId
            SaveDocmeePptxId(resultPptxId)

            ShareRibbon.GlobalStatusStripAll.ShowProgress("美化模板完成")
        Finally
            CloseReplaceProgressWindow(progressWindow.Item1)
        End Try
    End Function

    Private Function GetSlidePlainText(slide As PowerPoint.Slide) As String
        If slide Is Nothing Then Return ""

        Dim lines As New List(Of String)()
        Try
            For i = 1 To slide.Shapes.Count
                CollectShapeContextText(slide.Shapes(i), lines)
            Next
        Catch ex As Exception
            Debug.WriteLine("读取当前页文本失败: " & ex.Message)
        End Try

        Return String.Join(vbCrLf, lines.Where(Function(line) Not String.IsNullOrWhiteSpace(line)).Distinct())
    End Function

    Private Function GetSelectedPptTextTargets(modeName As String) As List(Of PptTextTarget)
        Dim targets As New List(Of PptTextTarget)()
        Dim allowBlankTextFrame = String.Equals(modeName, "填充", StringComparison.Ordinal)
        Dim slideContextText = GetCurrentSlideContextText()

        Try
            Dim sel = Globals.ThisAddIn.Application.ActiveWindow.Selection

            Select Case sel.Type
                Case PowerPoint.PpSelectionType.ppSelectionText
                    Dim text = sel.TextRange.Text
                    If allowBlankTextFrame OrElse Not String.IsNullOrWhiteSpace(text) Then
                        Dim selectedShape As PowerPoint.Shape = Nothing
                        Try
                            selectedShape = sel.ShapeRange(1)
                        Catch
                        End Try

                        targets.Add(New PptTextTarget() With {
                            .TextRange = sel.TextRange,
                            .Shape = selectedShape,
                            .OriginalText = If(text, "").Trim(),
                            .SlideContextText = slideContextText
                        })
                    End If

                Case PowerPoint.PpSelectionType.ppSelectionShapes
                    For i = 1 To sel.ShapeRange.Count
                        CollectShapeTextTargets(sel.ShapeRange(i), targets, allowBlankTextFrame, slideContextText)
                    Next
            End Select
        Catch ex As Exception
            Debug.WriteLine("读取 PPT 选中文本失败: " & ex.Message)
        End Try

        Return targets
    End Function

    Private Sub CollectShapeTextTargets(shape As PowerPoint.Shape, targets As List(Of PptTextTarget), allowBlankTextFrame As Boolean, slideContextText As String)
        Try
            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoGroup Then
                For i = 1 To shape.GroupItems.Count
                    CollectShapeTextTargets(shape.GroupItems(i), targets, allowBlankTextFrame, slideContextText)
                Next
                Return
            End If

            If shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                Dim table = shape.Table
                For row = 1 To table.Rows.Count
                    For col = 1 To table.Columns.Count
                        Dim cellShape = table.Cell(row, col).Shape
                        If cellShape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Dim cellText = ""
                            If cellShape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                cellText = cellShape.TextFrame.TextRange.Text
                            End If

                            If allowBlankTextFrame OrElse Not String.IsNullOrWhiteSpace(cellText) Then
                                targets.Add(New PptTextTarget() With {
                                    .TextRange = cellShape.TextFrame.TextRange,
                                    .Shape = cellShape,
                                    .OriginalText = If(cellText, "").Trim(),
                                    .SlideContextText = slideContextText
                                })
                            End If
                        End If
                    Next
                Next
                Return
            End If

            If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                Dim text = shape.TextFrame.TextRange.Text
                If allowBlankTextFrame OrElse Not String.IsNullOrWhiteSpace(text) Then
                    targets.Add(New PptTextTarget() With {
                        .TextRange = shape.TextFrame.TextRange,
                        .Shape = shape,
                        .OriginalText = text.Trim(),
                        .SlideContextText = slideContextText
                    })
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine("收集形状文本失败: " & ex.Message)
        End Try
    End Sub

    Private Function GetCurrentSlideContextText() As String
        Dim slide = GetCurrentSlide()
        If slide Is Nothing Then Return ""

        Dim lines As New List(Of String)()
        Try
            For i = 1 To slide.Shapes.Count
                CollectShapeContextText(slide.Shapes(i), lines)
            Next
        Catch ex As Exception
            Debug.WriteLine("收集当前页上下文失败: " & ex.Message)
        End Try

        Return String.Join(vbCrLf, lines.Distinct().Take(12))
    End Function

    Private Sub CollectShapeContextText(shape As PowerPoint.Shape, lines As List(Of String))
        Try
            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoGroup Then
                For i = 1 To shape.GroupItems.Count
                    CollectShapeContextText(shape.GroupItems(i), lines)
                Next
                Return
            End If

            If shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                Dim table = shape.Table
                For row = 1 To table.Rows.Count
                    For col = 1 To table.Columns.Count
                        CollectShapeContextText(table.Cell(row, col).Shape, lines)
                    Next
                Next
                Return
            End If

            If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
               shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                Dim text = shape.TextFrame.TextRange.Text
                If Not String.IsNullOrWhiteSpace(text) Then
                    lines.Add(text.Trim())
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine("收集形状上下文失败: " & ex.Message)
        End Try
    End Sub

    Private Function GetCurrentSlide() As PowerPoint.Slide
        Try
            Dim sel = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If sel IsNot Nothing AndAlso sel.SlideRange IsNot Nothing AndAlso sel.SlideRange.Count > 0 Then
                Return sel.SlideRange(1)
            End If
        Catch
        End Try

        Try
            Return TryCast(Globals.ThisAddIn.Application.ActiveWindow.View.Slide, PowerPoint.Slide)
        Catch
        End Try

        Return Nothing
    End Function

    Private Sub AutoFitPptTextShape(shape As PowerPoint.Shape)
        Try
            If shape Is Nothing OrElse
               shape.HasTextFrame <> Microsoft.Office.Core.MsoTriState.msoTrue OrElse
               shape.TextFrame.HasText <> Microsoft.Office.Core.MsoTriState.msoTrue Then Return

            Dim textFrame = shape.TextFrame
            Dim textRange = textFrame.TextRange
            Dim currentFontSize As Single = CSng(textRange.Font.Size)
            If currentFontSize <= 0 Then currentFontSize = 18.0F

            Dim minFontSize As Single = 10.0F
            Dim targetWidth As Single = CSng(Math.Max(1.0, shape.Width - textFrame.MarginLeft - textFrame.MarginRight))
            Dim targetHeight As Single = CSng(Math.Max(1.0, shape.Height - textFrame.MarginTop - textFrame.MarginBottom))
            Dim guard As Integer = 0

            While guard < 12 AndAlso currentFontSize > minFontSize AndAlso PptTextOverflows(textRange, targetWidth, targetHeight)
                currentFontSize = CSng(Math.Max(minFontSize, currentFontSize - 1.0F))
                textRange.Font.Size = currentFontSize
                guard += 1
            End While

            If PptTextOverflows(textRange, targetWidth, targetHeight) Then
                textFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText
            End If
        Catch ex As Exception
            Debug.WriteLine("适配文本框失败: " & ex.Message)
        End Try
    End Sub

    Private Function PptTextOverflows(textRange As PowerPoint.TextRange, targetWidth As Single, targetHeight As Single) As Boolean
        Try
            Return textRange.BoundHeight > targetHeight OrElse textRange.BoundWidth > targetWidth * 1.08F
        Catch
            Return False
        End Try
    End Function

    Private Function ShowAiCreationDialog() As AiCreationOptions
        Using dialog As New Form()
            OfficeAIStyleHelper.StyleFormDialog(dialog)
            dialog.Text = "AI创作"
            dialog.StartPosition = FormStartPosition.CenterScreen
            dialog.ClientSize = New Size(540, 590)
            dialog.BackColor = OfficeAIStyleHelper.BgPage

            Dim heroPanel As New Panel() With {
                .Location = New Point(OfficeAIStyleHelper.SpacingLg, OfficeAIStyleHelper.SpacingLg),
                .Size = New Size(dialog.ClientSize.Width - OfficeAIStyleHelper.SpacingLg * 2, 88),
                .BackColor = OfficeAIStyleHelper.BgSurface
            }
            dialog.Controls.Add(heroPanel)

            Dim heroIcon As New Label() With {
                .Text = "Ai",
                .Location = New Point(18, 20),
                .Size = New Size(46, 46),
                .BackColor = OfficeAIStyleHelper.BrandPrimaryLight,
                .ForeColor = OfficeAIStyleHelper.BrandPrimary,
                .Font = New Font("Microsoft YaHei UI", 15.0F, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleCenter
            }
            heroPanel.Controls.Add(heroIcon)

            Dim heroTitle As New Label() With {
                .Text = "AI 创作",
                .Location = New Point(78, 18),
                .Size = New Size(heroPanel.Width - 190, 28),
                .ForeColor = OfficeAIStyleHelper.TextPrimary,
                .Font = New Font("Microsoft YaHei UI", 14.0F, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleLeft
            }
            heroPanel.Controls.Add(heroTitle)

            Dim heroSubtitle As New Label() With {
                .Text = "选中文本后，选择一种处理方式并替换回当前页面",
                .Location = New Point(78, 52),
                .Size = New Size(heroPanel.Width - 96, 22),
                .ForeColor = OfficeAIStyleHelper.TextSecondary,
                .Font = OfficeAIStyleHelper.FontUi,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            heroPanel.Controls.Add(heroSubtitle)

            Dim heroBadge As New Label() With {
                .Text = "内容提效",
                .Location = New Point(heroPanel.Width - 108, 18),
                .Size = New Size(88, 28),
                .BackColor = OfficeAIStyleHelper.BrandPrimaryLight,
                .ForeColor = OfficeAIStyleHelper.BrandPrimary,
                .Font = OfficeAIStyleHelper.FontUiBold,
                .TextAlign = ContentAlignment.MiddleCenter
            }
            heroPanel.Controls.Add(heroBadge)

            Dim contentY As Integer = heroPanel.Bottom + OfficeAIStyleHelper.SpacingMd

            ' 提示标签
            Dim label As New Label() With {
                .Text = "选择处理方式",
                .Location = New Point(OfficeAIStyleHelper.SpacingLg, contentY),
                .AutoSize = True
            }
            OfficeAIStyleHelper.StyleLabelBody(label)
            dialog.Controls.Add(label)

            ' 卡片式模式选择
            Dim cardY As Integer = contentY + 26
            Dim cardW As Integer = dialog.ClientSize.Width - OfficeAIStyleHelper.SpacingLg * 2
            Dim cardH As Integer = 78
            Dim cardGap As Integer = 10
            Dim col1X As Integer = OfficeAIStyleHelper.SpacingLg

            Dim cardPolish = OfficeAIStyleHelper.CreateOptionCard("润", "润色", "优化表达，让文字更自然、专业", cardW, cardH)
            cardPolish.Location = New Point(col1X, cardY)
            Dim cardExpand = OfficeAIStyleHelper.CreateOptionCard("扩", "扩写", "补充细节，扩展成更完整表述", cardW, cardH)
            cardExpand.Location = New Point(col1X, cardPolish.Bottom + cardGap)
            Dim cardShorten = OfficeAIStyleHelper.CreateOptionCard("缩", "缩写", "精简文字，保留核心信息和重点", cardW, cardH)
            cardShorten.Location = New Point(col1X, cardExpand.Bottom + cardGap)
            Dim cardTranslate = OfficeAIStyleHelper.CreateOptionCard("译", "翻译", "选择目标语言，保留原意并优化表达", cardW, cardH)
            cardTranslate.Location = New Point(col1X, cardShorten.Bottom + cardGap)

            ' 卡片点击事件 - 使用 Dictionary 映射卡片到模式名
            Dim allCards = {cardPolish, cardExpand, cardShorten, cardTranslate}
            Dim cardModes As New Dictionary(Of Panel, String) From {
                {cardPolish, "润色"}, {cardExpand, "扩写"}, {cardShorten, "缩写"}, {cardTranslate, "翻译"}
            }
            Dim selectedMode As String = "润色"
            OfficeAIStyleHelper.SetCardSelected(cardPolish, True)

            ' 目标语言选择行 (仅翻译时可见) - 必须在卡片点击事件之前声明
            Dim langY As Integer = cardTranslate.Bottom + OfficeAIStyleHelper.SpacingMd

            Dim languageLabel As New Label() With {
                .Text = "目标语言：",
                .Location = New Point(OfficeAIStyleHelper.SpacingLg, langY + 3),
                .Size = New Size(82, OfficeAIStyleHelper.InputHeight),
                .AutoSize = False,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Visible = False
            }
            OfficeAIStyleHelper.StyleLabelBody(languageLabel)
            languageLabel.AutoSize = False
            languageLabel.TextAlign = ContentAlignment.MiddleLeft
            dialog.Controls.Add(languageLabel)

            Dim languageCombo As New ComboBox() With {
                .Location = New Point(112, langY),
                .Size = New Size(dialog.ClientSize.Width - 128, OfficeAIStyleHelper.InputHeight),
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Visible = False
            }
            languageCombo.Items.AddRange(New Object() {
                "英语", "中文", "日语", "韩语", "法语",
                "德语", "西班牙语", "葡萄牙语", "俄语", "意大利语"
            })
            languageCombo.SelectedIndex = 0
            OfficeAIStyleHelper.StyleComboBox(languageCombo)
            dialog.Controls.Add(languageCombo)

            For Each card In allCards
                AddHandler card.Click, Sub(s, e)
                    Dim clickedCard = DirectCast(s, Panel)
                    For Each c In allCards
                        OfficeAIStyleHelper.SetCardSelected(c, c Is clickedCard)
                    Next
                    selectedMode = cardModes(clickedCard)
                    Dim isTranslate = String.Equals(selectedMode, "翻译", StringComparison.Ordinal)
                    languageLabel.Visible = isTranslate
                    languageCombo.Visible = isTranslate
                End Sub
                dialog.Controls.Add(card)
            Next

            ' 底部按钮
            Dim btnY As Integer = dialog.ClientSize.Height - OfficeAIStyleHelper.ButtonHeight - OfficeAIStyleHelper.SpacingLg - 8
            Dim sepHorizontal = OfficeAIStyleHelper.CreateSeparator(dialog.ClientSize.Width - OfficeAIStyleHelper.SpacingLg * 2)
            sepHorizontal.Location = New Point(OfficeAIStyleHelper.SpacingLg, btnY - OfficeAIStyleHelper.SpacingSm)
            dialog.Controls.Add(sepHorizontal)

            Dim okButton As New Button() With {
                .Text = "开始创作",
                .Location = New Point(dialog.ClientSize.Width - 208, btnY),
                .Size = New Size(110, OfficeAIStyleHelper.ButtonHeight),
                .DialogResult = DialogResult.OK
            }
            OfficeAIStyleHelper.StyleButtonPrimary(okButton)
            dialog.Controls.Add(okButton)
            dialog.AcceptButton = okButton

            Dim cancelButton As New Button() With {
                .Text = "取消",
                .Location = New Point(dialog.ClientSize.Width - 90, btnY),
                .Size = New Size(74, OfficeAIStyleHelper.ButtonHeight),
                .DialogResult = DialogResult.Cancel
            }
            OfficeAIStyleHelper.StyleButtonSecondary(cancelButton)
            dialog.Controls.Add(cancelButton)
            dialog.CancelButton = cancelButton

            If dialog.ShowDialog() <> DialogResult.OK Then Return Nothing
            Dim selectedLanguage = If(String.Equals(selectedMode, "翻译", StringComparison.Ordinal), languageCombo.SelectedItem.ToString(), "")
            Return New AiCreationOptions() With {
                .ModeName = selectedMode,
                .TargetLanguageName = selectedLanguage,
                .TargetLanguageCode = GetDocmeeLanguageCode(selectedLanguage)
            }
        End Using
    End Function

    Private Async Function OptimizeSelectedTextWithDocmeeAsync(targets As List(Of PptTextTarget), options As AiCreationOptions) As Task(Of Integer)
        Dim modeName = options.ModeName
        Dim optimizeType = GetDocmeeTextOptimizeType(modeName)
        Dim client As New DocmeePptClient()
        Dim uiContext = SynchronizationContext.Current
        Dim changedCount = 0

        ShareRibbon.GlobalStatusStripAll.ShowProgress($"正在进行 AI创作：{modeName}...")

        For Each target In targets
            Dim originalText = If(target.OriginalText, "").Trim()
            If String.IsNullOrWhiteSpace(originalText) Then Continue For

            Dim lastPreview As String = ""
            Dim optimizedText = Await client.OptimizeTextAsync(
                originalText,
                optimizeType,
                targetLanguageName:=options.TargetLanguageName,
                targetLanguageCode:=options.TargetLanguageCode,
                progressHandler:=Sub(partialText)
                                     If String.IsNullOrWhiteSpace(partialText) OrElse String.Equals(partialText, lastPreview, StringComparison.Ordinal) Then Return
                                     lastPreview = partialText
                                     PostToUi(uiContext,
                                              Sub()
                                                  Try
                                                      target.TextRange.Text = partialText
                                                      If target.Shape IsNot Nothing Then AutoFitPptTextShape(target.Shape)
                                                      ShareRibbon.GlobalStatusStripAll.ShowProgress($"AI创作中：{modeName}...")
                                                  Catch ex As Exception
                                                      Debug.WriteLine("流式更新 PPT 文本失败: " & ex.Message)
                                                  End Try
                                              End Sub)
                                 End Sub)

            If Not String.IsNullOrWhiteSpace(optimizedText) Then
                target.TextRange.Text = optimizedText
                If target.Shape IsNot Nothing Then AutoFitPptTextShape(target.Shape)
                changedCount += 1
            End If
        Next

        ShareRibbon.GlobalStatusStripAll.ShowProgress($"AI创作完成，共处理 {changedCount} 处")
        Return changedCount
    End Function

    Private Shared Function GetDocmeeTextOptimizeType(modeName As String) As String
        Select Case modeName
            Case "润色"
                Return "rs"
            Case "扩写"
                Return "kx"
            Case "缩写"
                Return "sx"
            Case "翻译"
                Return "fy"
            Case Else
                Return "rs"
        End Select
    End Function

    Private Shared Function GetDocmeeLanguageCode(languageName As String) As String
        Select Case languageName
            Case "中文"
                Return "zh"
            Case "英语"
                Return "en"
            Case "日语"
                Return "ja"
            Case "韩语"
                Return "ko"
            Case "法语"
                Return "fr"
            Case "德语"
                Return "de"
            Case "西班牙语"
                Return "es"
            Case "葡萄牙语"
                Return "pt"
            Case "俄语"
                Return "ru"
            Case "意大利语"
                Return "it"
            Case Else
                Return ""
        End Select
    End Function

    Private Shared Sub PostToUi(context As SynchronizationContext, action As Action)
        If action Is Nothing Then Return
        If context Is Nothing Then
            action.Invoke()
        Else
            context.Post(Sub(state) action.Invoke(), Nothing)
        End If
    End Sub

    ' 文本翻译入口改为选中文本 AI 创作，使用 Docmee 流式 textOptimize 接口。
    Protected Overrides Async Sub TranslateButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            ShareRibbon.GlobalStatusStripAll.ShowProgress("AI内容提效：正在启动 AI创作...")
            Dim creationOptions = ShowAiCreationDialog()
            If creationOptions Is Nothing OrElse String.IsNullOrWhiteSpace(creationOptions.ModeName) Then Return

            Dim targets = GetSelectedPptTextTargets("AI创作")
            If targets.Count = 0 Then
                MessageBox.Show("请先选中 PPT 里的文字或文本框。", "AI创作", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim changedCount = Await OptimizeSelectedTextWithDocmeeAsync(targets, creationOptions)
            ' 结果已通过状态栏提示，无需弹窗打断用户操作
        Catch ex As Exception
            MessageBox.Show("AI创作出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Protected Overrides Sub ContinuationButton_Click(sender As Object, e As RibbonControlEventArgs)
    End Sub

    ' 主题生成PPT - 打开 Docmee V2 演示任务窗格
    Protected Overrides Sub TemplateFormatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Globals.ThisAddIn.ShowThemePptTaskPane()
        Catch ex As Exception
            MessageBox.Show("打开主题生成PPT面板出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 提取PPT演示文稿的完整结构为JSON格式
    ''' </summary>
    Private Function ExtractPresentationStructure(pres As Microsoft.Office.Interop.PowerPoint.Presentation, templateName As String) As JObject
        Dim result As New JObject()
        result("templateName") = templateName
        result("totalSlides") = pres.Slides.Count
        result("slideWidth") = pres.PageSetup.SlideWidth
        result("slideHeight") = pres.PageSetup.SlideHeight

        ' 幻灯片数组
        Dim slides As New JArray()

        ' 遍历幻灯片（最多30张）
        For i = 1 To Math.Min(pres.Slides.Count, 30)
            Dim slide = pres.Slides(i)
            Dim slideObj As New JObject()
            slideObj("slideIndex") = i
            slideObj("slideLayout") = GetLayoutName(slide.Layout)

            ' 元素数组：包含文本框、图片、表格等
            Dim elements As New JArray()
            Dim elementIndex As Integer = 0

            ' 遍历幻灯片中的形状
            For Each shape As Microsoft.Office.Interop.PowerPoint.Shape In slide.Shapes
                Dim elemObj As New JObject()
                elemObj("index") = elementIndex

                ' 判断形状类型
                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ' 文本框/占位符
                    Dim text = shape.TextFrame.TextRange.Text.Trim()
                    elemObj("type") = "textbox"
                    elemObj("text") = text
                    elemObj("placeholderType") = GetPlaceholderTypeName(shape)

                    ' 提取文本格式
                    Dim formatting As New JObject()
                    Try
                        Dim textRange = shape.TextFrame.TextRange
                        formatting("fontName") = If(textRange.Font.Name, "")
                        formatting("fontSize") = If(textRange.Font.Size > 0, CDec(textRange.Font.Size), 18)
                        formatting("bold") = (textRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue)
                        formatting("italic") = (textRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue)
                        formatting("underline") = (textRange.Font.Underline = Microsoft.Office.Core.MsoTriState.msoTrue)

                        ' 颜色
                        Try
                            Dim rgb = textRange.Font.Color.RGB
                            formatting("color") = $"#{rgb And &HFF:X2}{(rgb >> 8) And &HFF:X2}{(rgb >> 16) And &HFF:X2}"
                        Catch
                            formatting("color") = "auto"
                        End Try

                        ' 对齐方式
                        formatting("alignment") = GetPPTAlignmentString(textRange.ParagraphFormat.Alignment)
                    Catch ex As Exception
                        Debug.WriteLine($"提取PPT文本格式时出错: {ex.Message}")
                    End Try
                    elemObj("formatting") = formatting

                    ' 位置和大小
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)

                ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoPicture OrElse
                       shape.Type = Microsoft.Office.Core.MsoShapeType.msoLinkedPicture Then
                    ' 图片
                    elemObj("type") = "image"
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)

                ElseIf shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ' 表格
                    elemObj("type") = "table"
                    elemObj("rows") = shape.Table.Rows.Count
                    elemObj("columns") = shape.Table.Columns.Count
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)

                    ' 提取表格首行作为示例
                    Dim headerCells As New JArray()
                    Try
                        For c = 1 To shape.Table.Columns.Count
                            Dim cellText = shape.Table.Cell(1, c).Shape.TextFrame.TextRange.Text.Trim()
                            headerCells.Add(cellText)
                        Next
                        elemObj("headerCells") = headerCells
                    Catch
                        ' 忽略合并单元格等情况
                    End Try

                ElseIf shape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ' 图表
                    elemObj("type") = "chart"
                    elemObj("chartType") = shape.Chart.ChartType.ToString()
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)

                Else
                    ' 其他形状
                    elemObj("type") = "shape"
                    elemObj("shapeType") = shape.Type.ToString()
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)
                End If

                elements.Add(elemObj)
                elementIndex += 1
            Next

            slideObj("elements") = elements
            slides.Add(slideObj)
        Next

        result("slides") = slides
        Return result
    End Function

    Private Function GetLayoutName(layout As Microsoft.Office.Interop.PowerPoint.PpSlideLayout) As String
        Select Case layout
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitle : Return "标题幻灯片"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitleOnly : Return "仅标题"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText : Return "标题和内容"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTwoColumnText : Return "两栏内容"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank : Return "空白"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutContentWithCaption : Return "内容与标题"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutPictureWithCaption : Return "图片与标题"
            Case Else : Return "自定义"
        End Select
    End Function

    Private Function GetPPTAlignmentString(alignment As Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment) As String
        Select Case alignment
            Case Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft : Return "left"
            Case Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter : Return "center"
            Case Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight : Return "right"
            Case Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignJustify : Return "justify"
            Case Else : Return "left"
        End Select
    End Function

    Private Function GetPlaceholderTypeName(shape As Microsoft.Office.Interop.PowerPoint.Shape) As String
        Try
            If shape.PlaceholderFormat Is Nothing Then Return "文本框"
            Select Case shape.PlaceholderFormat.Type
                Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle : Return "标题"
                Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle : Return "居中标题"
                Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle : Return "副标题"
                Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody : Return "正文"
                Case Else : Return "内容"
            End Select
        Catch
            Return "文本"
        End Try
    End Function

    Private Function EscapeForJs(text As String) As String
        Return text.Replace("\", "\\").Replace("`", "\`").Replace("$", "\$").Replace(vbCr, "").Replace(vbLf, "\n")
    End Function
End Class
