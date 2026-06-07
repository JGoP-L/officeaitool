Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Office.Core
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class ThemePptTaskPane
    Inherits UserControl

    Private Const ThemePptPaneBuild As String = "2026.06.04.9"
    Private Const MaxConcurrentTemplateCoverLoads As Integer = 1
    Private Const TemplatePageSize As Integer = 20
    Private Const TemplateCoverHostName As String = "theme-ppt-covers.local"
    Private Const DocmeePptxIdTagName As String = "wenduoduoAI_DocmeePptxId"
    Private Const WM_SETREDRAW As Integer = &HB
    Private Const EM_LINESCROLL As Integer = &HB6
    Private Const EM_GETFIRSTVISIBLELINE As Integer = &HCE
    Private Const GenerationModeTitle As String = "标题生成"
    Private Const GenerationModeDocument As String = "文档生成"
    Private Const GenerationModeMarkdown As String = "Markdown大纲"
    Private Shared ReadOnly MarkdownPreviewPipelineLock As New Object()
    Private Shared _markdownPreviewPipeline As MarkdownPipeline
    Private Shared _markdownPreviewPipelineInitializeError As String

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    Private ReadOnly _pptApp As PowerPoint.Application
    Private ReadOnly _client As New DocmeePptClient()
    Private _outline As JObject
    Private _outlineMarkdown As String
    Private _taskId As String
    Private _lastTemplateLoadUsedFallback As Boolean
    Private _isTemplateLoading As Boolean
    Private _templateLoadCts As CancellationTokenSource
    Private _templateCoverCts As CancellationTokenSource
    Private _templateCoverFailureCount As Integer
    Private _templatePage As Integer = 1
    Private _templateHasNextPage As Boolean
    Private _isOutlineEditCompleted As Boolean
    Private _templateConfirmedForCurrentOutline As Boolean
    Private _confirmedTemplateId As String
    Private _outlinePreviewReady As Boolean
    Private _outlinePreviewInitializing As Boolean
    Private _pendingOutlinePreviewRender As Boolean
    Private _suppressOutlineEditorChange As Boolean
    Private _applyingMarkdownEditorHighlight As Boolean
    Private _markdownPreviewRenderGeneration As Integer
    Private _selectedDocumentPath As String
    Private _lastGeneratedPptId As String
    Private _lastImportedSlideStartIndex As Integer
    Private _lastImportedSlideCount As Integer

    Private ReadOnly _generationModeCombo As New ComboBox()
    Private ReadOnly _topicBox As New TextBox()
    Private ReadOnly _documentPanel As New TableLayoutPanel()
    Private ReadOnly _documentPathBox As New TextBox()
    Private ReadOnly _chooseDocumentButton As New Button()
    Private ReadOnly _generateButton As New Button()
    Private ReadOnly _insertButton As New Button()
    Private ReadOnly _finishOutlineEditButton As New Button()
    Private ReadOnly _configureDocmeeButton As New Button()
    Private ReadOnly _templateCombo As New ComboBox()
    Private ReadOnly _refreshTemplatesButton As New Button()
    Private ReadOnly _selectTemplateButton As New Button()
    Private ReadOnly _templateSectionLabel As New Label()
    Private ReadOnly _templateSectionPanel As New TableLayoutPanel()
    Private ReadOnly _outputBox As New TextBox()
    Private ReadOnly _contentPanel As New Panel()
    Private ReadOnly _outlineWorkspacePanel As New SplitContainer()
    Private ReadOnly _outlineEditor As New RichTextBox()
    Private ReadOnly _outlinePreviewWebView As New WebView2()
    Private ReadOnly _outlinePreviewDebounceTimer As New System.Windows.Forms.Timer()
    Private ReadOnly _templateCardPanel As New FlowLayoutPanel()
    Private ReadOnly _templateListBox As New ListBox()
    Private ReadOnly _templateWebView As New WebView2()
    Private ReadOnly _templatePaintGallery As New TemplateGalleryPaintControl()
    Private ReadOnly _templateCards As New Dictionary(Of String, Panel)()
    Private ReadOnly _templateSelectLabels As New Dictionary(Of String, Label)()
    Private ReadOnly _templateCoverBoxes As New Dictionary(Of String, PictureBox)()
    Private ReadOnly _templateCoverHosts As New Dictionary(Of String, Panel)()
    Private ReadOnly _templateCoverStatusLabels As New Dictionary(Of String, Label)()
    Private ReadOnly _templatePreviewImages As New Dictionary(Of String, System.Drawing.Image)()
    Private ReadOnly _templateCoverImageUrls As New Dictionary(Of String, String)()
    Private ReadOnly _templateCoverFilePaths As New Dictionary(Of String, String)()
    Private ReadOnly _templateCoverMessages As New Dictionary(Of String, String)()
    Private ReadOnly _statusLabel As New Label()
    Private _templateCoverLoadGeneration As Integer
    Private _templateWebViewReady As Boolean
    Private _templateWebViewInitializing As Boolean
    Private _pendingTemplateGalleryRender As Boolean
    Private ReadOnly _uiThreadId As Integer

    Public Sub New(pptApp As PowerPoint.Application)
        _uiThreadId = Thread.CurrentThread.ManagedThreadId
        _pptApp = pptApp
        AppendThemePptLog("Pane constructing. Build=" & ThemePptPaneBuild)
        BuildLayout()
        AddHandler Me.Load, AddressOf ThemePptTaskPane_Load
        AppendThemePptLog("Pane constructed.")
    End Sub

    Private Sub BuildLayout()
        Me.BackColor = OfficeAIStyleHelper.BgPage
        Me.Padding = New Padding(0)

        ' 顶部品牌色标题栏
        Dim headerPanel As New Panel()
        headerPanel.Dock = DockStyle.Top
        headerPanel.Height = 48
        headerPanel.BackColor = OfficeAIStyleHelper.BrandPrimary
        headerPanel.Padding = New Padding(OfficeAIStyleHelper.SpacingLg, 0, OfficeAIStyleHelper.SpacingLg, 0)

        Dim headerTitle As New Label()
        headerTitle.Text = "AI 生成 PPT"
        headerTitle.Font = New Font("Microsoft YaHei UI", 12.0!, FontStyle.Bold)
        headerTitle.ForeColor = Color.White
        headerTitle.AutoSize = True
        headerTitle.Location = New Point(OfficeAIStyleHelper.SpacingLg, 11)
        headerPanel.Controls.Add(headerTitle)

        Dim headerVersion As New Label()
        headerVersion.Text = "v" & ThemePptPaneBuild
        headerVersion.Font = New Font("Microsoft YaHei UI", 8.0!, FontStyle.Regular)
        headerVersion.ForeColor = Color.FromArgb(180, 180, 255)
        headerVersion.AutoSize = True
        headerVersion.Location = New Point(headerTitle.Right + 8, 18)
        headerPanel.Controls.Add(headerVersion)

        Me.Controls.Add(headerPanel)

        ' 主体滚动区域
        Dim scrollPanel As New Panel()
        scrollPanel.Dock = DockStyle.Fill
        scrollPanel.AutoScroll = True
        scrollPanel.Padding = New Padding(OfficeAIStyleHelper.SpacingLg)
        scrollPanel.BackColor = OfficeAIStyleHelper.BgPage

        Dim layout As New TableLayoutPanel()
        layout.Dock = DockStyle.Fill
        layout.AutoSize = False
        layout.ColumnCount = 1
        layout.RowCount = 10
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 200.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.BackColor = Color.Transparent

        ' --- 生成方式选择器（Segmented Control 风格）---
        Dim modeLabel As New Label()
        OfficeAIStyleHelper.StyleLabelHeading(modeLabel)
        modeLabel.Text = "生成方式"
        modeLabel.Margin = New Padding(0, 0, 0, OfficeAIStyleHelper.SpacingSm)

        Dim modeSegmentedPanel As New Panel()
        modeSegmentedPanel.Height = 36
        modeSegmentedPanel.Width = 400
        modeSegmentedPanel.BackColor = OfficeAIStyleHelper.BorderLight
        modeSegmentedPanel.Padding = New Padding(1)
        modeSegmentedPanel.Margin = New Padding(0, 0, 0, OfficeAIStyleHelper.SpacingMd)

        ' 三个模式按钮作为 Segmented Control
        Dim modes() As String = {GenerationModeTitle, GenerationModeDocument, GenerationModeMarkdown}
        Dim modeBtnWidth As Integer = 98
        For i As Integer = 0 To modes.Length - 1
            Dim modeBtn As New Button()
            modeBtn.Text = modes(i)
            modeBtn.FlatStyle = FlatStyle.Flat
            modeBtn.FlatAppearance.BorderSize = 0
            modeBtn.FlatAppearance.MouseOverBackColor = Color.Transparent
            modeBtn.FlatAppearance.MouseDownBackColor = Color.Transparent
            modeBtn.UseVisualStyleBackColor = False
            modeBtn.Height = 34
            modeBtn.Width = modeBtnWidth
            modeBtn.Left = 1 + i * (modeBtnWidth + 1)
            modeBtn.Top = 1
            modeBtn.Font = OfficeAIStyleHelper.FontUi
            modeBtn.Cursor = Cursors.Hand
            modeBtn.TextAlign = ContentAlignment.MiddleCenter
            modeBtn.Tag = i
            Dim captured As Button = modeBtn
            AddHandler modeBtn.Click, Sub(s, e)
                                          _generationModeCombo.SelectedIndex = CInt(captured.Tag)
                                      End Sub
            modeSegmentedPanel.Controls.Add(modeBtn)
        Next
        ' 事件只绑定一次
        AddHandler _generationModeCombo.SelectedIndexChanged, Sub(s, e)
                                                                 UpdateModeButtonsStyle(modeSegmentedPanel, _generationModeCombo.SelectedIndex)
                                                             End Sub
        UpdateModeButtonsStyle(modeSegmentedPanel, 0)

        ' --- 主题输入框 ---
        _topicBox.Dock = DockStyle.Fill
        _topicBox.Multiline = True
        _topicBox.ScrollBars = ScrollBars.Vertical
        _topicBox.Text = "AI 办公趋势"
        _topicBox.Margin = New Padding(0, 0, 0, OfficeAIStyleHelper.SpacingMd)
        OfficeAIStyleHelper.StyleTextBoxMultiline(_topicBox)

        ' --- 文档选择面板 ---
        _documentPanel.Dock = DockStyle.Fill
        _documentPanel.ColumnCount = 3
        _documentPanel.RowCount = 1
        _documentPanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        _documentPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        _documentPanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        _documentPanel.Margin = New Padding(0, 0, 0, OfficeAIStyleHelper.SpacingMd)
        _documentPanel.Visible = False
        _documentPanel.BackColor = Color.Transparent

        Dim documentLabel As New Label()
        OfficeAIStyleHelper.StyleLabelBody(documentLabel)
        documentLabel.Text = "文档"
        documentLabel.Margin = New Padding(0, 6, 10, 0)

        _documentPathBox.Dock = DockStyle.Fill
        _documentPathBox.ReadOnly = True
        OfficeAIStyleHelper.StyleTextBox(_documentPathBox)

        _chooseDocumentButton.Text = "选择"
        _chooseDocumentButton.Width = 66
        AddHandler _chooseDocumentButton.Click, AddressOf ChooseDocumentButton_Click
        OfficeAIStyleHelper.StyleButtonSmall(_chooseDocumentButton)

        _documentPanel.Controls.Add(documentLabel, 0, 0)
        _documentPanel.Controls.Add(_documentPathBox, 1, 0)
        _documentPanel.Controls.Add(_chooseDocumentButton, 2, 0)

        ' --- 按钮组 ---
        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.AutoSize = True
        buttonPanel.Dock = DockStyle.Fill
        buttonPanel.FlowDirection = FlowDirection.LeftToRight
        buttonPanel.WrapContents = True
        buttonPanel.Margin = New Padding(0, 0, 0, OfficeAIStyleHelper.SpacingMd)
        OfficeAIStyleHelper.StyleFlowPanel(buttonPanel)

        _generateButton.Text = "生成大纲"
        _generateButton.Width = 88
        _generateButton.Height = OfficeAIStyleHelper.ButtonHeight
        AddHandler _generateButton.Click, AddressOf GenerateButton_Click
        OfficeAIStyleHelper.StyleButtonPrimary(_generateButton)

        _finishOutlineEditButton.Text = "完成编辑"
        _finishOutlineEditButton.Width = 88
        _finishOutlineEditButton.Height = OfficeAIStyleHelper.ButtonHeight
        _finishOutlineEditButton.Enabled = False
        AddHandler _finishOutlineEditButton.Click, AddressOf FinishOutlineEditButton_Click
        OfficeAIStyleHelper.StyleButtonSecondary(_finishOutlineEditButton)

        _insertButton.Text = "生成并导入"
        _insertButton.Width = 104
        _insertButton.Height = OfficeAIStyleHelper.ButtonHeight
        _insertButton.Enabled = False
        AddHandler _insertButton.Click, AddressOf InsertButton_Click
        OfficeAIStyleHelper.StyleButtonAccent(_insertButton)

        _configureDocmeeButton.Text = "配置"
        _configureDocmeeButton.Width = 66
        _configureDocmeeButton.Height = OfficeAIStyleHelper.ButtonHeight
        AddHandler _configureDocmeeButton.Click, AddressOf ConfigureDocmeeButton_Click
        OfficeAIStyleHelper.StyleButtonSecondary(_configureDocmeeButton)

        buttonPanel.Controls.Add(_generateButton)
        buttonPanel.Controls.Add(_finishOutlineEditButton)
        buttonPanel.Controls.Add(_insertButton)
        buttonPanel.Controls.Add(_configureDocmeeButton)

        ' --- 模板选择区 ---
        Dim templateSectionLabel = _templateSectionLabel
        OfficeAIStyleHelper.StyleLabelHeading(templateSectionLabel)
        templateSectionLabel.Text = "选择模板"
        templateSectionLabel.Margin = New Padding(0, OfficeAIStyleHelper.SpacingSm, 0, OfficeAIStyleHelper.SpacingSm)
        templateSectionLabel.Visible = False

        Dim templatePanel = _templateSectionPanel
        templatePanel.Dock = DockStyle.Fill
        templatePanel.ColumnCount = 3
        templatePanel.RowCount = 1
        templatePanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        templatePanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        templatePanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        templatePanel.Margin = New Padding(0, 0, 0, OfficeAIStyleHelper.SpacingSm)
        templatePanel.BackColor = Color.Transparent

        _templateCombo.Dock = DockStyle.Fill
        _templateCombo.DropDownStyle = ComboBoxStyle.DropDownList
        _templateCombo.Enabled = False
        OfficeAIStyleHelper.StyleComboBox(_templateCombo)
        AddHandler _templateCombo.SelectedIndexChanged, AddressOf TemplateCombo_SelectedIndexChanged

        _refreshTemplatesButton.Text = "刷新"
        _refreshTemplatesButton.Width = 66
        _refreshTemplatesButton.Enabled = False
        AddHandler _refreshTemplatesButton.Click, AddressOf RefreshTemplatesButton_Click
        OfficeAIStyleHelper.StyleButtonSmall(_refreshTemplatesButton)

        _selectTemplateButton.Text = "预览模板"
        _selectTemplateButton.Width = 96
        _selectTemplateButton.Enabled = False
        AddHandler _selectTemplateButton.Click, AddressOf SelectTemplateButton_Click
        OfficeAIStyleHelper.StyleButtonSecondary(_selectTemplateButton)

        templatePanel.Controls.Add(_templateCombo, 0, 0)
        templatePanel.Controls.Add(_refreshTemplatesButton, 1, 0)
        templatePanel.Controls.Add(_selectTemplateButton, 2, 0)
        templatePanel.Visible = False

        ' --- 状态栏 ---
        _statusLabel.AutoSize = False
        _statusLabel.Height = 24
        _statusLabel.ForeColor = OfficeAIStyleHelper.TextSecondary
        _statusLabel.Font = OfficeAIStyleHelper.FontUiSmall
        _statusLabel.Text = "选择来源生成 Markdown 大纲，编辑完成后再选择模板生成 PPT。"
        _statusLabel.Margin = New Padding(0, 0, 0, OfficeAIStyleHelper.SpacingSm)
        _statusLabel.TextAlign = ContentAlignment.MiddleLeft

        ' --- 内容区域（大纲编辑/预览/模板卡片/输出）---
        _outlineWorkspacePanel.Dock = DockStyle.Fill
        _outlineWorkspacePanel.Orientation = Orientation.Horizontal
        _outlineWorkspacePanel.SplitterWidth = 6
        _outlineWorkspacePanel.Panel1MinSize = 180
        _outlineWorkspacePanel.Panel2MinSize = 100
        _outlineWorkspacePanel.Visible = False
        _outlineWorkspacePanel.BackColor = OfficeAIStyleHelper.BgPage

        Dim outlineEditorPanel As New TableLayoutPanel()
        outlineEditorPanel.Dock = DockStyle.Fill
        outlineEditorPanel.ColumnCount = 1
        outlineEditorPanel.RowCount = 2
        outlineEditorPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        outlineEditorPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        outlineEditorPanel.BackColor = Color.Transparent

        Dim outlineEditorLabel As New Label()
        OfficeAIStyleHelper.StyleLabelHeading(outlineEditorLabel)
        outlineEditorLabel.Text = "Markdown 源码"
        outlineEditorLabel.Margin = New Padding(0, 0, 0, 4)

        _outlineEditor.Dock = DockStyle.Fill
        _outlineEditor.Multiline = True
        _outlineEditor.ReadOnly = False
        _outlineEditor.ScrollBars = RichTextBoxScrollBars.Vertical
        _outlineEditor.WordWrap = True
        _outlineEditor.AcceptsTab = True
        OfficeAIStyleHelper.StyleRichTextBox(_outlineEditor)
        AddHandler _outlineEditor.TextChanged, AddressOf OutlineEditor_TextChanged

        outlineEditorPanel.Controls.Add(outlineEditorLabel, 0, 0)
        outlineEditorPanel.Controls.Add(_outlineEditor, 0, 1)

        Dim outlinePreviewPanel As New TableLayoutPanel()
        outlinePreviewPanel.Dock = DockStyle.Fill
        outlinePreviewPanel.ColumnCount = 1
        outlinePreviewPanel.RowCount = 2
        outlinePreviewPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        outlinePreviewPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        outlinePreviewPanel.BackColor = Color.Transparent

        Dim outlinePreviewLabel As New Label()
        OfficeAIStyleHelper.StyleLabelHeading(outlinePreviewLabel)
        outlinePreviewLabel.Text = "Markdown 预览"
        outlinePreviewLabel.Margin = New Padding(0, 0, 0, 4)

        _outlinePreviewWebView.Dock = DockStyle.Fill
        _outlinePreviewWebView.DefaultBackgroundColor = OfficeAIStyleHelper.BgSurface

        _outlinePreviewDebounceTimer.Interval = 350
        AddHandler _outlinePreviewDebounceTimer.Tick, AddressOf OutlinePreviewDebounceTimer_Tick

        outlinePreviewPanel.Controls.Add(outlinePreviewLabel, 0, 0)
        outlinePreviewPanel.Controls.Add(_outlinePreviewWebView, 0, 1)

        _outlineWorkspacePanel.Panel1.Controls.Add(outlineEditorPanel)
        _outlineWorkspacePanel.Panel2.Controls.Add(outlinePreviewPanel)

        _outputBox.Dock = DockStyle.Fill
        _outputBox.Multiline = True
        _outputBox.ReadOnly = True
        _outputBox.ScrollBars = ScrollBars.Vertical
        _outputBox.Visible = True
        OfficeAIStyleHelper.StyleTextBox(_outputBox)
        _outputBox.BackColor = OfficeAIStyleHelper.BgSurface

        _contentPanel.Dock = DockStyle.Fill
        _contentPanel.Margin = New Padding(0, 0, 0, 0)

        _templateCardPanel.Dock = DockStyle.Fill
        _templateCardPanel.AutoScroll = True
        _templateCardPanel.FlowDirection = FlowDirection.TopDown
        _templateCardPanel.WrapContents = False
        _templateCardPanel.BackColor = OfficeAIStyleHelper.BgSurface
        _templateCardPanel.Visible = False
        AddHandler _templateCardPanel.Resize, AddressOf TemplateCardPanel_Resize

        _templateListBox.Dock = DockStyle.Fill
        _templateListBox.DrawMode = DrawMode.OwnerDrawFixed
        _templateListBox.ItemHeight = 182
        _templateListBox.IntegralHeight = False
        _templateListBox.BorderStyle = BorderStyle.FixedSingle
        _templateListBox.BackColor = OfficeAIStyleHelper.BgSurface
        _templateListBox.Visible = False
        AddHandler _templateListBox.DrawItem, AddressOf TemplateListBox_DrawItem
        AddHandler _templateListBox.SelectedIndexChanged, AddressOf TemplateListBox_SelectedIndexChanged

        _templateWebView.Dock = DockStyle.Fill
        _templateWebView.Visible = False
        _templateWebView.DefaultBackgroundColor = OfficeAIStyleHelper.BgSurface

        _templatePaintGallery.Dock = DockStyle.Fill
        _templatePaintGallery.Visible = False
        _templatePaintGallery.BackColor = OfficeAIStyleHelper.BgSurface
        AddHandler _templatePaintGallery.TemplateSelected, AddressOf TemplatePaintGallery_TemplateSelected

        _contentPanel.Controls.Add(_templateListBox)
        _contentPanel.Controls.Add(_templateCardPanel)
        _contentPanel.Controls.Add(_outlineWorkspacePanel)
        _contentPanel.Controls.Add(_outputBox)

        ' 底部提示
        Dim hintLabel As New Label()
        hintLabel.AutoSize = False
        hintLabel.Dock = DockStyle.Fill
        hintLabel.Height = 36
        OfficeAIStyleHelper.StyleLabelHint(hintLabel)
        hintLabel.TextAlign = ContentAlignment.MiddleLeft
        hintLabel.Text = "版本 " & ThemePptPaneBuild & " | Docmee 地址和 token 可配置"

        ' --- 组装布局 ---
        layout.Controls.Add(modeLabel, 0, 0)
        layout.Controls.Add(modeSegmentedPanel, 0, 1)
        layout.Controls.Add(_topicBox, 0, 2)
        layout.Controls.Add(_documentPanel, 0, 3)
        layout.Controls.Add(buttonPanel, 0, 4)
        layout.Controls.Add(templateSectionLabel, 0, 5)
        layout.Controls.Add(templatePanel, 0, 6)
        layout.Controls.Add(_statusLabel, 0, 7)
        layout.Controls.Add(_contentPanel, 0, 8)
        layout.Controls.Add(hintLabel, 0, 9)

        scrollPanel.Controls.Add(layout)
        Me.Controls.Add(scrollPanel)

        ' 初始化隐藏的 ComboBox（作为模式数据源）
        _generationModeCombo.Visible = False
        _generationModeCombo.DropDownStyle = ComboBoxStyle.DropDownList
        _generationModeCombo.Items.AddRange(New Object() {GenerationModeTitle, GenerationModeDocument, GenerationModeMarkdown})
        _generationModeCombo.SelectedIndex = 0
        AddHandler _generationModeCombo.SelectedIndexChanged, AddressOf GenerationModeCombo_SelectedIndexChanged
        Me.Controls.Add(_generationModeCombo)

        UpdateGenerationModeUi()
    End Sub

    ''' <summary>更新 Segmented Control 按钮样式</summary>
    Private Sub UpdateModeButtonsStyle(container As Panel, selectedIndex As Integer)
        For Each ctrl As Control In container.Controls
            Dim btn = TryCast(ctrl, Button)
            If btn Is Nothing Then Continue For
            Dim idx = CInt(btn.Tag)
            If idx = selectedIndex Then
                btn.BackColor = OfficeAIStyleHelper.BgSurface
                btn.ForeColor = OfficeAIStyleHelper.BrandPrimary
                btn.Font = OfficeAIStyleHelper.FontUiBold
            Else
                btn.BackColor = Color.Transparent
                btn.ForeColor = OfficeAIStyleHelper.TextSecondary
                btn.Font = OfficeAIStyleHelper.FontUi
            End If
        Next
    End Sub

    ''' <summary>控制主题输入框可见性并折叠所在行</summary>
    Private Sub SetTopicBoxVisible(visible As Boolean)
        _topicBox.Visible = visible
        Dim layout = TryCast(_topicBox.Parent, TableLayoutPanel)
        If layout Is Nothing OrElse layout.RowCount <= 2 Then Return
        layout.RowStyles(2).Height = If(visible, 200.0F, 0)
    End Sub

    ''' <summary>控制模板选择区域可见性</summary>
    Private Sub SetTemplateSectionVisible(visible As Boolean)
        _templateSectionLabel.Visible = visible
        _templateSectionPanel.Visible = visible
    End Sub

    Private Async Sub ThemePptTaskPane_Load(sender As Object, e As EventArgs)
        AppendThemePptLog("Pane load: WebView2 lazy initialization requested.")
        Await InitializeOutlinePreviewWebViewAsync()
    End Sub

    Private Sub GenerationModeCombo_SelectedIndexChanged(sender As Object, e As EventArgs)
        UpdateGenerationModeUi()
    End Sub

    Private Sub UpdateGenerationModeUi()
        Dim mode = GetSelectedGenerationMode()
        _documentPanel.Visible = String.Equals(mode, GenerationModeDocument, StringComparison.Ordinal)

        Select Case mode
            Case GenerationModeDocument
                If String.Equals(_topicBox.Text.Trim(), "AI 办公趋势", StringComparison.Ordinal) Then
                    _topicBox.Text = "可选：补充文档生成要求，例如突出汇报重点。"
                End If
                SetStatus("选择 Word 或其他文档，生成 Markdown 大纲后可编辑并选择模板。")
            Case GenerationModeMarkdown
                If String.Equals(_topicBox.Text.Trim(), "AI 办公趋势", StringComparison.Ordinal) OrElse
                   _topicBox.Text.Trim().StartsWith("可选：", StringComparison.Ordinal) Then
                    _topicBox.Text = "# 演示主题" & vbCrLf & vbCrLf & "## 章节一" & vbCrLf & "### 页面标题" & vbCrLf & "- 要点内容"
                End If
                SetStatus("粘贴 Markdown 大纲，确认后再选择模板生成 PPT。")
            Case Else
                If _topicBox.Text.Trim().StartsWith("可选：", StringComparison.Ordinal) Then
                    _topicBox.Text = "AI 办公趋势"
                End If
                SetStatus("输入主题生成 Markdown 大纲，编辑完成后再选择模板生成 PPT。")
        End Select
    End Sub

    Private Function GetSelectedGenerationMode() As String
        Dim selectedMode = TryCast(_generationModeCombo.SelectedItem, String)
        If String.IsNullOrWhiteSpace(selectedMode) Then Return GenerationModeTitle
        Return selectedMode
    End Function

    Private Sub ChooseDocumentButton_Click(sender As Object, e As EventArgs)
        Using dialog As New OpenFileDialog()
            dialog.Title = "选择用于生成 PPT 的文档"
            dialog.Filter = "Docmee 支持的文档|*.doc;*.docx;*.pdf;*.ppt;*.pptx;*.txt;*.md;*.xls;*.xlsx;*.csv;*.html;*.epub;*.mobi;*.xmind;*.mm|所有文件|*.*"
            dialog.Multiselect = False

            If dialog.ShowDialog(Me.FindForm()) <> DialogResult.OK Then Return

            _selectedDocumentPath = dialog.FileName
            _documentPathBox.Text = _selectedDocumentPath
            SetStatus("已选择文档：" & Path.GetFileName(_selectedDocumentPath))
        End Using
    End Sub

    Private Function GetSelectedDocumentPath() As String
        If Not String.IsNullOrWhiteSpace(_selectedDocumentPath) Then Return _selectedDocumentPath.Trim()
        Return If(_documentPathBox.Text, "").Trim()
    End Function

    Private Async Function InitializeOutlinePreviewWebViewAsync() As Task
        If _outlinePreviewReady OrElse _outlinePreviewInitializing Then Return

        _outlinePreviewInitializing = True
        Try
            AppendThemePptLog("Outline preview WebView2 initialize start.")
            WebView2Loader.EnsureWebView2Loader()
            Dim userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAiThemePptMarkdownPreviewWebView2")
            Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)
            Await _outlinePreviewWebView.EnsureCoreWebView2Async(env)

            If _outlinePreviewWebView.CoreWebView2 Is Nothing Then
                AppendThemePptLog("Outline preview WebView2 initialize failed: CoreWebView2 is null.")
                Return
            End If

            _outlinePreviewWebView.CoreWebView2.Settings.IsScriptEnabled = False
            _outlinePreviewWebView.CoreWebView2.Settings.IsWebMessageEnabled = False
            _outlinePreviewWebView.CoreWebView2.Settings.AreDevToolsEnabled = True
            _outlinePreviewReady = True
            AppendThemePptLog("Outline preview WebView2 initialize success.")

            If _pendingOutlinePreviewRender Then
                ScheduleMarkdownPreviewUpdate(True)
            Else
                _outlinePreviewWebView.NavigateToString(BuildMarkdownPreviewHtml(""))
            End If
        Catch ex As Exception
            AppendThemePptLog("Outline preview WebView2 initialize exception: " & ex.ToString())
            SetStatus("Markdown 预览初始化失败：" & ex.Message)
        Finally
            _outlinePreviewInitializing = False
        End Try
    End Function

    Private Async Function InitializeTemplateWebViewAsync() As Task
        If _templateWebViewReady OrElse _templateWebViewInitializing Then Return

        _templateWebViewInitializing = True
        Try
            AppendThemePptLog("WebView2 initialize start.")
            WebView2Loader.EnsureWebView2Loader()
            Dim userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAiThemePptWebView2")
            Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)
            Await _templateWebView.EnsureCoreWebView2Async(env)

            If _templateWebView.CoreWebView2 IsNot Nothing Then
                _templateWebView.CoreWebView2.Settings.IsScriptEnabled = True
                _templateWebView.CoreWebView2.Settings.IsWebMessageEnabled = True
                _templateWebView.CoreWebView2.Settings.AreDevToolsEnabled = True
                Directory.CreateDirectory(GetTemplateCoverCacheDirectory())
                _templateWebView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    TemplateCoverHostName,
                    GetTemplateCoverCacheDirectory(),
                    CoreWebView2HostResourceAccessKind.Allow)
                AddHandler _templateWebView.CoreWebView2.NavigationCompleted, AddressOf TemplateWebView_NavigationCompleted
                AddHandler _templateWebView.CoreWebView2.WebMessageReceived, AddressOf TemplateWebView_WebMessageReceived
                _templateWebViewReady = True
                AppendThemePptLog("WebView2 initialize success.")
                RenderTemplateGallery()
            Else
                AppendThemePptLog("WebView2 initialize failed: CoreWebView2 is null.")
                SetStatus("模板列表 WebView2 初始化失败：CoreWebView2 不可用。")
            End If
        Catch ex As Exception
            AppendThemePptLog("WebView2 initialize exception: " & ex.ToString())
            SetStatus("模板列表 WebView2 初始化失败：" & ex.Message)
        Finally
            _templateWebViewInitializing = False
        End Try
    End Function

    Private Sub ShowOutlineOutput()
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() ShowOutlineOutput(), MethodInvoker))
            Return
        End If

        _outlineWorkspacePanel.Visible = False
        _templateCardPanel.Visible = False
        _templateListBox.Visible = False
        _templateWebView.Visible = False
        _templatePaintGallery.Visible = False
        _outputBox.Visible = True
        _outputBox.BringToFront()
    End Sub

    Private Sub ShowOutlineEditor()
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() ShowOutlineEditor(), MethodInvoker))
            Return
        End If

        _templateCardPanel.Visible = False
        _templateListBox.Visible = False
        _templateWebView.Visible = False
        _templatePaintGallery.Visible = False
        _outputBox.Visible = False
        _outlineWorkspacePanel.Visible = True
        _outlineWorkspacePanel.BringToFront()
        BeginInvokeIfAlive(CType(Sub() ApplyOutlineWorkspaceLayout(), MethodInvoker))
    End Sub

    Private Sub ApplyOutlineWorkspaceLayout()
        If _outlineWorkspacePanel.IsDisposed OrElse Not _outlineWorkspacePanel.Visible Then Return

        Try
            Dim availableHeight = _outlineWorkspacePanel.Height - _outlineWorkspacePanel.SplitterWidth
            If availableHeight <= 0 Then Return

            Dim desiredEditorHeight = Math.Max(_outlineWorkspacePanel.Panel1MinSize, CInt(availableHeight * 0.62))
            Dim maxEditorHeight = Math.Max(_outlineWorkspacePanel.Panel1MinSize, availableHeight - _outlineWorkspacePanel.Panel2MinSize)
            desiredEditorHeight = Math.Min(desiredEditorHeight, maxEditorHeight)
            If desiredEditorHeight > 0 Then _outlineWorkspacePanel.SplitterDistance = desiredEditorHeight
        Catch ex As Exception
            AppendThemePptLog("Apply outline workspace layout failed: " & ex.ToString())
        End Try
    End Sub

    Private Sub ShowTemplateGallery()
        If _templateCombo.Items.Count = 0 Then
            ShowOutlineEditor()
            Return
        End If

        ShowOutlineEditor()
        AppendThemePptLog("ShowTemplateGallery skipped in task pane; use WebView2 dialog. count=" &
                          _templateCombo.Items.Count.ToString() & ", selected=" & GetSelectedTemplateId())
    End Sub

    Private Sub RenderTemplateGallery()
        ' The PowerPoint task pane is sensitive to hosted child HWNDs and direct GDI redraws.
        ' Keep template rendering on the lightweight ListBox path only.
    End Sub

    Private Sub UpdateTemplatePaintGallery()
        Try
            Dim templates = GetCurrentTemplatesSnapshot()
            _templatePaintGallery.SetData(templates, _templatePreviewImages, _templateCoverMessages, GetSelectedTemplateId())
            AppendThemePptLog("PaintGallery data updated: templates=" & templates.Count.ToString() &
                              ", images=" & _templatePreviewImages.Count.ToString() &
                              ", messages=" & _templateCoverMessages.Count.ToString() &
                              ", size=" & _templatePaintGallery.Width.ToString() & "x" & _templatePaintGallery.Height.ToString())
        Catch ex As Exception
            AppendThemePptLog("PaintGallery data update exception: " & ex.ToString())
            SetStatus("模板列表渲染失败：" & ex.Message)
        End Try
    End Sub

    Private Sub NavigateTemplateGalleryFile()
        Try
            Dim cacheDirectory = GetTemplateCoverCacheDirectory()
            Directory.CreateDirectory(cacheDirectory)
            Dim htmlPath = Path.Combine(cacheDirectory, "gallery.html")
            File.WriteAllText(htmlPath, BuildTemplateGalleryHtml(), Encoding.UTF8)
            AppendThemePptLog("WebView2 gallery HTML written: " & htmlPath)
            _templateWebView.CoreWebView2.Navigate("https://" & TemplateCoverHostName & "/gallery.html?t=" & DateTime.UtcNow.Ticks.ToString())
        Catch ex As Exception
            SetStatus("模板列表 HTML 写入失败：" & ex.Message)
        End Try
    End Sub

    Private Sub TemplateWebView_NavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        If Not e.IsSuccess Then
            SetStatus("模板列表渲染失败：" & e.WebErrorStatus.ToString())
            Return
        End If

        If _templateCombo.Items.Count > 0 AndAlso _templateCoverImageUrls.Count >= _templateCombo.Items.Count Then
            SetStatus("模板列表已渲染，封面加载完成。")
        End If
    End Sub

    Private Function BuildTemplateGalleryHtml() As String
        Dim selectedId = GetSelectedTemplateId()
        Dim builder As New StringBuilder()

        builder.AppendLine("<!doctype html>")
        builder.AppendLine("<html><head><meta charset=""utf-8""><style>")
        builder.AppendLine("html,body{margin:0;padding:0;background:#fff;font-family:'Microsoft YaHei UI','Microsoft YaHei',Arial,sans-serif;color:#272d37;font-size:12px;}")
        builder.AppendLine(".wrap{padding:0 4px 10px 0;box-sizing:border-box;}")
        builder.AppendLine(".card{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;background:#fff;margin:0 0 10px 0;padding:8px;cursor:pointer;}")
        builder.AppendLine(".card.selected{border:2px solid #ea580c;background:#fff7ed;padding:7px;}")
        builder.AppendLine(".cover{width:100%;aspect-ratio:16/9;background:#fff8f1;border:1px solid #e2e8f0;display:flex;align-items:center;justify-content:center;overflow:hidden;box-sizing:border-box;}")
        builder.AppendLine(".cover img{display:block;width:100%;height:100%;object-fit:contain;background:#fff;}")
        builder.AppendLine(".fallback{width:100%;height:100%;box-sizing:border-box;border-left:5px solid #ea580c;padding:18px 14px;display:flex;flex-direction:column;justify-content:space-between;}")
        builder.AppendLine(".fallback-title{font-weight:700;color:#272d37;line-height:1.5;word-break:break-word;}")
        builder.AppendLine(".fallback-status{color:#64748b;line-height:1.5;word-break:break-word;}")
        builder.AppendLine(".fallback-status.error{color:#b91c1c;}")
        builder.AppendLine(".title{margin-top:8px;font-size:13px;font-weight:700;line-height:1.45;word-break:break-word;}")
        builder.AppendLine(".meta{margin-top:6px;color:#64748b;line-height:1.4;word-break:break-word;}")
        builder.AppendLine(".status{margin-top:6px;color:#64748b;line-height:1.45;word-break:break-word;}")
        builder.AppendLine(".status.error{color:#b91c1c;}")
        builder.AppendLine(".btn{margin-top:8px;height:28px;min-width:96px;display:inline-flex;align-items:center;justify-content:center;background:#f1f5f9;color:#272d37;font-weight:700;}")
        builder.AppendLine(".selected .btn{background:#ea580c;color:#fff;}")
        builder.AppendLine("</style></head><body><div class=""wrap"">")

        For Each item In _templateCombo.Items
            Dim template = TryCast(item, DocmeeTemplateInfo)
            If template Is Nothing OrElse String.IsNullOrWhiteSpace(template.Id) Then Continue For

            Dim isSelected = String.Equals(template.Id, selectedId, StringComparison.Ordinal)
            Dim status = If(_templateCoverMessages.ContainsKey(template.Id), _templateCoverMessages(template.Id), "")
            Dim hasError = status.StartsWith("封面加载失败", StringComparison.Ordinal)
            Dim imageUrl = If(_templateCoverImageUrls.ContainsKey(template.Id), _templateCoverImageUrls(template.Id), "")
            Dim title = If(String.IsNullOrWhiteSpace(template.Name), template.Id, template.Name)

            builder.Append("<div class=""card")
            If isSelected Then builder.Append(" selected")
            builder.Append(""" data-id=""").Append(EscapeHtmlAttribute(template.Id)).Append(""">")
            builder.AppendLine("<div class=""cover"">")
            If Not String.IsNullOrWhiteSpace(imageUrl) Then
                builder.Append("<img src=""").Append(EscapeHtmlAttribute(imageUrl)).Append(""" alt=""").Append(EscapeHtmlAttribute(title)).Append(""">")
            Else
                builder.Append("<div class=""fallback""><div class=""fallback-title"">").Append(EscapeHtml(title)).Append("</div>")
                builder.Append("<div class=""fallback-status")
                If hasError Then builder.Append(" error")
                builder.Append(""">").Append(EscapeHtml(If(String.IsNullOrWhiteSpace(status), "封面加载中...", status))).Append("</div></div>")
            End If
            builder.AppendLine("</div>")
            builder.Append("<div class=""title"">").Append(EscapeHtml(title)).AppendLine("</div>")
            builder.Append("<div class=""meta"">").Append(EscapeHtml(BuildTemplateMetaText(template))).AppendLine("</div>")
            If Not String.IsNullOrWhiteSpace(status) AndAlso String.IsNullOrWhiteSpace(imageUrl) Then
                builder.Append("<div class=""status")
                If hasError Then builder.Append(" error")
                builder.Append(""">").Append(EscapeHtml(status)).AppendLine("</div>")
            End If
            builder.Append("<div class=""btn"">").Append(If(isSelected, "已选择", "选择模板")).AppendLine("</div>")
            builder.AppendLine("</div>")
        Next

        builder.AppendLine("</div><script>")
        builder.AppendLine("document.addEventListener('click',function(e){var card=e.target.closest('.card');if(!card)return;window.chrome.webview.postMessage({type:'selectTemplate',id:card.getAttribute('data-id')});});")
        builder.AppendLine("</script></body></html>")
        Return builder.ToString()
    End Function

    Private Sub TemplateWebView_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim payload = JObject.Parse(e.WebMessageAsJson)
            If Not String.Equals(TryGetString(payload("type")), "selectTemplate", StringComparison.Ordinal) Then Return

            Dim templateId = TryGetString(payload("id"))
            If String.IsNullOrWhiteSpace(templateId) Then Return

            Dim template = FindTemplateById(templateId)
            If template IsNot Nothing Then SelectTemplate(template)
        Catch ex As Exception
            SetStatus("模板选择失败：" & ex.Message)
        End Try
    End Sub

    Private Sub TemplatePaintGallery_TemplateSelected(sender As Object, template As DocmeeTemplateInfo)
        AppendThemePptLog("PaintGallery template selected: " & If(template Is Nothing, "", template.Id))
        SelectTemplate(template)
    End Sub

    Private Function GetSelectedTemplateId() As String
        Dim selectedTemplate = TryCast(_templateCombo.SelectedItem, DocmeeTemplateInfo)
        If selectedTemplate Is Nothing Then Return ""
        Return If(selectedTemplate.Id, "")
    End Function

    Private Function FindTemplateById(templateId As String) As DocmeeTemplateInfo
        If String.IsNullOrWhiteSpace(templateId) Then Return Nothing

        For Each item In _templateCombo.Items
            Dim template = TryCast(item, DocmeeTemplateInfo)
            If template IsNot Nothing AndAlso String.Equals(template.Id, templateId, StringComparison.Ordinal) Then
                Return template
            End If
        Next

        Return Nothing
    End Function

    Private Sub ShowTemplateSelectionDialog()
        Dim templates = GetCurrentTemplatesSnapshot()
        If templates.Count = 0 Then
            SetStatus("没有可预览的模板。")
            Return
        End If

        AppendThemePptLog("Template dialog opening. count=" & templates.Count.ToString() &
                          ", selected=" & GetSelectedTemplateId())

        Using dialog As New TemplateSelectionForm(templates,
                                                  GetSelectedTemplateId(),
                                                  AddressOf BuildTemplateCoverUrl,
                                                  _templatePage,
                                                  TemplatePageSize,
                                                  AddressOf LoadTemplatePageForDialog)
            Dim owner = Me.FindForm()
            Dim result As DialogResult
            If owner IsNot Nothing Then
                result = dialog.ShowDialog(owner)
            Else
                result = dialog.ShowDialog()
            End If

            If result = DialogResult.OK AndAlso dialog.SelectedTemplate IsNot Nothing Then
                If dialog.CurrentPage <> _templatePage Then
                    _templatePage = dialog.CurrentPage
                    _templateHasNextPage = dialog.HasNextPage
                    PopulateTemplates(dialog.CurrentTemplates)
                End If

                SelectTemplate(dialog.SelectedTemplate)
                _templateConfirmedForCurrentOutline = True
                _confirmedTemplateId = dialog.SelectedTemplate.Id
                RefreshActionButtons()

                Dim displayName = If(String.IsNullOrWhiteSpace(dialog.SelectedTemplate.Name),
                                     dialog.SelectedTemplate.Id,
                                     dialog.SelectedTemplate.Name)
                SetStatus("已选择模板：" & displayName)
            End If
        End Using
    End Sub

    Private Function LoadTemplatePageForDialog(page As Integer) As List(Of DocmeeTemplateInfo)
        Dim safePage = Math.Max(1, page)
        AppendThemePptLog("Template dialog page load requested: page=" & safePage.ToString())
        Return LoadTemplatesInBackgroundAsync(safePage, CancellationToken.None).GetAwaiter().GetResult()
    End Function

    Private Function CanChooseTemplate() As Boolean
        Return Not _isTemplateLoading AndAlso
               _templateCombo.Items.Count > 0 AndAlso
               Not String.IsNullOrWhiteSpace(GetEditedMarkdown())
    End Function

    Private Function CanGenerateFromTemplate() As Boolean
        Dim selectedTemplate = TryCast(_templateCombo.SelectedItem, DocmeeTemplateInfo)
        Return CanChooseTemplate() AndAlso
               selectedTemplate IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(selectedTemplate.Id)
    End Function

    Private Sub RefreshActionButtons()
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() RefreshActionButtons(), MethodInvoker))
            Return
        End If

        _finishOutlineEditButton.Enabled = Not String.IsNullOrWhiteSpace(GetEditedMarkdown()) AndAlso Not _isOutlineEditCompleted
        _refreshTemplatesButton.Enabled = _isOutlineEditCompleted AndAlso Not _isTemplateLoading
        _selectTemplateButton.Enabled = CanChooseTemplate()
        _insertButton.Enabled = CanGenerateFromTemplate()
    End Sub

    Private Function GetEditedMarkdown() As String
        Return If(_outlineEditor.Text, "").Trim()
    End Function

    Private Sub SetOutlineEditorText(markdown As String)
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() SetOutlineEditorText(markdown), MethodInvoker))
            Return
        End If

        _suppressOutlineEditorChange = True
        Try
            _outlineEditor.Text = NormalizeMarkdownForEditing(markdown)
            ApplyMarkdownEditorHighlight()
            _outlineEditor.SelectionStart = 0
            _outlineEditor.ScrollToCaret()
        Finally
            _suppressOutlineEditorChange = False
        End Try

        ScheduleMarkdownPreviewUpdate(True)
    End Sub

    Private Function NormalizeMarkdownForEditing(markdown As String) As String
        Dim value = If(markdown, "").Trim()
        If String.IsNullOrWhiteSpace(value) Then Return ""

        value = value.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        value = value.Replace("\r\n", vbLf).Replace("\n", vbLf).Replace("\r", vbLf)
        value = value.Replace(ChrW(&HA0), " ")

        value = System.Text.RegularExpressions.Regex.Replace(value, "[ \t]+\n", vbLf)
        value = System.Text.RegularExpressions.Regex.Replace(value, "([^\n])\s+(#{1,6}\s+)", "$1" & vbLf & vbLf & "$2")
        value = System.Text.RegularExpressions.Regex.Replace(value, "([^\n])\s+([-*+]\s+)", "$1" & vbLf & "$2")
        value = System.Text.RegularExpressions.Regex.Replace(value, "([^\n])\s+(\d{1,2}[\.\)]\s+)", "$1" & vbLf & "$2")
        value = System.Text.RegularExpressions.Regex.Replace(value, "\n{3,}", vbLf & vbLf)
        value = NormalizeGeneratedMarkdownHeadingBodies(value)

        Return value.Trim().Replace(vbLf, Environment.NewLine)
    End Function

    Private Function NormalizeGeneratedMarkdownHeadingBodies(markdown As String) As String
        Dim builder As New StringBuilder()
        Dim lines = If(markdown, "").Split(New String() {vbLf}, StringSplitOptions.None)

        For Each rawLine In lines
            Dim line = If(rawLine, "").TrimEnd()
            Dim headingLine As String = Nothing
            Dim bodyLine As String = Nothing

            If TrySplitGeneratedMarkdownHeadingBody(line, headingLine, bodyLine) Then
                builder.Append(headingLine).Append(vbLf)
                builder.Append(vbLf)
                builder.Append(bodyLine).Append(vbLf)
            Else
                builder.Append(line).Append(vbLf)
            End If
        Next

        Dim normalized = builder.ToString().Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).TrimEnd()
        Return System.Text.RegularExpressions.Regex.Replace(normalized, "\n{3,}", vbLf & vbLf)
    End Function

    Private Function TrySplitGeneratedMarkdownHeadingBody(line As String, ByRef headingLine As String, ByRef bodyLine As String) As Boolean
        If String.IsNullOrWhiteSpace(line) OrElse Not line.StartsWith("#", StringComparison.Ordinal) Then Return False

        Dim headingLevel = 0
        While headingLevel < line.Length AndAlso line(headingLevel) = "#"c
            headingLevel += 1
        End While

        If headingLevel < 2 OrElse headingLevel > 6 Then Return False
        If headingLevel >= line.Length OrElse Not Char.IsWhiteSpace(line(headingLevel)) Then Return False

        Dim content = line.Substring(headingLevel).Trim()
        If content.Length < 18 Then Return False

        Dim splitIndex = FindGeneratedHeadingBodySplitIndex(content)
        If splitIndex <= 0 Then Return False

        Dim headingText = content.Substring(0, splitIndex).Trim()
        Dim bodyText = content.Substring(splitIndex).Trim()
        If headingText.Length < 4 OrElse bodyText.Length < 8 Then Return False

        headingLine = New String("#"c, headingLevel) & " " & headingText
        bodyLine = bodyText
        Return True
    End Function

    Private Function FindGeneratedHeadingBodySplitIndex(content As String) As Integer
        Dim markers = New String() {
            " 据", " 根据", " 预计", " 其中", " 海外", " 国内", " 用户", " 企业",
            " 核心", " 目前", " 通过", " 同时", " 此外", " 随着", " Microsoft",
            " Google", " IDC", " Gartner", "据IDC", "根据IDC", "据Gartner", "根据Gartner"
        }

        Dim best = -1
        For Each marker In markers
            Dim startIndex = If(marker.StartsWith(" ", StringComparison.Ordinal), 4, 6)
            Dim index = content.IndexOf(marker, startIndex, StringComparison.OrdinalIgnoreCase)
            If index >= 0 AndAlso (best < 0 OrElse index < best) Then
                best = index
            End If
        Next

        Return best
    End Function

    Private Sub ApplyMarkdownEditorHighlight()
        If _applyingMarkdownEditorHighlight Then Return
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() ApplyMarkdownEditorHighlight(), MethodInvoker))
            Return
        End If
        If _outlineEditor.IsDisposed Then Return

        _applyingMarkdownEditorHighlight = True
        Dim redrawSuspended = False
        Try
            Dim text = If(_outlineEditor.Text, "")
            Dim selectionStart = Math.Min(_outlineEditor.SelectionStart, text.Length)
            Dim selectionLength = Math.Min(_outlineEditor.SelectionLength, Math.Max(0, text.Length - selectionStart))
            Dim firstVisibleLine = GetOutlineEditorFirstVisibleLine()

            _outlineEditor.SuspendLayout()
            If _outlineEditor.IsHandleCreated Then
                SetOutlineEditorRedraw(False)
                redrawSuspended = True
            End If

            _outlineEditor.SelectAll()
            _outlineEditor.SelectionFont = New Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular)
            _outlineEditor.SelectionColor = Color.FromArgb(31, 41, 55)
            _outlineEditor.SelectionBackColor = Color.White

            Dim index = 0
            While index < text.Length
                Dim lineEnd = text.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf}, index)
                Dim lineLength = If(lineEnd >= 0, lineEnd - index, text.Length - index)
                If lineLength > 0 Then
                    HighlightMarkdownEditorLine(text.Substring(index, lineLength), index, lineLength)
                End If

                If lineEnd < 0 Then Exit While
                If text(lineEnd) = ControlChars.Cr AndAlso lineEnd + 1 < text.Length AndAlso text(lineEnd + 1) = ControlChars.Lf Then
                    index = lineEnd + 2
                Else
                    index = lineEnd + 1
                End If
            End While

            _outlineEditor.Select(selectionStart, selectionLength)
            RestoreOutlineEditorFirstVisibleLine(firstVisibleLine)
        Catch ex As Exception
            AppendThemePptLog("Markdown editor highlight failed: " & ex.ToString())
        Finally
            If redrawSuspended Then
                SetOutlineEditorRedraw(True)
            End If
            _outlineEditor.ResumeLayout()
            _applyingMarkdownEditorHighlight = False
        End Try
    End Sub

    Private Function GetOutlineEditorFirstVisibleLine() As Integer
        If _outlineEditor Is Nothing OrElse _outlineEditor.IsDisposed OrElse Not _outlineEditor.IsHandleCreated Then Return -1
        Return SendMessage(_outlineEditor.Handle, EM_GETFIRSTVISIBLELINE, IntPtr.Zero, IntPtr.Zero).ToInt32()
    End Function

    Private Sub RestoreOutlineEditorFirstVisibleLine(firstVisibleLine As Integer)
        If firstVisibleLine < 0 Then Return
        If _outlineEditor Is Nothing OrElse _outlineEditor.IsDisposed OrElse Not _outlineEditor.IsHandleCreated Then Return

        Dim currentVisibleLine = GetOutlineEditorFirstVisibleLine()
        If currentVisibleLine < 0 Then Return

        Dim delta = firstVisibleLine - currentVisibleLine
        If delta <> 0 Then
            SendMessage(_outlineEditor.Handle, EM_LINESCROLL, IntPtr.Zero, New IntPtr(delta))
        End If
    End Sub

    Private Sub SetOutlineEditorRedraw(enabled As Boolean)
        If _outlineEditor Is Nothing OrElse _outlineEditor.IsDisposed OrElse Not _outlineEditor.IsHandleCreated Then Return

        SendMessage(_outlineEditor.Handle, WM_SETREDRAW, If(enabled, New IntPtr(1), IntPtr.Zero), IntPtr.Zero)
        If enabled Then
            _outlineEditor.Invalidate()
        End If
    End Sub

    Private Sub HighlightMarkdownEditorLine(line As String, lineStart As Integer, lineLength As Integer)
        Dim trimmed = line.TrimStart()
        If String.IsNullOrWhiteSpace(trimmed) Then Return

        Dim leadingSpaces = line.Length - trimmed.Length
        Dim style As FontStyle = FontStyle.Regular
        Dim size As Single = 9.5F
        Dim highlightColor As Color = Color.FromArgb(31, 41, 55)

        If trimmed.StartsWith("#", StringComparison.Ordinal) Then
            Dim level = 0
            While level < trimmed.Length AndAlso trimmed(level) = "#"c
                level += 1
            End While

            If level > 0 AndAlso level <= 6 AndAlso level < trimmed.Length AndAlso Char.IsWhiteSpace(trimmed(level)) Then
                style = FontStyle.Bold
                If level = 1 Then
                    size = 12.0F
                    highlightColor = Color.FromArgb(153, 27, 27)
                ElseIf level = 2 Then
                    size = 10.5F
                    highlightColor = Color.FromArgb(194, 65, 12)
                Else
                    size = 9.8F
                    highlightColor = Color.FromArgb(75, 85, 99)
                End If
            End If
        ElseIf trimmed.StartsWith("- ", StringComparison.Ordinal) OrElse
               trimmed.StartsWith("* ", StringComparison.Ordinal) OrElse
               trimmed.StartsWith("+ ", StringComparison.Ordinal) OrElse
               System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\d{1,2}[\.\)]\s+") Then
            highlightColor = Color.FromArgb(55, 65, 81)
        ElseIf trimmed.StartsWith("```", StringComparison.Ordinal) Then
            highlightColor = Color.FromArgb(37, 99, 235)
            style = FontStyle.Bold
        End If

        _outlineEditor.Select(lineStart, lineLength)
        _outlineEditor.SelectionFont = New Font("Microsoft YaHei UI", size, style)
        _outlineEditor.SelectionColor = highlightColor

        If leadingSpaces > 0 AndAlso lineLength > leadingSpaces Then
            _outlineEditor.Select(lineStart, leadingSpaces)
            _outlineEditor.SelectionColor = Color.FromArgb(156, 163, 175)
        End If
    End Sub

    Private Sub MarkOutlineEditingRequired()
        _isOutlineEditCompleted = False
        _templateConfirmedForCurrentOutline = False
        _confirmedTemplateId = ""
        RefreshActionButtons()
    End Sub

    Private Sub ClearGeneratedPptState()
        _lastGeneratedPptId = ""
        _lastImportedSlideStartIndex = 0
        _lastImportedSlideCount = 0
    End Sub

    Private Sub ApplyOutlineEditCompletion()
        _outlineMarkdown = GetEditedMarkdown()
        _isOutlineEditCompleted = True
        _templateConfirmedForCurrentOutline = False
        _confirmedTemplateId = ""
        SetTemplateSectionVisible(True)
        RefreshActionButtons()
    End Sub

    Private Async Sub FinishOutlineEditButton_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(GetEditedMarkdown()) Then
            MessageBox.Show("请先填写 Markdown 大纲。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            ValidateEditedMarkdownForDocmee(GetEditedMarkdown())
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Markdown 大纲", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End Try

        ApplyOutlineEditCompletion()
        SetStatus("Markdown 大纲编辑完成，正在加载模板...")

        If _templateCombo.Items.Count = 0 Then
            Await LoadTemplatesAsync()
        Else
            RefreshActionButtons()
        End If

        If _templateCombo.Items.Count > 0 AndAlso _lastTemplateLoadUsedFallback Then
            SetStatus($"Markdown 大纲编辑完成，模板接口失败，已使用内置模板 {_templateCombo.Items.Count} 个，请预览并选择模板后生成。")
        ElseIf _templateCombo.Items.Count > 0 Then
            SetStatus("Markdown 大纲编辑完成，请预览并选择模板后生成。")
        End If
    End Sub

    Private Sub OutlineEditor_TextChanged(sender As Object, e As EventArgs)
        If Not _suppressOutlineEditorChange Then
            ApplyMarkdownEditorHighlight()
        End If

        ScheduleMarkdownPreviewUpdate()

        If _suppressOutlineEditorChange Then Return

        Dim hadCompletedEdit = _isOutlineEditCompleted OrElse _templateConfirmedForCurrentOutline
        _isOutlineEditCompleted = False
        _templateConfirmedForCurrentOutline = False
        _confirmedTemplateId = ""
        ClearGeneratedPptState()
        RefreshActionButtons()

        If hadCompletedEdit Then
            SetStatus("Markdown 大纲已修改，请先完成编辑，再选择模板生成。")
        End If
    End Sub

    Private Sub UpdateMarkdownPreview()
        ScheduleMarkdownPreviewUpdate(True)
    End Sub

    Private Sub ScheduleMarkdownPreviewUpdate(Optional immediate As Boolean = False)
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() ScheduleMarkdownPreviewUpdate(immediate), MethodInvoker))
            Return
        End If

        _pendingOutlinePreviewRender = True
        _markdownPreviewRenderGeneration += 1
        _outlinePreviewDebounceTimer.Stop()

        Try
            If immediate Then
                BeginInvokeIfAlive(CType(Sub() OutlinePreviewDebounceTimer_Tick(_outlinePreviewDebounceTimer, EventArgs.Empty), MethodInvoker))
            Else
                _outlinePreviewDebounceTimer.Start()
            End If
        Catch ex As Exception
            _pendingOutlinePreviewRender = True
            AppendThemePptLog("Schedule markdown preview failed: " & ex.ToString())
        End Try
    End Sub

    Private Function BeginInvokeIfAlive(action As MethodInvoker) As Boolean
        If action Is Nothing OrElse Me.IsDisposed OrElse Not Me.IsHandleCreated Then Return False

        Try
            Me.BeginInvoke(action)
            Return True
        Catch ex As ObjectDisposedException
            Return False
        Catch ex As InvalidOperationException
            Return False
        End Try
    End Function

    Private Function IsOnPaneUiThread() As Boolean
        Return Thread.CurrentThread.ManagedThreadId = _uiThreadId AndAlso Not Me.InvokeRequired
    End Function

    Private Async Sub OutlinePreviewDebounceTimer_Tick(sender As Object, e As EventArgs)
        Try
            If Not IsOnPaneUiThread() Then
                BeginInvokeIfAlive(CType(Sub() OutlinePreviewDebounceTimer_Tick(sender, e), MethodInvoker))
                Return
            End If

            _outlinePreviewDebounceTimer.Stop()

            If Not _outlinePreviewReady OrElse _outlinePreviewWebView.CoreWebView2 Is Nothing Then
                _pendingOutlinePreviewRender = True
                Return
            End If

            _pendingOutlinePreviewRender = False
            Dim renderGeneration = _markdownPreviewRenderGeneration
            Dim markdownSnapshot = GetEditedMarkdown()
            Dim previewHtml = Await Task.Run(Function() BuildMarkdownPreviewHtml(markdownSnapshot))

            If Not IsOnPaneUiThread() Then
                BeginInvokeIfAlive(CType(Sub() NavigateMarkdownPreview(renderGeneration, previewHtml), MethodInvoker))
                Return
            End If

            NavigateMarkdownPreview(renderGeneration, previewHtml)
        Catch ex As Exception
            _pendingOutlinePreviewRender = True
            AppendThemePptLog("Markdown preview update failed: " & ex.ToString())
        End Try
    End Sub

    Private Sub NavigateMarkdownPreview(renderGeneration As Integer, previewHtml As String)
        Try
            If Not IsOnPaneUiThread() Then
                BeginInvokeIfAlive(CType(Sub() NavigateMarkdownPreview(renderGeneration, previewHtml), MethodInvoker))
                Return
            End If

            If renderGeneration <> _markdownPreviewRenderGeneration Then Return
            If _outlinePreviewWebView.IsDisposed Then Return

            If Not _outlinePreviewReady OrElse _outlinePreviewWebView.CoreWebView2 Is Nothing Then
                _pendingOutlinePreviewRender = True
                Return
            End If

            _outlinePreviewWebView.NavigateToString(previewHtml)
        Catch ex As Exception
            _pendingOutlinePreviewRender = True
            AppendThemePptLog("Markdown preview navigate failed: " & ex.ToString())
        End Try
    End Sub

    Private Shared Function GetMarkdownPreviewPipeline() As MarkdownPipeline
        If _markdownPreviewPipeline IsNot Nothing Then Return _markdownPreviewPipeline

        SyncLock MarkdownPreviewPipelineLock
            If _markdownPreviewPipeline IsNot Nothing Then Return _markdownPreviewPipeline

            Try
                _markdownPreviewPipeline = New MarkdownPipelineBuilder().UseAdvancedExtensions().DisableHtml().Build()
                _markdownPreviewPipelineInitializeError = Nothing
            Catch ex As Exception
                _markdownPreviewPipelineInitializeError = ex.Message
                AppendThemePptLog("Markdown preview pipeline initialize failed: " & ex.ToString())
            End Try

            Return _markdownPreviewPipeline
        End SyncLock
    End Function

    Private Function BuildMarkdownPreviewHtml(markdown As String) As String
        Dim safeMarkdown = If(markdown, "")
        Dim renderedMarkdown As String

        Try
            Dim previewPipeline = GetMarkdownPreviewPipeline()
            If previewPipeline Is Nothing Then
                Throw New InvalidOperationException("Markdown preview renderer unavailable: " & If(_markdownPreviewPipelineInitializeError, "unknown error"))
            End If

            renderedMarkdown = Markdig.Markdown.ToHtml(safeMarkdown, previewPipeline)
        Catch ex As Exception
            renderedMarkdown = "<pre class=""error"">Markdown 渲染失败：" & EscapeHtml(ex.Message) & "</pre>"
        End Try

        If String.IsNullOrWhiteSpace(renderedMarkdown) Then
            renderedMarkdown = "<p class=""empty"">暂无 Markdown 大纲</p>"
        End If

        Dim builder As New StringBuilder()
        builder.AppendLine("<!doctype html>")
        builder.AppendLine("<html lang=""zh-CN""><head><meta charset=""utf-8"">")
        builder.AppendLine("<meta name=""viewport"" content=""width=device-width,initial-scale=1"">")
        builder.AppendLine("<style>")
        builder.AppendLine("html,body{margin:0;padding:0;background:#fff;color:#1f2937;font-family:'Microsoft YaHei UI','Segoe UI',Arial,sans-serif;font-size:13px;line-height:1.62;}")
        builder.AppendLine(".markdown-body{box-sizing:border-box;padding:12px 14px 18px;word-break:break-word;}")
        builder.AppendLine("h1{font-size:22px;line-height:1.25;margin:0 0 12px;color:#111827;border-bottom:1px solid #e5e7eb;padding-bottom:8px;}")
        builder.AppendLine("h2{font-size:17px;line-height:1.35;margin:18px 0 8px;color:#1f2937;}h3{font-size:15px;margin:14px 0 6px;color:#374151;}")
        builder.AppendLine("p{margin:8px 0;}ul,ol{padding-left:22px;margin:8px 0;}li{margin:4px 0;}blockquote{margin:10px 0;padding:8px 12px;border-left:4px solid #f97316;background:#fff7ed;color:#4b5563;}")
        builder.AppendLine("code{font-family:Consolas,'Courier New',monospace;background:#f3f4f6;border-radius:4px;padding:1px 4px;}pre{background:#111827;color:#f9fafb;border-radius:6px;padding:10px;overflow:auto;}pre code{background:transparent;color:inherit;padding:0;}")
        builder.AppendLine("table{border-collapse:collapse;width:100%;margin:10px 0;}th,td{border:1px solid #d1d5db;padding:6px 8px;text-align:left;}th{background:#f3f4f6;}a{color:#2563eb}.empty{color:#6b7280}.error{background:#fef2f2;color:#991b1b;white-space:pre-wrap;}")
        builder.AppendLine("</style></head><body><main class=""markdown-body"">")
        builder.AppendLine(renderedMarkdown)
        builder.AppendLine("</main></body></html>")
        Return builder.ToString()
    End Function

    Private Shared Function TryGetString(token As JToken) As String
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return ""
        Return token.ToString()
    End Function

    Private Shared Function EscapeHtml(value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        Return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("""", "&quot;").Replace("'", "&#39;")
    End Function

    Private Shared Function EscapeHtmlAttribute(value As String) As String
        Return EscapeHtml(value)
    End Function

    Private Shared Function GetTemplateCoverCacheDirectory() As String
        Return Path.Combine(Path.GetTempPath(), "OfficeAiThemePptCovers")
    End Function

    Friend Shared Function GetThemePptLogPath() As String
        Return Path.Combine(Path.GetTempPath(), "OfficeAiThemePpt", "theme-ppt.log")
    End Function

    Friend Shared Sub AppendThemePptLog(message As String)
        Try
            Dim logPath = GetThemePptLogPath()
            Dim logDirectory = Path.GetDirectoryName(logPath)
            If Not String.IsNullOrWhiteSpace(logDirectory) Then Directory.CreateDirectory(logDirectory)
            File.AppendAllText(logPath,
                               DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") & " [" & ThemePptPaneBuild & "] " & message & Environment.NewLine,
                               Encoding.UTF8)
        Catch
        End Try
    End Sub

    Private Shared Function GetSafeFileName(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return Guid.NewGuid().ToString("N")

        Dim builder As New StringBuilder()
        For Each ch In value
            If Char.IsLetterOrDigit(ch) OrElse ch = "_"c OrElse ch = "-"c Then
                builder.Append(ch)
            Else
                builder.Append("_"c)
            End If
        Next

        If builder.Length = 0 Then Return Guid.NewGuid().ToString("N")
        Return builder.ToString()
    End Function

    Private Shared Function GetImageFileExtension(bytes As Byte()) As String
        If bytes IsNot Nothing AndAlso bytes.Length >= 4 Then
            If bytes(0) = &HFF AndAlso bytes(1) = &HD8 Then Return ".jpg"
            If bytes(0) = &H89 AndAlso bytes(1) = &H50 AndAlso bytes(2) = &H4E AndAlso bytes(3) = &H47 Then Return ".png"
            If bytes(0) = &H47 AndAlso bytes(1) = &H49 AndAlso bytes(2) = &H46 Then Return ".gif"
            If bytes.Length >= 12 AndAlso bytes(0) = &H52 AndAlso bytes(1) = &H49 AndAlso bytes(2) = &H46 AndAlso bytes(3) = &H46 AndAlso bytes(8) = &H57 AndAlso bytes(9) = &H45 AndAlso bytes(10) = &H42 AndAlso bytes(11) = &H50 Then Return ".webp"
        End If

        Return ".png"
    End Function

    Private Function SaveTemplateCoverFile(bytes As Byte(), templateId As String, loadGeneration As Integer) As String
        If bytes Is Nothing OrElse bytes.Length = 0 Then Return ""

        Dim cacheDirectory = GetTemplateCoverCacheDirectory()
        Directory.CreateDirectory(cacheDirectory)
        Dim fileName = GetSafeFileName(templateId) & "_" & loadGeneration.ToString() & "_" & Guid.NewGuid().ToString("N") & GetImageFileExtension(bytes)
        Dim filePath = Path.Combine(cacheDirectory, fileName)
        File.WriteAllBytes(filePath, bytes)

        If _templateCoverFilePaths.ContainsKey(templateId) Then
            Try
                File.Delete(_templateCoverFilePaths(templateId))
            Catch
            End Try
        End If

        _templateCoverFilePaths(templateId) = filePath
        AppendThemePptLog("Cover file saved: id=" & templateId & ", path=" & filePath & ", bytes=" & bytes.Length.ToString())
        Return "https://" & TemplateCoverHostName & "/" & fileName
    End Function

    Private Async Sub GenerateButton_Click(sender As Object, e As EventArgs)
        Await GenerateOutlineAsync()
    End Sub

    Private Async Sub RefreshTemplatesButton_Click(sender As Object, e As EventArgs)
        If Not _isOutlineEditCompleted Then
            MessageBox.Show("请先完成 Markdown 大纲编辑。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Await LoadTemplatesAsync(_templatePage)
    End Sub

    Private Async Sub SelectTemplateButton_Click(sender As Object, e As EventArgs)
        If _isTemplateLoading Then Return

        If Not _isOutlineEditCompleted Then
            MessageBox.Show("请先完成 Markdown 大纲编辑。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If _templateCombo.Items.Count = 0 Then
            Await LoadTemplatesAsync()
        End If

        If _templateCombo.Items.Count = 0 Then Return
        ShowTemplateSelectionDialog()
    End Sub

    Private Sub ConfigureDocmeeButton_Click(sender As Object, e As EventArgs)
        ShowDocmeeSettingsDialog()
    End Sub

    Private Sub ShowDocmeeSettingsDialog()
        Using dialog As New Form()
            dialog.Text = "Docmee配置"
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MaximizeBox = False
            dialog.MinimizeBox = False
            dialog.ClientSize = New Size(460, 182)

            Dim layout As New TableLayoutPanel()
            layout.Dock = DockStyle.Fill
            layout.Padding = New Padding(14)
            layout.ColumnCount = 2
            layout.RowCount = 4
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 86.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 38.0F))

            Dim baseUrlLabel As New Label() With {
                .Text = "接口地址",
                .AutoSize = True,
                .Margin = New Padding(0, 7, 8, 0)
            }
            Dim baseUrlBox As New TextBox() With {
                .Dock = DockStyle.Fill,
                .Text = DocmeePptClient.GetConfiguredApiBaseUrl()
            }

            Dim tokenLabel As New Label() With {
                .Text = "Token",
                .AutoSize = True,
                .Margin = New Padding(0, 7, 8, 0)
            }
            Dim tokenBox As New TextBox() With {
                .Dock = DockStyle.Fill,
                .Text = DocmeePptClient.GetConfiguredToken(),
                .UseSystemPasswordChar = True
            }

            Dim hintLabel As New Label() With {
                .Dock = DockStyle.Fill,
                .ForeColor = Color.FromArgb(86, 94, 108),
                .Text = "正式环境请填写 Docmee 开放平台地址和 token；留空会继续使用配置文件、环境变量或测试默认值。"
            }

            Dim actionPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft
            }
            Dim okButton As New Button() With {
                .Text = "保存",
                .Width = 78,
                .Height = 28,
                .DialogResult = DialogResult.OK
            }
            Dim cancelButton As New Button() With {
                .Text = "取消",
                .Width = 78,
                .Height = 28,
                .DialogResult = DialogResult.Cancel
            }
            actionPanel.Controls.Add(okButton)
            actionPanel.Controls.Add(cancelButton)

            layout.Controls.Add(baseUrlLabel, 0, 0)
            layout.Controls.Add(baseUrlBox, 1, 0)
            layout.Controls.Add(tokenLabel, 0, 1)
            layout.Controls.Add(tokenBox, 1, 1)
            layout.Controls.Add(hintLabel, 0, 2)
            layout.SetColumnSpan(hintLabel, 2)
            layout.Controls.Add(actionPanel, 0, 3)
            layout.SetColumnSpan(actionPanel, 2)

            dialog.Controls.Add(layout)
            dialog.AcceptButton = okButton
            dialog.CancelButton = cancelButton

            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return

            Dim apiBaseUrl = baseUrlBox.Text.Trim()
            If Not String.IsNullOrWhiteSpace(apiBaseUrl) AndAlso
               Not apiBaseUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) AndAlso
               Not apiBaseUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("Docmee 接口地址需要以 http:// 或 https:// 开头。", "Docmee配置", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ShareRibbon.ConfigSettings.SaveDocmeeSettings(apiBaseUrl, tokenBox.Text)
            _templateConfirmedForCurrentOutline = False
            _confirmedTemplateId = ""
            _templateCombo.Items.Clear()
            _templateCombo.Enabled = False
            _selectTemplateButton.Enabled = False
            CancelTemplateCoverLoad()
            ClearTemplateCoverImages()
            SetStatus("Docmee 配置已保存，请刷新模板或重新生成大纲。")
            RefreshActionButtons()
        End Using
    End Sub

    Private Async Function GenerateOutlineAsync() As Task
        Dim mode = GetSelectedGenerationMode()

        _generateButton.Enabled = False
        _insertButton.Enabled = False
        _finishOutlineEditButton.Enabled = False
        ShowOutlineOutput()
        _outputBox.Clear()
        _outline = Nothing
        _outlineMarkdown = ""
        SetOutlineEditorText("")
        _isOutlineEditCompleted = False
        _templateConfirmedForCurrentOutline = False
        _confirmedTemplateId = ""
        ClearGeneratedPptState()
        RefreshActionButtons()
        SetTopicBoxVisible(False)

        Try
            Select Case mode
                Case GenerationModeDocument
                    _outlineMarkdown = Await GenerateOutlineFromDocumentAsync()
                Case GenerationModeMarkdown
                    _outlineMarkdown = PrepareMarkdownOutlineFromInput()
                    _taskId = ""
                Case Else
                    Dim topic = _topicBox.Text.Trim()
                    If String.IsNullOrWhiteSpace(topic) Then
                        MessageBox.Show("请输入 PPT 主题。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If

                    Dim requestContent = BuildRequestContent(topic)
                    SetStatus("正在创建 Docmee 主题任务...")
                    _taskId = Await _client.CreateTaskAsync(requestContent)

                    SetStatus("正在生成 PPT Markdown 大纲...")
                    _outputBox.Clear()
                    _outlineMarkdown = Await _client.GenerateMarkdownContentAsync(_taskId, AddressOf AppendOutlineStreamText)
            End Select

            SetOutlineEditorText(_outlineMarkdown.Trim())
            MarkOutlineEditingRequired()
            ShowOutlineEditor()
            SetStatus("大纲已生成，请编辑 Markdown，完成编辑后再选择模板生成。")
        Catch ex As Exception
            AppendThemePptLog("GenerateOutlineAsync exception: " & ex.ToString())
            SetStatus("生成失败。")
            ShowOutlineOutput()
            SetTopicBoxVisible(True)
            MessageBox.Show("主题生成PPT失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _generateButton.Enabled = True
            RefreshActionButtons()
        End Try
    End Function

    Private Async Function GenerateOutlineFromDocumentAsync() As Task(Of String)
        Dim documentPath = GetSelectedDocumentPath()
        If String.IsNullOrWhiteSpace(documentPath) Then
            Throw New InvalidOperationException("请先选择要生成 PPT 的文档。")
        End If
        If Not File.Exists(documentPath) Then
            Throw New FileNotFoundException("未找到要生成 PPT 的文档。", documentPath)
        End If

        SetStatus("正在上传文档创建 Docmee 任务...")
        AppendTaskPaneLine("文档路径: " & documentPath)
        _taskId = Await _client.CreateFileTaskAsync(documentPath)

        SetStatus("正在根据文档生成 PPT Markdown 大纲...")
        _outputBox.Clear()
        Return Await _client.GenerateMarkdownContentAsync(_taskId, AddressOf AppendOutlineStreamText, GetDocumentPrompt())
    End Function

    Private Function GetDocumentPrompt() As String
        Dim text = _topicBox.Text.Trim()
        If text.StartsWith("可选：", StringComparison.Ordinal) Then Return ""
        Return text
    End Function

    Private Function PrepareMarkdownOutlineFromInput() As String
        Dim markdown = _topicBox.Text.Trim()
        If String.IsNullOrWhiteSpace(markdown) Then
            Throw New InvalidOperationException("请先粘贴 Markdown 大纲。")
        End If
        If Not markdown.StartsWith("#", StringComparison.Ordinal) Then
            Throw New InvalidOperationException("Markdown 大纲应以 # 一级标题开始。")
        End If

        Return NormalizeMarkdownForEditing(markdown)
    End Function

    Private Sub ValidateEditedMarkdownForDocmee(markdown As String)
        If String.IsNullOrWhiteSpace(markdown) Then
            Throw New InvalidOperationException("请先填写 Markdown 大纲。")
        End If

        Dim h1Count = 0
        Dim h2Count = 0
        Dim normalized = markdown.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        For Each rawLine In normalized.Split(New String() {vbLf}, StringSplitOptions.None)
            Dim line = rawLine.Trim()
            If Not line.StartsWith("#", StringComparison.Ordinal) Then Continue For

            Dim headingLevel = 0
            While headingLevel < line.Length AndAlso line(headingLevel) = "#"c
                headingLevel += 1
            End While

            Dim headingText = line.Substring(headingLevel).Trim()
            If String.IsNullOrWhiteSpace(headingText) Then
                Throw New InvalidOperationException("Markdown 标题不能只有 #，请补充标题内容。")
            End If

            If headingLevel = 1 Then h1Count += 1
            If headingLevel = 2 Then h2Count += 1
        Next

        If h1Count <> 1 Then
            Throw New InvalidOperationException("Docmee Markdown 大纲需要且只能有一个一级标题：# 主题。")
        End If
        If h2Count < 1 Then
            Throw New InvalidOperationException("Docmee Markdown 大纲至少需要一个二级章节：## 章节。")
        End If
    End Sub

    Private Sub AppendOutlineStreamText(chunkText As String)
        If String.IsNullOrEmpty(chunkText) Then Return
        If Me.IsDisposed OrElse _outputBox.IsDisposed Then Return

        If _outputBox.InvokeRequired Then
            BeginInvokeIfAlive(CType(Sub() AppendOutlineStreamText(chunkText), MethodInvoker))
            Return
        End If

        _outputBox.AppendText(chunkText)
        _outputBox.SelectionStart = _outputBox.TextLength
        _outputBox.ScrollToCaret()
    End Sub

    Private Async Sub InsertButton_Click(sender As Object, e As EventArgs)
        Await GenerateAndImportPptxAsync()
    End Sub

    Private Async Function LoadTemplatesAsync(Optional page As Integer = 1) As Task
        If _isTemplateLoading Then Return

        Dim requestedPage = Math.Max(1, page)
        _isTemplateLoading = True
        AppendThemePptLog("LoadTemplatesAsync start. page=" & requestedPage.ToString())
        _refreshTemplatesButton.Enabled = False
        _selectTemplateButton.Enabled = False
        _lastTemplateLoadUsedFallback = False
        CancelTemplateCoverLoad()
        CancelTemplateLoad()
        _templateLoadCts = New CancellationTokenSource()
        Dim loadCts = _templateLoadCts
        Dim cancellationToken = loadCts.Token

        Try
            SetStatus("正在加载 Docmee 模板...")
            Await Task.Yield()
            Dim templates = Await LoadTemplatesInBackgroundAsync(requestedPage, cancellationToken)
            cancellationToken.ThrowIfCancellationRequested()
            AppendThemePptLog("LoadTemplatesAsync fetched count=" & If(templates Is Nothing, 0, templates.Count).ToString())
            _templatePage = requestedPage
            _templateHasNextPage = templates IsNot Nothing AndAlso templates.Count >= TemplatePageSize
            PopulateTemplates(templates)

            If _templateCombo.Items.Count = 0 Then
                SetStatus("未获取到可用模板。")
            Else
                SetStatus($"已加载 {_templateCombo.Items.Count} 个模板。")
            End If
        Catch ex As OperationCanceledException
            SetStatus("模板加载已取消。")
        Catch ex As Exception
            AppendThemePptLog("LoadTemplatesAsync exception: " & ex.ToString())
            AppendTemplateLoadFailure(ex)

            Dim fallbackTemplates = DocmeePptClient.GetFallbackTemplates()
            If fallbackTemplates.Count > 0 Then
                _lastTemplateLoadUsedFallback = True
                _templatePage = 1
                _templateHasNextPage = False
                PopulateTemplates(fallbackTemplates)
                SetStatus($"模板接口失败，已使用内置模板 {fallbackTemplates.Count} 个。")
                ShowTemplateGallery()
            Else
                _templateCombo.Enabled = False
                SetStatus("模板加载失败。")
                ShowOutlineOutput()
                MessageBox.Show("加载模板失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Finally
            _isTemplateLoading = False
            If Object.ReferenceEquals(_templateLoadCts, loadCts) Then
                _templateLoadCts.Dispose()
                _templateLoadCts = Nothing
            End If
            RefreshActionButtons()
        End Try
    End Function

    Private Function LoadTemplatesInBackgroundAsync(page As Integer, cancellationToken As CancellationToken) As Task(Of List(Of DocmeeTemplateInfo))
        Return Task.Run(Function() As List(Of DocmeeTemplateInfo)
                            cancellationToken.ThrowIfCancellationRequested()
                            Return _client.ListTemplatesAsync(Math.Max(1, page), TemplatePageSize, cancellationToken).GetAwaiter().GetResult()
                        End Function, cancellationToken)
    End Function

    Private Sub PopulateTemplates(templates As IEnumerable(Of DocmeeTemplateInfo))
        AppendThemePptLog("PopulateTemplates start.")
        _templateCoverLoadGeneration += 1
        _templateCoverFailureCount = 0
        CancelTemplateCoverLoad()
        ClearTemplateCoverImages()
        _templateCombo.Items.Clear()
        _templateListBox.Items.Clear()
        _templateCards.Clear()
        _templateSelectLabels.Clear()
        _templateCoverHosts.Clear()
        _templateCoverStatusLabels.Clear()
        _templateCoverImageUrls.Clear()
        _templateCoverFilePaths.Clear()
        _templateCoverMessages.Clear()

        _templateCardPanel.SuspendLayout()
        Try
            ClearTemplateCardPanel()

            If templates IsNot Nothing Then
                For Each template In templates
                    If template Is Nothing OrElse String.IsNullOrWhiteSpace(template.Id) Then Continue For
                    _templateCombo.Items.Add(template)
                Next
            End If
        Finally
            _templateCardPanel.ResumeLayout(True)
        End Try

        _templateCombo.Enabled = False
        If _templateCombo.Items.Count > 0 AndAlso _templateCombo.SelectedIndex < 0 Then
            _templateCombo.SelectedIndex = 0
        End If
        RefreshTemplateSelectionStyles()
        RefreshActionButtons()
        ResizeTemplateCards()
        AppendThemePptLog("PopulateTemplates completed: count=" & _templateCombo.Items.Count.ToString() & ", generation=" & _templateCoverLoadGeneration.ToString())
    End Sub

    Private Sub ClearTemplateCardPanel()
        Do While _templateCardPanel.Controls.Count > 0
            Dim card = _templateCardPanel.Controls(0)
            _templateCardPanel.Controls.RemoveAt(0)
            card.Dispose()
        Loop
    End Sub

    Private Sub ClearTemplateCoverImages()
        For Each pair In _templateCoverBoxes
            If pair.Value IsNot Nothing AndAlso pair.Value.Image IsNot Nothing Then
                Dim image = pair.Value.Image
                pair.Value.Image = Nothing
                image.Dispose()
            End If
        Next

        For Each pair In _templateCoverHosts
            If pair.Value IsNot Nothing AndAlso pair.Value.BackgroundImage IsNot Nothing Then
                Dim image = pair.Value.BackgroundImage
                pair.Value.BackgroundImage = Nothing
                image.Dispose()
            End If
        Next

        For Each pair In _templatePreviewImages
            If pair.Value IsNot Nothing Then
                pair.Value.Dispose()
            End If
        Next

        For Each pair In _templateCoverFilePaths
            If Not String.IsNullOrWhiteSpace(pair.Value) Then
                Try
                    File.Delete(pair.Value)
                Catch
                End Try
            End If
        Next

        _templateCoverBoxes.Clear()
        _templateCoverHosts.Clear()
        _templatePreviewImages.Clear()
        _templateCoverImageUrls.Clear()
        _templateCoverFilePaths.Clear()
    End Sub

    Private Async Sub BeginLoadTemplateCovers(loadGeneration As Integer, cancellationToken As CancellationToken)
        Dim templates = GetCurrentTemplatesSnapshot()
        Dim semaphore As New SemaphoreSlim(MaxConcurrentTemplateCoverLoads)
        AppendThemePptLog("BeginLoadTemplateCovers: generation=" & loadGeneration.ToString() & ", templates=" & templates.Count.ToString())

        Try
            Await Task.Delay(200, cancellationToken).ConfigureAwait(False)
            AppendThemePptLog("BeginLoadTemplateCovers scheduling: generation=" & loadGeneration.ToString())

            Dim loadTasks As New List(Of Task)()
            Dim skippedWithoutCover As Integer = 0
            For Each template In templates
                If loadGeneration <> _templateCoverLoadGeneration Then Return
                If cancellationToken.IsCancellationRequested Then Return
                If template Is Nothing Then Continue For
                If String.IsNullOrWhiteSpace(template.CoverUrl) Then
                    skippedWithoutCover += 1
                    AppendThemePptLog("Cover skipped: id=" & template.Id & ", missing coverUrl.")
                    MarkTemplateCoverUnavailable(template.Id, loadGeneration, "模板未返回封面地址")
                    Continue For
                End If

                Dim currentTemplate = template
                loadTasks.Add(LoadSingleTemplateCoverAsync(currentTemplate, loadGeneration, semaphore, cancellationToken))
            Next
            AppendThemePptLog("BeginLoadTemplateCovers scheduled: tasks=" & loadTasks.Count.ToString() &
                              ", skippedWithoutCover=" & skippedWithoutCover.ToString())

            Await Task.WhenAll(loadTasks)
            UpdateTemplateCoverLoadStatus(loadGeneration)
        Catch ex As OperationCanceledException
        Finally
            semaphore.Dispose()
        End Try
    End Sub

    Private Async Function LoadSingleTemplateCoverAsync(template As DocmeeTemplateInfo, loadGeneration As Integer, semaphore As SemaphoreSlim, cancellationToken As CancellationToken) As Task
        Try
            Await semaphore.WaitAsync(cancellationToken)
            Try
                If loadGeneration <> _templateCoverLoadGeneration Then Return
                cancellationToken.ThrowIfCancellationRequested()

                Dim bytes As Byte() = Nothing
                Dim lastError As Exception = Nothing
                For attempt As Integer = 1 To 3
                    Try
                        AppendThemePptLog("Cover download start: id=" & template.Id & ", attempt=" & attempt.ToString())
                        bytes = Await DownloadTemplateCoverInBackgroundAsync(BuildTemplateCoverUrl(template.CoverUrl), cancellationToken)
                        AppendThemePptLog("Cover download success: id=" & template.Id & ", bytes=" & If(bytes Is Nothing, 0, bytes.Length).ToString())
                        Exit For
                    Catch ex As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        AppendThemePptLog("Cover download attempt failed: id=" & template.Id & ", attempt=" & attempt.ToString() & ", error=" & ex.Message)
                        lastError = ex
                    End Try

                    If bytes IsNot Nothing Then Exit For
                    If attempt < 3 Then Await Task.Delay(300 * attempt, cancellationToken)
                Next

                If bytes Is Nothing AndAlso lastError IsNot Nothing Then Throw lastError
                If loadGeneration <> _templateCoverLoadGeneration Then Return
                cancellationToken.ThrowIfCancellationRequested()

                Dim coverVirtualUrl = ""
                Dim coverImage As System.Drawing.Image = Await Task.Run(Function() LoadTemplateCoverImageFromBytes(bytes, template.Id), cancellationToken)

                SetTemplateCoverImage(template.Id, coverImage, loadGeneration, coverVirtualUrl)
            Finally
                semaphore.Release()
            End Try
        Catch ex As OperationCanceledException
        Catch ex As Exception
            AppendThemePptLog("Cover unavailable: id=" & template.Id & ", error=" & ex.ToString())
            MarkTemplateCoverUnavailable(template.Id, loadGeneration, ex.Message)
        End Try
    End Function

    Private Sub UpdateTemplateCoverLoadStatus(loadGeneration As Integer)
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() UpdateTemplateCoverLoadStatus(loadGeneration), MethodInvoker))
            Return
        End If
        If Me.IsDisposed Then Return

        If loadGeneration <> _templateCoverLoadGeneration Then Return

        Dim total = _templateListBox.Items.Count
        If total <= 0 Then Return

        Dim failed = 0
        For Each pair In _templateCoverMessages
            If Not String.IsNullOrWhiteSpace(pair.Value) AndAlso
               pair.Value.StartsWith("封面加载失败", StringComparison.Ordinal) Then
                failed += 1
            End If
        Next

        Dim success = _templatePreviewImages.Count
        If success > 0 AndAlso failed = 0 Then
            SetStatus("模板封面加载完成。")
        ElseIf success > 0 Then
            SetStatus("已显示 " & success & "/" & total & " 张模板封面，" & failed & " 张失败，卡片内已显示原因。")
        ElseIf failed > 0 Then
            SetStatus("模板封面加载失败，卡片内已显示具体原因。")
        End If
    End Sub

    Private Function DownloadTemplateCoverInBackgroundAsync(coverUrl As String, cancellationToken As CancellationToken) As Task(Of Byte())
        Return Task.Run(Function() As Byte()
                            cancellationToken.ThrowIfCancellationRequested()
                            Return _client.DownloadTemplateCoverAsync(coverUrl, cancellationToken).GetAwaiter().GetResult()
                        End Function, cancellationToken)
    End Function

    Private Function LoadTemplateCoverImageFromBytes(bytes As Byte(), templateId As String) As System.Drawing.Image
        If bytes Is Nothing OrElse bytes.Length = 0 Then
            Throw New InvalidOperationException("模板封面图片为空。")
        End If

        Dim safeId = If(String.IsNullOrWhiteSpace(templateId), Guid.NewGuid().ToString("N"), templateId)
        Dim tempPath = Path.Combine(Path.GetTempPath(), "wenduoduo_cover_" & safeId & "_" & Guid.NewGuid().ToString("N") & ".png")
        File.WriteAllBytes(tempPath, bytes)

        Try
            Using loaded As System.Drawing.Image = System.Drawing.Image.FromFile(tempPath)
                Return CType(New Bitmap(loaded), System.Drawing.Image)
            End Using
        Finally
            Try
                File.Delete(tempPath)
            Catch
            End Try
        End Try
    End Function

    Private Sub CancelTemplateLoad()
        If _templateLoadCts IsNot Nothing Then
            _templateLoadCts.Cancel()
            _templateLoadCts.Dispose()
            _templateLoadCts = Nothing
        End If
    End Sub

    Private Sub CancelTemplateCoverLoad()
        If _templateCoverCts IsNot Nothing Then
            _templateCoverCts.Cancel()
            _templateCoverCts.Dispose()
            _templateCoverCts = Nothing
        End If
    End Sub

    Private Function GetCurrentTemplatesSnapshot() As List(Of DocmeeTemplateInfo)
        Dim templates As New List(Of DocmeeTemplateInfo)()
        For Each item In _templateCombo.Items
            Dim template = TryCast(item, DocmeeTemplateInfo)
            If template IsNot Nothing Then templates.Add(template)
        Next
        Return templates
    End Function

    Private Sub SetTemplateCoverImage(templateId As String, image As Image, loadGeneration As Integer, Optional coverVirtualUrl As String = "")
        If Not IsOnPaneUiThread() Then
            If Not BeginInvokeIfAlive(CType(Sub() SetTemplateCoverImage(templateId, image, loadGeneration, coverVirtualUrl), MethodInvoker)) AndAlso image IsNot Nothing Then
                image.Dispose()
            End If
            Return
        End If
        If Me.IsDisposed Then
            If image IsNot Nothing Then image.Dispose()
            Return
        End If

        If loadGeneration <> _templateCoverLoadGeneration OrElse
           String.IsNullOrWhiteSpace(templateId) Then
            If image IsNot Nothing Then image.Dispose()
            Return
        End If

        Dim oldImage As System.Drawing.Image = Nothing
        If _templateCoverBoxes.ContainsKey(templateId) Then
            Dim cover = _templateCoverBoxes(templateId)
            oldImage = cover.Image
            cover.Image = CreateScaledTemplateImage(image, 640, 360)
            cover.Visible = True
            cover.BringToFront()
            cover.Invalidate()
            If cover.IsHandleCreated Then cover.Update()
        End If

        If _templatePreviewImages.ContainsKey(templateId) Then
            _templatePreviewImages(templateId).Dispose()
        End If
        _templatePreviewImages(templateId) = CreateScaledTemplateImage(image, 320, 180)
        If Not String.IsNullOrWhiteSpace(coverVirtualUrl) Then
            _templateCoverImageUrls(templateId) = coverVirtualUrl
        End If
        _templateCoverMessages(templateId) = ""
        AppendThemePptLog("Cover image set: id=" & templateId &
                          ", size=" & If(image Is Nothing, "null", image.Width.ToString() & "x" & image.Height.ToString()) &
                          ", galleryImages=" & _templatePreviewImages.Count.ToString())
        InvalidateTemplateListItem(templateId)

        If _templateCoverHosts.ContainsKey(templateId) Then
            Dim host = _templateCoverHosts(templateId)
            Dim oldBackground = host.BackgroundImage
            host.BackgroundImage = Nothing
            host.Invalidate()
            If oldBackground IsNot Nothing Then oldBackground.Dispose()
        End If

        If _templateCoverStatusLabels.ContainsKey(templateId) Then
            _templateCoverStatusLabels(templateId).Visible = False
        End If

        If oldImage IsNot Nothing Then oldImage.Dispose()
        If image IsNot Nothing Then image.Dispose()
    End Sub

    Private Sub InvalidateTemplateListItem(templateId As String)
        If _templateListBox Is Nothing OrElse String.IsNullOrWhiteSpace(templateId) Then Return
        If Not _templateListBox.IsHandleCreated OrElse _templateListBox.Items.Count = 0 Then Return

        For index As Integer = 0 To _templateListBox.Items.Count - 1
            Dim item = TryCast(_templateListBox.Items(index), DocmeeTemplateInfo)
            If item IsNot Nothing AndAlso String.Equals(item.Id, templateId, StringComparison.Ordinal) Then
                Dim bounds = _templateListBox.GetItemRectangle(index)
                If bounds.Width > 0 AndAlso bounds.Height > 0 Then
                    _templateListBox.Invalidate(bounds)
                Else
                    _templateListBox.Invalidate()
                End If
                Return
            End If
        Next
    End Sub

    Private Sub MarkTemplateCoverUnavailable(templateId As String, loadGeneration As Integer, Optional reason As String = "")
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() MarkTemplateCoverUnavailable(templateId, loadGeneration, reason), MethodInvoker))
            Return
        End If
        If Me.IsDisposed Then Return

        If loadGeneration <> _templateCoverLoadGeneration OrElse
           String.IsNullOrWhiteSpace(templateId) Then Return

        If _templateCoverBoxes.ContainsKey(templateId) Then
            _templateCoverBoxes(templateId).Visible = False
        End If
        Dim failureMessage = If(String.IsNullOrWhiteSpace(reason), "封面加载失败", reason.Trim())
        If failureMessage.Length > 80 Then failureMessage = failureMessage.Substring(0, 80) & "..."
        _templateCoverMessages(templateId) = "封面加载失败：" & failureMessage
        If _templateCoverImageUrls.ContainsKey(templateId) Then _templateCoverImageUrls.Remove(templateId)
        InvalidateTemplateListItem(templateId)

        If _templateCoverStatusLabels.ContainsKey(templateId) Then
            If _templateCoverHosts.ContainsKey(templateId) Then
                Dim host = _templateCoverHosts(templateId)
                Dim template = TryCast(host.Tag, DocmeeTemplateInfo)
                Dim oldBackground = host.BackgroundImage
                host.BackgroundImage = CreateTemplatePreviewBitmap(template, "封面加载失败：" & failureMessage)
                host.BackgroundImageLayout = ImageLayout.Stretch
                host.Invalidate()
                If oldBackground IsNot Nothing Then oldBackground.Dispose()
            End If
            Dim statusLabel = _templateCoverStatusLabels(templateId)
            statusLabel.Text = "封面加载失败：" & failureMessage
            statusLabel.ForeColor = Color.FromArgb(185, 28, 28)
            statusLabel.Visible = True
            statusLabel.BringToFront()
        End If
        _templateCoverFailureCount += 1

        If _templateCoverFailureCount = 1 Then
            Dim message = If(String.IsNullOrWhiteSpace(reason), "未知错误", reason.Trim())
            If message.Length > 120 Then message = message.Substring(0, 120) & "..."
            SetStatus("部分模板封面加载失败，已显示文字预览: " & message)
        End If
    End Sub

    Private Function CreateTemplateCard(template As DocmeeTemplateInfo) As Panel
        Dim card As New Panel()
        card.Width = Math.Max(220, _templateCardPanel.ClientSize.Width - 24)
        card.Height = 244
        card.Padding = New Padding(8)
        card.Margin = New Padding(0, 0, 0, 10)
        card.BackColor = Color.White
        card.BorderStyle = BorderStyle.FixedSingle
        card.Tag = template
        card.Cursor = Cursors.Hand

        Dim previewPanel As New Panel()
        previewPanel.Name = "TemplateCoverHost"
        previewPanel.BackColor = Color.FromArgb(255, 248, 241)
        previewPanel.BorderStyle = BorderStyle.FixedSingle
        previewPanel.BackgroundImage = CreateTemplatePreviewBitmap(template, "封面加载中...")
        previewPanel.BackgroundImageLayout = ImageLayout.Stretch
        previewPanel.Tag = template
        previewPanel.Cursor = Cursors.Hand

        Dim previewBadge As New Label()
        previewBadge.Name = "TemplatePreviewBadge"
        previewBadge.AutoSize = False
        previewBadge.Text = "模板预览"
        previewBadge.TextAlign = ContentAlignment.MiddleLeft
        previewBadge.ForeColor = Color.FromArgb(234, 88, 12)
        previewBadge.Font = New Font(Me.Font.FontFamily, 8.5F, FontStyle.Bold)
        previewBadge.Tag = template
        previewBadge.Cursor = Cursors.Hand

        Dim previewTitle As New Label()
        previewTitle.Name = "TemplatePreviewTitle"
        previewTitle.AutoSize = False
        previewTitle.Text = If(String.IsNullOrWhiteSpace(template.Name), template.Id, template.Name)
        previewTitle.TextAlign = ContentAlignment.MiddleLeft
        previewTitle.AutoEllipsis = True
        previewTitle.ForeColor = Color.FromArgb(39, 45, 55)
        previewTitle.Font = New Font(Me.Font.FontFamily, 10.0F, FontStyle.Bold)
        previewTitle.Tag = template
        previewTitle.Cursor = Cursors.Hand

        Dim previewMeta As New Label()
        previewMeta.Name = "TemplatePreviewMeta"
        previewMeta.AutoSize = False
        previewMeta.Text = BuildTemplateMetaText(template)
        previewMeta.TextAlign = ContentAlignment.MiddleLeft
        previewMeta.AutoEllipsis = True
        previewMeta.ForeColor = Color.FromArgb(86, 94, 108)
        previewMeta.Font = New Font(Me.Font.FontFamily, 8.5F, FontStyle.Regular)
        previewMeta.Tag = template
        previewMeta.Cursor = Cursors.Hand

        Dim cover As New PictureBox()
        cover.Name = "TemplateCoverImage"
        cover.BackColor = Color.FromArgb(248, 250, 252)
        cover.Dock = DockStyle.Fill
        cover.SizeMode = PictureBoxSizeMode.Zoom
        cover.Tag = template
        cover.Cursor = Cursors.Hand
        cover.Visible = False

        Dim coverStatusLabel As New Label()
        coverStatusLabel.Name = "TemplateCoverStatus"
        coverStatusLabel.AutoSize = False
        coverStatusLabel.Text = "封面加载中..."
        coverStatusLabel.TextAlign = ContentAlignment.MiddleCenter
        coverStatusLabel.ForeColor = Color.FromArgb(86, 94, 108)
        coverStatusLabel.BackColor = Color.Transparent
        coverStatusLabel.Font = New Font(Me.Font.FontFamily, 9.0F, FontStyle.Regular)
        coverStatusLabel.Tag = template
        coverStatusLabel.Cursor = Cursors.Hand

        previewPanel.Controls.Add(previewBadge)
        previewPanel.Controls.Add(previewTitle)
        previewPanel.Controls.Add(previewMeta)
        previewPanel.Controls.Add(cover)
        previewPanel.Controls.Add(coverStatusLabel)

        Dim nameLabel As New Label()
        nameLabel.Name = "TemplateName"
        nameLabel.AutoSize = False
        nameLabel.TextAlign = ContentAlignment.MiddleLeft
        nameLabel.AutoEllipsis = True
        nameLabel.ForeColor = Color.FromArgb(39, 45, 55)
        nameLabel.Font = New Font(Me.Font.FontFamily, 9.0F, FontStyle.Bold)
        nameLabel.Text = If(String.IsNullOrWhiteSpace(template.Name), template.Id, template.Name)
        nameLabel.Tag = template
        nameLabel.Cursor = Cursors.Hand

        Dim detailLabel As New Label()
        detailLabel.Name = "TemplateMeta"
        detailLabel.AutoSize = False
        detailLabel.TextAlign = ContentAlignment.MiddleLeft
        detailLabel.AutoEllipsis = True
        detailLabel.ForeColor = Color.FromArgb(86, 94, 108)
        detailLabel.Font = New Font(Me.Font.FontFamily, 8.0F, FontStyle.Regular)
        detailLabel.Text = BuildTemplateMetaText(template) & If(String.IsNullOrWhiteSpace(template.Id), "", " | ID " & template.Id)
        detailLabel.Tag = template
        detailLabel.Cursor = Cursors.Hand

        Dim selectLabel As New Label()
        selectLabel.Name = "TemplateSelect"
        selectLabel.AutoSize = False
        selectLabel.Text = "选择模板"
        selectLabel.TextAlign = ContentAlignment.MiddleCenter
        selectLabel.ForeColor = Color.FromArgb(39, 45, 55)
        selectLabel.BackColor = Color.FromArgb(241, 245, 249)
        selectLabel.Font = New Font(Me.Font.FontFamily, 9.0F, FontStyle.Bold)
        selectLabel.Tag = template
        selectLabel.Cursor = Cursors.Hand

        card.Controls.Add(previewPanel)
        card.Controls.Add(nameLabel)
        card.Controls.Add(detailLabel)
        card.Controls.Add(selectLabel)
        AddHandler card.Resize, Sub() LayoutTemplateCard(card)
        LayoutTemplateCard(card)
        AddTemplateCardClickHandlers(card, template)

        If Not String.IsNullOrWhiteSpace(template.Id) Then
            _templateCards(template.Id) = card
            _templateSelectLabels(template.Id) = selectLabel
            _templateCoverHosts(template.Id) = previewPanel
            _templateCoverBoxes(template.Id) = cover
            _templateCoverStatusLabels(template.Id) = coverStatusLabel
        End If

        Return card
    End Function

    Private Sub LayoutTemplateCard(card As Panel)
        If card Is Nothing Then Return

        Dim left = card.Padding.Left
        Dim top = card.Padding.Top
        Dim innerWidth = Math.Max(120, card.ClientSize.Width - card.Padding.Horizontal)
        Dim coverHeight = 130
        Dim nameHeight = 28
        Dim metaHeight = 22
        Dim buttonHeight = 30
        Dim gap = 6

        Dim coverHost = FindTemplateCardChild(card, "TemplateCoverHost")
        If coverHost IsNot Nothing Then
            coverHost.Bounds = New Rectangle(left, top, innerWidth, coverHeight)
            LayoutTemplateCoverHost(coverHost)
        End If

        Dim nameLabel = FindTemplateCardChild(card, "TemplateName")
        If nameLabel IsNot Nothing Then
            nameLabel.Bounds = New Rectangle(left, top + coverHeight + gap, innerWidth, nameHeight)
        End If

        Dim metaLabel = FindTemplateCardChild(card, "TemplateMeta")
        If metaLabel IsNot Nothing Then
            metaLabel.Bounds = New Rectangle(left, top + coverHeight + gap + nameHeight, innerWidth, metaHeight)
        End If

        Dim selectLabel = FindTemplateCardChild(card, "TemplateSelect")
        If selectLabel IsNot Nothing Then
            selectLabel.Bounds = New Rectangle(left, top + coverHeight + gap + nameHeight + metaHeight + gap, innerWidth, buttonHeight)
        End If
    End Sub

    Private Sub LayoutTemplateCoverHost(coverHost As Control)
        Dim contentWidth = Math.Max(80, coverHost.ClientSize.Width - 24)
        Dim contentLeft = 12

        Dim previewBadge = FindTemplateCardChild(coverHost, "TemplatePreviewBadge")
        If previewBadge IsNot Nothing Then
            previewBadge.Bounds = New Rectangle(contentLeft, 8, contentWidth, 22)
        End If

        Dim previewTitle = FindTemplateCardChild(coverHost, "TemplatePreviewTitle")
        If previewTitle IsNot Nothing Then
            previewTitle.Bounds = New Rectangle(contentLeft, 38, contentWidth, 44)
        End If

        Dim previewMeta = FindTemplateCardChild(coverHost, "TemplatePreviewMeta")
        If previewMeta IsNot Nothing Then
            previewMeta.Bounds = New Rectangle(contentLeft, 92, contentWidth, 20)
        End If

        Dim cover = FindTemplateCardChild(coverHost, "TemplateCoverImage")
        If cover IsNot Nothing Then
            cover.Bounds = New Rectangle(0, 0, coverHost.ClientSize.Width, coverHost.ClientSize.Height)
            cover.BringToFront()
        End If

        Dim coverStatus = FindTemplateCardChild(coverHost, "TemplateCoverStatus")
        If coverStatus IsNot Nothing Then
            coverStatus.Bounds = New Rectangle(contentLeft, Math.Max(8, coverHost.ClientSize.Height - 34), contentWidth, 24)
            If Not cover.Visible Then coverStatus.BringToFront()
        End If
    End Sub

    Private Function FindTemplateCardChild(parent As Control, childName As String) As Control
        If parent Is Nothing OrElse String.IsNullOrWhiteSpace(childName) Then Return Nothing

        For Each child As Control In parent.Controls
            If String.Equals(child.Name, childName, StringComparison.Ordinal) Then Return child
        Next

        Return Nothing
    End Function

    Private Function BuildTemplateCoverUrl(coverUrl As String) As String
        If String.IsNullOrWhiteSpace(coverUrl) Then Return coverUrl

        Dim trimmedUrl = coverUrl.Trim()
        Dim separator = If(trimmedUrl.Contains("?"), "&", "?")
        Return trimmedUrl & separator & "token=" & Uri.EscapeDataString(DocmeePptClient.GetConfiguredToken())
    End Function

    Private Function CreateTemplatePreviewBitmap(template As DocmeeTemplateInfo, statusText As String) As Bitmap
        Dim width = 640
        Dim height = 360
        Dim bitmap As New Bitmap(width, height)

        Using g = Graphics.FromImage(bitmap)
            g.Clear(Color.FromArgb(255, 248, 241))
            Using borderPen As New Pen(Color.FromArgb(226, 232, 240), 6.0F)
                g.DrawRectangle(borderPen, 3, 3, width - 6, height - 6)
            End Using

            Using accentBrush As New SolidBrush(Color.FromArgb(234, 88, 12))
                g.FillRectangle(accentBrush, 0, 0, width, 18)
                g.FillRectangle(accentBrush, 0, height - 18, width, 18)
            End Using

            Dim title = If(template Is Nothing OrElse String.IsNullOrWhiteSpace(template.Name), "模板预览", template.Name.Trim())
            Dim meta = BuildTemplateMetaText(template)
            Dim status = If(String.IsNullOrWhiteSpace(statusText), "封面加载中...", statusText.Trim())

            Using titleFont As New Font(Me.Font.FontFamily, 42.0F, FontStyle.Bold),
                  metaFont As New Font(Me.Font.FontFamily, 24.0F, FontStyle.Regular),
                  statusFont As New Font(Me.Font.FontFamily, 22.0F, FontStyle.Regular),
                  titleBrush As New SolidBrush(Color.FromArgb(39, 45, 55)),
                  metaBrush As New SolidBrush(Color.FromArgb(86, 94, 108)),
                  statusBrush As New SolidBrush(Color.FromArgb(185, 28, 28))

                Dim format As New StringFormat()
                format.Trimming = StringTrimming.EllipsisWord
                format.FormatFlags = StringFormatFlags.LineLimit

                g.DrawString(title, titleFont, titleBrush, New RectangleF(48, 90, width - 96, 150), format)
                g.DrawString(meta, metaFont, metaBrush, New RectangleF(48, 265, width - 96, 58), format)
                g.DrawString(status, statusFont, statusBrush, New RectangleF(48, 385, width - 96, 72), format)
            End Using
        End Using

        Return bitmap
    End Function

    Private Function BuildTemplateMetaText(template As DocmeeTemplateInfo) As String
        Dim parts As New List(Of String)()
        If template IsNot Nothing Then
            If Not String.IsNullOrWhiteSpace(template.Category) Then parts.Add(template.Category.Trim())
            If Not String.IsNullOrWhiteSpace(template.Style) Then parts.Add(template.Style.Trim())
        End If

        If parts.Count = 0 Then Return "Docmee 模板"
        Return String.Join(" / ", parts)
    End Function

    Private Sub AddTemplateCardClickHandlers(control As Control, template As DocmeeTemplateInfo)
        AddHandler control.Click, Sub(sender, args) SelectTemplate(template)

        For Each child As Control In control.Controls
            AddTemplateCardClickHandlers(child, template)
        Next
    End Sub

    Private Sub SelectTemplate(template As DocmeeTemplateInfo)
        If template Is Nothing Then Return

        Dim selectedIndex = -1
        For index As Integer = 0 To _templateCombo.Items.Count - 1
            Dim item = TryCast(_templateCombo.Items(index), DocmeeTemplateInfo)
            If item IsNot Nothing AndAlso String.Equals(item.Id, template.Id, StringComparison.Ordinal) Then
                selectedIndex = index
                Exit For
            End If
        Next

        If selectedIndex < 0 Then
            _templateCombo.Items.Add(template)
            selectedIndex = _templateCombo.Items.Count - 1
        End If

        If _templateCombo.SelectedIndex <> selectedIndex Then
            _templateCombo.SelectedIndex = selectedIndex
        End If

        RefreshTemplateSelectionStyles()
    End Sub

    Private Sub TemplateCombo_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim selectedTemplate = TryCast(_templateCombo.SelectedItem, DocmeeTemplateInfo)
        If selectedTemplate Is Nothing OrElse
           Not String.Equals(selectedTemplate.Id, _confirmedTemplateId, StringComparison.Ordinal) Then
            _templateConfirmedForCurrentOutline = False
        End If

        RefreshTemplateSelectionStyles()
        RefreshActionButtons()
    End Sub

    Private Sub TemplateListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim selectedTemplate = TryCast(_templateListBox.SelectedItem, DocmeeTemplateInfo)
        If selectedTemplate Is Nothing Then Return

        For index As Integer = 0 To _templateCombo.Items.Count - 1
            Dim item = TryCast(_templateCombo.Items(index), DocmeeTemplateInfo)
            If item IsNot Nothing AndAlso String.Equals(item.Id, selectedTemplate.Id, StringComparison.Ordinal) Then
                If _templateCombo.SelectedIndex <> index Then _templateCombo.SelectedIndex = index
                Exit For
            End If
        Next

        _templateListBox.Invalidate()
    End Sub

    Private Sub TemplateListBox_DrawItem(sender As Object, e As DrawItemEventArgs)
        If e.Index < 0 OrElse e.Index >= _templateListBox.Items.Count Then Return

        Dim template = TryCast(_templateListBox.Items(e.Index), DocmeeTemplateInfo)
        If template Is Nothing Then Return

        Dim isSelected = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        Dim cardBounds = Rectangle.Inflate(e.Bounds, -6, -5)
        Dim imageBounds = New Rectangle(cardBounds.Left + 10, cardBounds.Top + 10, Math.Min(240, Math.Max(160, cardBounds.Width \ 3)), cardBounds.Height - 20)
        Dim textLeft = imageBounds.Right + 12
        Dim textBounds = New Rectangle(textLeft, cardBounds.Top + 12, Math.Max(120, cardBounds.Right - textLeft - 12), cardBounds.Height - 24)

        Using backgroundBrush As New SolidBrush(If(isSelected, Color.FromArgb(255, 245, 235), Color.White))
            e.Graphics.FillRectangle(backgroundBrush, e.Bounds)
            e.Graphics.FillRectangle(backgroundBrush, cardBounds)
        End Using

        Using borderPen As New Pen(If(isSelected, Color.FromArgb(234, 88, 12), Color.FromArgb(203, 213, 225)), If(isSelected, 2.0F, 1.0F))
            e.Graphics.DrawRectangle(borderPen, cardBounds)
        End Using

        DrawTemplateListPreview(e.Graphics, template, imageBounds)

        Dim title = If(String.IsNullOrWhiteSpace(template.Name), template.Id, template.Name)
        Dim meta = BuildTemplateMetaText(template)
        Dim status = If(_templateCoverMessages.ContainsKey(template.Id), _templateCoverMessages(template.Id), "")

        Using titleFont As New Font(Me.Font.FontFamily, 10.0F, FontStyle.Bold),
              metaFont As New Font(Me.Font.FontFamily, 8.5F, FontStyle.Regular),
              statusFont As New Font(Me.Font.FontFamily, 8.0F, FontStyle.Regular),
              titleBrush As New SolidBrush(Color.FromArgb(39, 45, 55)),
              metaBrush As New SolidBrush(Color.FromArgb(86, 94, 108)),
              statusBrush As New SolidBrush(If(status.StartsWith("封面加载失败", StringComparison.Ordinal), Color.FromArgb(185, 28, 28), Color.FromArgb(86, 94, 108)))

            Dim format As New StringFormat()
            format.Trimming = StringTrimming.EllipsisWord
            format.FormatFlags = StringFormatFlags.LineLimit

            e.Graphics.DrawString(title, titleFont, titleBrush, New RectangleF(textBounds.Left, textBounds.Top, textBounds.Width, 42), format)
            e.Graphics.DrawString(meta, metaFont, metaBrush, New RectangleF(textBounds.Left, textBounds.Top + 48, textBounds.Width, 22), format)
            If Not String.IsNullOrWhiteSpace(status) Then
                e.Graphics.DrawString(status, statusFont, statusBrush, New RectangleF(textBounds.Left, textBounds.Top + 78, textBounds.Width, 38), format)
            End If
        End Using

        Dim buttonBounds = New Rectangle(textBounds.Left, cardBounds.Bottom - 38, Math.Min(120, textBounds.Width), 28)
        Using buttonBrush As New SolidBrush(If(isSelected, Color.FromArgb(234, 88, 12), Color.FromArgb(241, 245, 249))),
              buttonTextBrush As New SolidBrush(If(isSelected, Color.White, Color.FromArgb(39, 45, 55))),
              buttonFont As New Font(Me.Font.FontFamily, 8.5F, FontStyle.Bold)
            e.Graphics.FillRectangle(buttonBrush, buttonBounds)
            TextRenderer.DrawText(e.Graphics, If(isSelected, "已选择", "选择模板"), buttonFont, buttonBounds, If(isSelected, Color.White, Color.FromArgb(39, 45, 55)), TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
        End Using
    End Sub

    Private Sub DrawTemplateListPreview(graphics As Graphics, template As DocmeeTemplateInfo, bounds As Rectangle)
        Using backgroundBrush As New SolidBrush(Color.FromArgb(255, 248, 241)),
              borderPen As New Pen(Color.FromArgb(226, 232, 240))
            graphics.FillRectangle(backgroundBrush, bounds)
            graphics.DrawRectangle(borderPen, bounds)
        End Using

        If template IsNot Nothing AndAlso _templatePreviewImages.ContainsKey(template.Id) Then
            DrawImageZoom(graphics, _templatePreviewImages(template.Id), bounds)
            Return
        End If

        Dim title = If(template Is Nothing OrElse String.IsNullOrWhiteSpace(template.Name), "模板预览", template.Name.Trim())
        Dim status = If(template IsNot Nothing AndAlso _templateCoverMessages.ContainsKey(template.Id), _templateCoverMessages(template.Id), "封面加载中...")

        Using accentBrush As New SolidBrush(Color.FromArgb(234, 88, 12)),
              titleFont As New Font(Me.Font.FontFamily, 9.0F, FontStyle.Bold),
              statusFont As New Font(Me.Font.FontFamily, 8.0F, FontStyle.Regular),
              titleBrush As New SolidBrush(Color.FromArgb(39, 45, 55)),
              statusBrush As New SolidBrush(If(status.StartsWith("封面加载失败", StringComparison.Ordinal), Color.FromArgb(185, 28, 28), Color.FromArgb(86, 94, 108)))

            graphics.FillRectangle(accentBrush, bounds.Left, bounds.Top, 5, bounds.Height)
            Dim inner = Rectangle.Inflate(bounds, -12, -10)
            TextRenderer.DrawText(graphics, title, titleFont, New Rectangle(inner.Left, inner.Top + 18, inner.Width, 42), Color.FromArgb(39, 45, 55), TextFormatFlags.WordBreak Or TextFormatFlags.EndEllipsis)
            TextRenderer.DrawText(graphics, status, statusFont, New Rectangle(inner.Left, inner.Bottom - 44, inner.Width, 40), If(status.StartsWith("封面加载失败", StringComparison.Ordinal), Color.FromArgb(185, 28, 28), Color.FromArgb(86, 94, 108)), TextFormatFlags.WordBreak Or TextFormatFlags.EndEllipsis)
        End Using
    End Sub

    Private Sub DrawImageZoom(graphics As Graphics, image As System.Drawing.Image, bounds As Rectangle)
        If image Is Nothing OrElse bounds.Width <= 0 OrElse bounds.Height <= 0 Then Return

        Dim scale = Math.Min(bounds.Width / CDbl(image.Width), bounds.Height / CDbl(image.Height))
        Dim width = CInt(image.Width * scale)
        Dim height = CInt(image.Height * scale)
        Dim left = bounds.Left + (bounds.Width - width) \ 2
        Dim top = bounds.Top + (bounds.Height - height) \ 2
        graphics.DrawImage(image, New Rectangle(left, top, width, height))
    End Sub

    Private Function CreateScaledTemplateImage(source As System.Drawing.Image, maxWidth As Integer, maxHeight As Integer) As System.Drawing.Image
        If source Is Nothing Then Return Nothing

        Dim scale = Math.Min(maxWidth / CDbl(source.Width), maxHeight / CDbl(source.Height))
        scale = Math.Min(1.0R, scale)
        Dim width = Math.Max(1, CInt(Math.Round(source.Width * scale)))
        Dim height = Math.Max(1, CInt(Math.Round(source.Height * scale)))
        Dim bitmap As New Bitmap(width, height)

        Using g As Graphics = Graphics.FromImage(bitmap)
            g.Clear(Color.White)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.PixelOffsetMode = PixelOffsetMode.HighQuality
            g.DrawImage(source, New Rectangle(0, 0, width, height))
        End Using

        Return bitmap
    End Function

    Private Sub RefreshTemplateSelectionStyles()
        Dim selectedTemplate = TryCast(_templateCombo.SelectedItem, DocmeeTemplateInfo)
        Dim selectedId = If(selectedTemplate Is Nothing, "", selectedTemplate.Id)

        For Each pair In _templateCards
            Dim isSelected = String.Equals(pair.Key, selectedId, StringComparison.Ordinal)
            pair.Value.BackColor = If(isSelected, Color.FromArgb(255, 245, 235), Color.White)
            pair.Value.Padding = If(isSelected, New Padding(5), New Padding(8))
            LayoutTemplateCard(pair.Value)

            If _templateSelectLabels.ContainsKey(pair.Key) Then
                Dim selectLabel = _templateSelectLabels(pair.Key)
                selectLabel.Text = If(isSelected, "已选择", "选择模板")
                selectLabel.BackColor = If(isSelected, Color.FromArgb(234, 88, 12), Color.FromArgb(241, 245, 249))
                selectLabel.ForeColor = If(isSelected, Color.White, Color.FromArgb(39, 45, 55))
            End If
        Next
    End Sub

    Private Sub TemplateCardPanel_Resize(sender As Object, e As EventArgs)
        ResizeTemplateCards()
    End Sub

    Private Sub ResizeTemplateCards()
        Dim cardWidth = Math.Max(220, _templateCardPanel.ClientSize.Width - 24)
        For Each card As Control In _templateCardPanel.Controls
            card.Width = cardWidth
        Next
    End Sub

    Private Async Function GenerateAndImportPptxAsync() As Task
        Dim markdown = GetEditedMarkdown()
        If String.IsNullOrWhiteSpace(markdown) Then
            MessageBox.Show("请先生成或填写大纲。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            ValidateEditedMarkdownForDocmee(markdown)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Markdown 大纲", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End Try

        Dim selectedTemplate = TryCast(_templateCombo.SelectedItem, DocmeeTemplateInfo)
        If selectedTemplate Is Nothing OrElse String.IsNullOrWhiteSpace(selectedTemplate.Id) Then
            MessageBox.Show("请先选择模板。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        _generateButton.Enabled = False
        _insertButton.Enabled = False
        _finishOutlineEditButton.Enabled = False
        _refreshTemplatesButton.Enabled = False
        _selectTemplateButton.Enabled = False
        ShowOutlineOutput()
        ClearGeneratedPptState()

        Try
            _outlineMarkdown = markdown
            _templateConfirmedForCurrentOutline = True
            _confirmedTemplateId = selectedTemplate.Id

            SetStatus("正在创建模板生成任务...")
            AppendTaskPaneLine("使用模板ID: " & selectedTemplate.Id)
            Dim pptTaskId = Await _client.CreateMarkdownTaskAsync(markdown)
            AppendTaskPaneLine("生成任务ID: " & pptTaskId)

            SetStatus("正在按所选模板生成 PPTX...")
            Dim pptInfo = Await _client.GeneratePptxAsync(pptTaskId, selectedTemplate.Id, markdown)
            AppendTaskPaneLine("PPT ID: " & pptInfo.Id)
            AppendTaskPaneLine("返回模板ID: " & pptInfo.TemplateId)
            If Not String.IsNullOrWhiteSpace(pptInfo.CoverUrl) Then
                AppendTaskPaneLine("PPT 封面预览: " & pptInfo.CoverUrl)
            End If

            If Not String.Equals(pptInfo.TemplateId, selectedTemplate.Id, StringComparison.Ordinal) Then
                Throw New InvalidOperationException($"Docmee 返回的模板ID与所选模板不一致。所选: {selectedTemplate.Id}，返回: {pptInfo.TemplateId}")
            End If
            _lastGeneratedPptId = pptInfo.Id
            SaveDocmeePptxIdToCurrentPresentation(_lastGeneratedPptId)

            SetStatus("正在获取 PPTX 下载地址...")
            Dim fileUrl = Await _client.DownloadPptxAsync(pptInfo.Id, True)
            AppendTaskPaneLine("PPTX 下载地址: " & fileUrl)

            SetStatus("正在下载 PPTX...")
            Dim localPath = Path.Combine(Path.GetTempPath(), $"wenduoduoAI_{pptInfo.Id}.pptx")
            AppendTaskPaneLine("本地保存路径: " & localPath)
            Await _client.DownloadPptxFileAsync(fileUrl, localPath)

            SetStatus("正在导入当前演示文稿...")
            Dim importedCount = ImportPptxIntoPresentation(localPath)
            CaptureImportedSlideRange(importedCount)
            AppendTaskPaneLine("已导入页数: " & importedCount.ToString())
            SetStatus($"已生成并导入当前演示文稿，共 {importedCount} 页。")
        Catch ex As Exception
            SetStatus("生成或导入失败。")
            AppendTaskPaneLine("生成并导入失败: " & ex.Message)
            ClearGeneratedPptState()
            MessageBox.Show("生成并导入 PPT 失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _generateButton.Enabled = True
            RefreshActionButtons()
        End Try
    End Function

    Private Sub SaveDocmeePptxIdToCurrentPresentation(pptxId As String)
        If String.IsNullOrWhiteSpace(pptxId) Then Return

        Try
            Dim presentation = GetOrCreatePresentation()
            If presentation Is Nothing Then Return

            Try
                presentation.Tags.Delete(DocmeePptxIdTagName)
            Catch
            End Try

            presentation.Tags.Add(DocmeePptxIdTagName, pptxId.Trim())
        Catch ex As Exception
            AppendThemePptLog("Save Docmee PPT ID failed: " & ex.Message)
        End Try
    End Sub

    Private Function ReplaceImportedSlideRange(localPath As String) As Integer
        Dim target = GetOrCreatePresentation()
        If _lastImportedSlideStartIndex <= 0 OrElse _lastImportedSlideCount <= 0 Then
            Dim appendedCount = ImportPptxIntoPresentation(localPath)
            CaptureImportedSlideRange(appendedCount)
            Return appendedCount
        End If

        Dim originalStart = Math.Max(1, Math.Min(_lastImportedSlideStartIndex, target.Slides.Count + 1))
        Dim originalCount = Math.Min(_lastImportedSlideCount, Math.Max(0, target.Slides.Count - originalStart + 1))
        If originalCount <= 0 Then
            Dim appendedCount = ImportPptxIntoPresentation(localPath)
            CaptureImportedSlideRange(appendedCount)
            Return appendedCount
        End If

        Dim insertedCount = ImportPptxIntoPresentation(localPath, originalStart - 1)
        For slideIndex As Integer = originalStart + insertedCount + originalCount - 1 To originalStart + insertedCount Step -1
            If slideIndex > 0 AndAlso slideIndex <= target.Slides.Count Then
                target.Slides(slideIndex).Delete()
            End If
        Next

        _lastImportedSlideStartIndex = originalStart
        _lastImportedSlideCount = insertedCount
        RestoreActiveSlide(target, originalStart)
        Return insertedCount
    End Function

    Private Sub AppendTemplateLoadFailure(ex As Exception)
        If Me.IsDisposed OrElse _outputBox.IsDisposed Then Return
        If _outputBox.InvokeRequired Then
            BeginInvokeIfAlive(CType(Sub() AppendTemplateLoadFailure(ex), MethodInvoker))
            Return
        End If

        Dim builder As New StringBuilder()
        builder.AppendLine()
        builder.AppendLine("模板接口加载失败，已尝试使用内置演示模板。")

        Dim current = ex
        Dim depth = 1
        Do While current IsNot Nothing
            builder.AppendLine($"错误 {depth}: {current.GetType().Name}: {current.Message}")
            current = current.InnerException
            depth += 1
        Loop

        If _outputBox.TextLength > 0 AndAlso Not _outputBox.Text.EndsWith(vbCrLf) Then
            _outputBox.AppendText(vbCrLf)
        End If

        _outputBox.AppendText(builder.ToString())
        _outputBox.SelectionStart = _outputBox.TextLength
        _outputBox.ScrollToCaret()
    End Sub

    Private Sub AppendTaskPaneLine(text As String)
        If Me.IsDisposed OrElse _outputBox.IsDisposed Then Return
        If _outputBox.InvokeRequired Then
            BeginInvokeIfAlive(CType(Sub() AppendTaskPaneLine(text), MethodInvoker))
            Return
        End If

        If _outputBox.TextLength > 0 AndAlso Not _outputBox.Text.EndsWith(vbCrLf) Then
            _outputBox.AppendText(vbCrLf)
        End If

        _outputBox.AppendText(vbCrLf & text & vbCrLf)
        _outputBox.SelectionStart = _outputBox.TextLength
        _outputBox.ScrollToCaret()
    End Sub

    Public Sub InsertOutlineIntoPresentation(outline As JObject)
        If outline Is Nothing Then
            Throw New InvalidOperationException("请先生成大纲。")
        End If

        Dim presentation = GetOrCreatePresentation()
        Dim root = GetContentRoot(outline)
        Dim presentationTitle = GetNodeText(root, "主题生成PPT")
        Dim sections = GetChildren(root)

        CreateCoverSlide(presentation, presentationTitle)

        If sections.Count = 0 Then
            Dim fallbackLines = CollectBodyLines(root, 0)
            If fallbackLines.Count = 0 Then fallbackLines.Add("- " & presentationTitle)
            CreateContentSlide(presentation, "内容大纲", fallbackLines)
            Return
        End If

        For Each child In sections
            Dim childObj = TryCast(child, JObject)
            If childObj Is Nothing Then Continue For
            If IsDocmeeCoverPage(childObj) AndAlso sections.Count > 1 Then Continue For

            Dim slideTitle = GetNodeText(childObj, "内容页")
            Dim bodyLines = CollectBodyLines(childObj, 0)
            If bodyLines.Count = 0 Then bodyLines.Add("- " & slideTitle)
            CreateContentSlide(presentation, slideTitle, bodyLines)
        Next
    End Sub

    Private Function ImportPptxIntoPresentation(downloadPath As String, Optional insertAfterIndex As Integer = -1) As Integer
        If String.IsNullOrWhiteSpace(downloadPath) OrElse Not File.Exists(downloadPath) Then
            Throw New FileNotFoundException("未找到下载后的 PPTX 文件。", downloadPath)
        End If

        Dim target = GetOrCreatePresentation()
        Dim originalSlideIndex = CaptureActiveSlideIndex(target)
        Dim importedSlides As New List(Of PowerPoint.Slide)()
        Dim normalizedInsertAfterIndex = NormalizeInsertAfterIndex(target, insertAfterIndex)

        Try
            importedSlides = ImportPptxFileIntoPresentation(target, downloadPath, normalizedInsertAfterIndex)
            If importedSlides.Count > 0 Then
                AppendTaskPaneLine("导入方式: InsertFromFile")
            Else
                AppendTaskPaneLine("InsertFromFile 未导入幻灯片，尝试复制粘贴。")
                importedSlides = CopyPptxSlidesIntoPresentation(target, downloadPath, normalizedInsertAfterIndex)
            End If
        Catch ex As Exception
            AppendTaskPaneLine("InsertFromFile 导入失败，尝试复制粘贴: " & ex.Message)
            importedSlides = CopyPptxSlidesIntoPresentation(target, downloadPath, normalizedInsertAfterIndex)
        Finally
            RestoreActiveSlide(target, originalSlideIndex)
        End Try

        If importedSlides.Count = 0 Then
            Throw New InvalidOperationException("PPTX 已下载，但没有成功导入任何幻灯片。")
        End If

        FixInsertedSlideReadability(importedSlides)
        RestoreImportedSlideBackgrounds(downloadPath, importedSlides)
        Return importedSlides.Count
    End Function

    Private Function NormalizeInsertAfterIndex(target As PowerPoint.Presentation, insertAfterIndex As Integer) As Integer
        If target Is Nothing Then Return 0
        If insertAfterIndex < 0 Then Return target.Slides.Count
        Return Math.Max(0, Math.Min(insertAfterIndex, target.Slides.Count))
    End Function

    Private Function ImportPptxFileIntoPresentation(target As PowerPoint.Presentation, downloadPath As String, insertAfterIndex As Integer) As List(Of PowerPoint.Slide)
        Dim insertedCount = target.Slides.InsertFromFile(downloadPath, insertAfterIndex)
        Dim importedSlides As New List(Of PowerPoint.Slide)()

        If insertedCount <= 0 Then Return importedSlides

        For slideIndex As Integer = insertAfterIndex + 1 To insertAfterIndex + insertedCount
            If slideIndex > 0 AndAlso slideIndex <= target.Slides.Count Then
                importedSlides.Add(target.Slides(slideIndex))
            End If
        Next

        Return importedSlides
    End Function

    Private Function CopyPptxSlidesIntoPresentation(target As PowerPoint.Presentation, downloadPath As String, insertAfterIndex As Integer) As List(Of PowerPoint.Slide)
        Dim sourcePresentation As PowerPoint.Presentation = Nothing
        Dim importedSlides As New List(Of PowerPoint.Slide)()

        Try
            sourcePresentation = _pptApp.Presentations.Open(downloadPath,
                ReadOnly:=MsoTriState.msoTrue,
                Untitled:=MsoTriState.msoFalse,
                WithWindow:=MsoTriState.msoFalse)

            For slideIndex As Integer = 1 To sourcePresentation.Slides.Count
                Dim pastedSlide = TryPasteSlideWithSourceFormatting(target, sourcePresentation.Slides(slideIndex), insertAfterIndex + importedSlides.Count)
                If pastedSlide IsNot Nothing Then importedSlides.Add(pastedSlide)
            Next
        Finally
            If sourcePresentation IsNot Nothing Then
                sourcePresentation.Close()
            End If
        End Try

        Return importedSlides
    End Function

    Private Sub CaptureImportedSlideRange(importedCount As Integer)
        If importedCount <= 0 Then
            _lastImportedSlideStartIndex = 0
            _lastImportedSlideCount = 0
            Return
        End If

        Dim target = GetOrCreatePresentation()
        _lastImportedSlideCount = importedCount
        _lastImportedSlideStartIndex = Math.Max(1, target.Slides.Count - importedCount + 1)
    End Sub

    Private Function CaptureActiveSlideIndex(target As PowerPoint.Presentation) As Integer
        Try
            If target Is Nothing OrElse target.Slides.Count = 0 OrElse
               _pptApp.ActiveWindow Is Nothing OrElse _pptApp.ActiveWindow.Selection Is Nothing Then
                Return 0
            End If

            Dim selection = _pptApp.ActiveWindow.Selection
            If selection.SlideRange IsNot Nothing AndAlso selection.SlideRange.Count > 0 Then
                Return selection.SlideRange(1).SlideIndex
            End If
        Catch
        End Try

        Return 0
    End Function

    Private Sub RestoreActiveSlide(target As PowerPoint.Presentation, slideIndex As Integer)
        If target Is Nothing OrElse slideIndex <= 0 OrElse slideIndex > target.Slides.Count Then Return

        Try
            If target.Windows.Count <= 0 Then Return

            target.Windows(1).Activate()
            target.Windows(1).View.GotoSlide(slideIndex)
            target.Slides(slideIndex).Select()
        Catch
            ' 恢复用户原来的编辑页失败不应阻断 PPT 导入。
        End Try
    End Sub

    Private Function TryPasteSlideWithSourceFormatting(target As PowerPoint.Presentation, sourceSlide As PowerPoint.Slide, insertAfterIndex As Integer) As PowerPoint.Slide
        Dim beforeCount = target.Slides.Count
        Dim destinationIndex = NormalizeInsertAfterIndex(target, insertAfterIndex)

        Try
            sourceSlide.Copy()
            ActivateDestinationAtPosition(target, destinationIndex)
            _pptApp.CommandBars.ExecuteMso("PasteSourceFormatting")

            If WaitForPastedSlide(target, beforeCount) Then
                Dim pastedSlide = GetSelectedSlide(target)
                If pastedSlide IsNot Nothing Then
                    Dim moveToIndex = Math.Min(destinationIndex + 1, target.Slides.Count)
                    pastedSlide.MoveTo(moveToIndex)
                    Return target.Slides(moveToIndex)
                End If

                Return target.Slides(Math.Min(destinationIndex + 1, target.Slides.Count))
            End If
        Catch
        End Try

        sourceSlide.Copy()
        Dim pastedSlides = target.Slides.Paste(Math.Min(destinationIndex + 1, target.Slides.Count + 1))
        If pastedSlides IsNot Nothing AndAlso pastedSlides.Count > 0 Then Return pastedSlides(1)
        Return Nothing
    End Function

    Private Sub ActivateDestinationAtEnd(target As PowerPoint.Presentation)
        If target.Windows.Count <= 0 Then Return

        target.Windows(1).Activate()
        If target.Slides.Count > 0 Then
            target.Windows(1).View.GotoSlide(target.Slides.Count)
        End If
    End Sub

    Private Sub ActivateDestinationAtPosition(target As PowerPoint.Presentation, insertAfterIndex As Integer)
        If target.Windows.Count <= 0 Then Return

        target.Windows(1).Activate()
        If target.Slides.Count <= 0 Then Return

        Dim gotoIndex = Math.Max(1, Math.Min(insertAfterIndex, target.Slides.Count))
        target.Windows(1).View.GotoSlide(gotoIndex)
    End Sub

    Private Function WaitForPastedSlide(target As PowerPoint.Presentation, beforeCount As Integer) As Boolean
        For attempt As Integer = 1 To 10
            If target.Slides.Count > beforeCount Then Return True
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(50)
        Next

        Return target.Slides.Count > beforeCount
    End Function

    Private Function GetSelectedSlide(target As PowerPoint.Presentation) As PowerPoint.Slide
        Try
            Dim selection = _pptApp.ActiveWindow.Selection
            If selection IsNot Nothing AndAlso
               selection.SlideRange IsNot Nothing AndAlso
               selection.SlideRange.Count > 0 Then
                Return selection.SlideRange(1)
            End If
        Catch
        End Try

        Return Nothing
    End Function

    Private Sub FixInsertedSlideReadability(importedSlides As List(Of PowerPoint.Slide))
        For Each slide As PowerPoint.Slide In importedSlides
            Dim slideBackgroundLuminance = GetSlideBackgroundLuminance(slide)
            If slideBackgroundLuminance < 0 Then Continue For

            For Each shape As PowerPoint.Shape In slide.Shapes
                FixShapeTextReadability(shape, slideBackgroundLuminance)
            Next
        Next
    End Sub

    Private Sub FixShapeTextReadability(shape As PowerPoint.Shape, slideBackgroundLuminance As Double)
        If shape Is Nothing Then Return

        Try
            If shape.Type = MsoShapeType.msoGroup Then
                For itemIndex As Integer = 1 To shape.GroupItems.Count
                    FixShapeTextReadability(shape.GroupItems(itemIndex), slideBackgroundLuminance)
                Next
                Return
            End If

            If shape.HasTable = MsoTriState.msoTrue Then
                For rowIndex As Integer = 1 To shape.Table.Rows.Count
                    For columnIndex As Integer = 1 To shape.Table.Columns.Count
                        FixShapeTextReadability(shape.Table.Cell(rowIndex, columnIndex).Shape, slideBackgroundLuminance)
                    Next
                Next
            End If

            If shape.HasTextFrame <> MsoTriState.msoTrue OrElse
               shape.TextFrame.HasText <> MsoTriState.msoTrue Then
                Return
            End If

            Dim textRange = shape.TextFrame.TextRange
            Dim textLuminance = GetColorLuminance(textRange.Font.Color.RGB)
            Dim backgroundLuminance = GetShapeBackgroundLuminance(shape, slideBackgroundLuminance)
            Dim contrastGap = Math.Abs(backgroundLuminance - textLuminance)

            If backgroundLuminance >= 180 AndAlso textLuminance >= 165 AndAlso contrastGap < 95 Then
                textRange.Font.Color.RGB = RGB(45, 52, 64)
            ElseIf backgroundLuminance <= 95 AndAlso textLuminance <= 110 AndAlso contrastGap < 80 Then
                textRange.Font.Color.RGB = RGB(255, 255, 255)
            End If
        Catch
            ' 单个形状格式读取失败时，不阻断 PPT 导入。
        End Try
    End Sub

    Private Function GetSlideBackgroundLuminance(slide As PowerPoint.Slide) As Double
        Try
            Dim fill = slide.Background.Fill
            If fill IsNot Nothing AndAlso fill.Type = MsoFillType.msoFillSolid Then
                Return GetColorLuminance(fill.ForeColor.RGB)
            End If
        Catch
        End Try
        Return -1.0R
    End Function

    Private Function GetShapeBackgroundLuminance(shape As PowerPoint.Shape, fallbackLuminance As Double) As Double
        Try
            If shape.Fill.Visible = MsoTriState.msoTrue Then
                Return GetColorLuminance(shape.Fill.ForeColor.RGB)
            End If
        Catch
        End Try

        Return fallbackLuminance
    End Function

    Private Function GetColorLuminance(rgbValue As Integer) As Double
        Dim red = rgbValue And &HFF
        Dim green = (rgbValue >> 8) And &HFF
        Dim blue = (rgbValue >> 16) And &HFF

        Return (0.299R * red) + (0.587R * green) + (0.114R * blue)
    End Function

    ''' <summary>从源PPTX复制背景到已导入的幻灯片</summary>
    Private Sub RestoreImportedSlideBackgrounds(sourcePath As String, importedSlides As List(Of PowerPoint.Slide))
        If importedSlides.Count = 0 Then Return
        If String.IsNullOrWhiteSpace(sourcePath) OrElse Not File.Exists(sourcePath) Then Return

        Dim sourcePres As PowerPoint.Presentation = Nothing
        Try
            sourcePres = _pptApp.Presentations.Open(sourcePath,
                ReadOnly:=MsoTriState.msoTrue,
                Untitled:=MsoTriState.msoFalse,
                WithWindow:=MsoTriState.msoFalse)

            Dim maxCount = Math.Min(importedSlides.Count, sourcePres.Slides.Count)
            For i As Integer = 0 To maxCount - 1
                Try
                    CopySlideBackground(sourcePres.Slides(i + 1), importedSlides(i))
                Catch
                End Try
            Next
        Catch ex As Exception
            AppendThemePptLog("Restore backgrounds failed: " & ex.Message)
        Finally
            If sourcePres IsNot Nothing Then
                sourcePres.Close()
            End If
        End Try
    End Sub

    ''' <summary>将源幻灯片的背景复制到目标幻灯片</summary>
    Private Sub CopySlideBackground(source As PowerPoint.Slide, dest As PowerPoint.Slide)
        If source Is Nothing OrElse dest Is Nothing Then Return

        Try
            ' 如果源幻灯片使用母版背景，目标也已继承，无需处理
            If source.FollowMasterBackground = MsoTriState.msoTrue Then Return

            dest.FollowMasterBackground = MsoTriState.msoFalse

            Dim srcFill = source.Background.Fill
            Dim dstFill = dest.Background.Fill

            Select Case srcFill.Type
                Case MsoFillType.msoFillSolid
                    dstFill.ForeColor.RGB = srcFill.ForeColor.RGB
                    dstFill.BackColor.RGB = srcFill.BackColor.RGB
                    dstFill.Visible = MsoTriState.msoTrue

                Case MsoFillType.msoFillGradient
                    dstFill.ForeColor.RGB = srcFill.ForeColor.RGB
                    dstFill.BackColor.RGB = srcFill.BackColor.RGB
                    dstFill.Visible = MsoTriState.msoTrue

                Case Else
                    ' 图片/纹理/图案背景：导出源幻灯片无形状版本作为背景图
                    Dim tempSlide = source.Duplicate()(1)
                    Try
                        While tempSlide.Shapes.Count > 0
                            tempSlide.Shapes(1).Delete()
                        End While

                        Dim bgPath = Path.Combine(Path.GetTempPath(), $"pptbgb_{Guid.NewGuid():N}.png")
                        tempSlide.Export(bgPath, "PNG")

                        dstFill.UserPicture(bgPath)
                        dstFill.Visible = MsoTriState.msoTrue

                        Try
                            File.Delete(bgPath)
                        Catch
                        End Try
                    Finally
                        tempSlide.Delete()
                    End Try
            End Select
        Catch
        End Try
    End Sub

    Private Function GetOrCreatePresentation() As PowerPoint.Presentation
        Try
            If _pptApp.Presentations.Count > 0 Then
                Return _pptApp.ActivePresentation
            End If
        Catch
        End Try

        Return _pptApp.Presentations.Add(MsoTriState.msoTrue)
    End Function

    Private Sub CreateCoverSlide(presentation As PowerPoint.Presentation, titleText As String)
        Dim slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly)
        ApplySlideTheme(slide, presentation, True)
        SetTitleText(slide, titleText, 34.0F, True)
    End Sub

    Private Sub CreateContentSlide(presentation As PowerPoint.Presentation, titleText As String, bodyLines As List(Of String))
        Dim slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly)
        ApplySlideTheme(slide, presentation, False)
        SetTitleText(slide, titleText, 24.0F, False)

        Dim left = 54.0F
        Dim top = 108.0F
        Dim width = CSng(presentation.PageSetup.SlideWidth - left * 2)
        Dim height = CSng(presentation.PageSetup.SlideHeight - top - 46.0F)
        Dim bodyShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height)
        Dim bodyRange = bodyShape.TextFrame.TextRange
        bodyRange.Text = String.Join(vbCrLf, bodyLines)
        bodyRange.Font.Name = "微软雅黑"
        bodyRange.Font.Size = 17
        bodyRange.Font.Color.RGB = RGB(45, 52, 64)
        bodyShape.TextFrame.MarginLeft = 12
        bodyShape.TextFrame.MarginRight = 12
        bodyShape.TextFrame.MarginTop = 8
        bodyShape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText
    End Sub

    Private Sub ApplySlideTheme(slide As PowerPoint.Slide, presentation As PowerPoint.Presentation, isCover As Boolean)
        slide.FollowMasterBackground = MsoTriState.msoFalse
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = If(isCover, RGB(242, 247, 255), RGB(248, 250, 252))

        Dim accent = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 10, CSng(presentation.PageSetup.SlideHeight))
        accent.Fill.ForeColor.RGB = RGB(26, 115, 232)
        accent.Line.Visible = MsoTriState.msoFalse
    End Sub

    Private Sub SetTitleText(slide As PowerPoint.Slide, titleText As String, fontSize As Single, center As Boolean)
        Dim titleShape As PowerPoint.Shape = Nothing
        Try
            If slide.Shapes.HasTitle = MsoTriState.msoTrue Then
                titleShape = slide.Shapes.Title
            End If
        Catch
        End Try

        If titleShape Is Nothing Then
            titleShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 54, 38, 620, 54)
        End If

        Dim titleRange = titleShape.TextFrame.TextRange
        titleRange.Text = titleText
        titleRange.Font.Name = "微软雅黑"
        titleRange.Font.Size = fontSize
        titleRange.Font.Bold = MsoTriState.msoTrue
        titleRange.Font.Color.RGB = RGB(22, 32, 48)
        titleRange.ParagraphFormat.Alignment = If(center, PowerPoint.PpParagraphAlignment.ppAlignCenter, PowerPoint.PpParagraphAlignment.ppAlignLeft)

        If center Then
            titleShape.Left = 64
            titleShape.Top = 170
            titleShape.Width = 600
            titleShape.Height = 110
        Else
            titleShape.Left = 54
            titleShape.Top = 40
            titleShape.Width = 620
            titleShape.Height = 54
        End If
    End Sub

    Private Function BuildRequestContent(topic As String) As String
        If topic.Contains("PPT") OrElse topic.Contains("ppt") OrElse topic.StartsWith("请生成") Then
            Return topic
        End If

        Return $"请生成一份关于 {topic} 的详细PPT，每个章节包含3-5个要点，内容充实具体，包含数据、案例和实践建议。"
    End Function

    Private Function RenderOutlineText(outline As JObject) As String
        Dim root = GetContentRoot(outline)
        Dim builder As New StringBuilder()
        AppendOutlineLine(builder, root, 0)
        Return builder.ToString().Trim()
    End Function

    Private Function ConvertOutlineToMarkdown(outline As JObject) As String
        Dim root = GetContentRoot(outline)
        Dim title = CleanMarkdown(GetNodeText(root, "主题生成PPT"))
        Dim builder As New StringBuilder()
        builder.AppendLine("# " & title)
        builder.AppendLine()

        Dim sections = GetChildren(root)
        If sections.Count = 0 Then
            For Each line In CollectBodyLines(root, 0)
                builder.AppendLine(NormalizeMarkdownBullet(line))
            Next
            Return builder.ToString().Trim()
        End If

        For Each child In sections
            Dim childObj = TryCast(child, JObject)
            If childObj Is Nothing Then Continue For
            If IsDocmeeCoverPage(childObj) AndAlso sections.Count > 1 Then Continue For

            builder.AppendLine("## " & CleanMarkdown(GetNodeText(childObj, "内容页")))
            For Each line In CollectBodyLines(childObj, 0)
                builder.AppendLine(NormalizeMarkdownBullet(line))
            Next
            builder.AppendLine()
        Next

        Return builder.ToString().Trim()
    End Function

    Private Sub AppendOutlineLine(builder As StringBuilder, node As JObject, level As Integer)
        Dim text = GetNodeText(node, "")
        If Not String.IsNullOrWhiteSpace(text) Then
            builder.Append(New String(" "c, level * 2))
            builder.AppendLine("- " & text)
        End If

        For Each line In GetDirectContentLines(node, level + 1)
            builder.AppendLine(line)
        Next

        For Each child In GetChildren(node)
            Dim childObj = TryCast(child, JObject)
            If childObj IsNot Nothing Then AppendOutlineLine(builder, childObj, level + 1)
        Next
    End Sub

    Private Function CollectBodyLines(node As JObject, level As Integer) As List(Of String)
        Dim lines As New List(Of String)()
        lines.AddRange(GetDirectContentLines(node, level))

        For Each child In GetChildren(node)
            Dim childObj = TryCast(child, JObject)
            If childObj Is Nothing Then Continue For

            Dim text = GetNodeText(childObj, "")
            If Not String.IsNullOrWhiteSpace(text) Then
                lines.Add(New String(" "c, level * 2) & "- " & text)
            End If

            lines.AddRange(CollectBodyLines(childObj, level + 1))
        Next
        Return lines
    End Function

    Private Function GetContentRoot(outline As JObject) As JObject
        Dim children = GetChildren(outline)
        If Not HasNodeText(outline) AndAlso children.Count = 1 Then
            Dim onlyChild = TryCast(children(0), JObject)
            If onlyChild IsNot Nothing Then Return onlyChild
        End If

        Return outline
    End Function

    Private Function GetChildren(node As JObject) As JArray
        If node Is Nothing Then Return New JArray()

        Dim children = TryCast(node("children"), JArray)
        If children IsNot Nothing Then Return children

        children = TryCast(node("items"), JArray)
        If children IsNot Nothing Then Return children

        children = TryCast(node("subTitle"), JArray)
        If children IsNot Nothing Then Return children

        children = TryCast(node("pages"), JArray)
        If children IsNot Nothing Then Return children

        Return New JArray()
    End Function

    Private Function GetNodeText(node As JObject, fallback As String) As String
        If node Is Nothing Then Return fallback

        Dim candidates = {"overall_theme", "name", "title", "subtitle", "text"}
        For Each key In candidates
            Dim token = node(key)
            If token IsNot Nothing AndAlso token.Type = JTokenType.String Then
                Dim value = token.ToString().Trim()
                If Not String.IsNullOrWhiteSpace(value) Then Return value
            End If
        Next

        Return fallback
    End Function

    Private Function IsDocmeeCoverPage(node As JObject) As Boolean
        Dim pageType = If(node("page_type"), node("pageType"))
        If pageType Is Nothing OrElse pageType.Type <> JTokenType.String Then Return False
        Return String.Equals(pageType.ToString(), "cover", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function GetDirectContentLines(node As JObject, level As Integer) As List(Of String)
        Dim lines As New List(Of String)()
        If node Is Nothing Then Return lines

        AddTextLines(lines, node("subtitle"), level)
        AddTextLines(lines, node("text"), level)

        Dim content = node("content")
        If TypeOf content Is JObject Then
            Dim contentObj = DirectCast(content, JObject)
            AddTextLines(lines, contentObj("subtitle"), level)
            AddTextLines(lines, contentObj("text"), level)

            For Each prop As JProperty In contentObj.Properties()
                If prop.Name = "subtitle" OrElse prop.Name = "text" Then Continue For
                If prop.Value.Type = JTokenType.String Then
                    AddTextLines(lines, prop.Value, level)
                End If
            Next
        Else
            AddTextLines(lines, content, level)
        End If

        Return lines
    End Function

    Private Sub AddTextLines(lines As List(Of String), token As JToken, level As Integer)
        If token Is Nothing OrElse token.Type <> JTokenType.String Then Return

        Dim rawText = token.ToString().Trim()
        If String.IsNullOrWhiteSpace(rawText) Then Return

        Dim parts = rawText.Replace(vbCrLf, vbLf).Split({vbLf}, StringSplitOptions.None)
        For Each part In parts
            Dim value = part.Trim()
            If String.IsNullOrWhiteSpace(value) Then Continue For
            lines.Add(New String(" "c, level * 2) & "- " & value)
        Next
    End Sub

    Private Function NormalizeMarkdownBullet(line As String) As String
        If String.IsNullOrWhiteSpace(line) Then Return "- "

        Dim value = line.Trim()
        While value.StartsWith("-")
            value = value.Substring(1).Trim()
        End While

        Return "- " & CleanMarkdown(value)
    End Function

    Private Function CleanMarkdown(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return ""
        Return text.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
    End Function

    Private Function HasNodeText(node As JObject) As Boolean
        Return Not String.IsNullOrWhiteSpace(GetNodeText(node, ""))
    End Function

    Private Class TemplateGalleryPaintControl
        Inherits ScrollableControl

        Private Const OuterPadding As Integer = 6
        Private Const CardGap As Integer = 10
        Private Const CardPadding As Integer = 8
        Private Const CardExtraHeight As Integer = 104
        Private ReadOnly _templates As New List(Of DocmeeTemplateInfo)()
        Private ReadOnly _images As New Dictionary(Of String, Image)()
        Private ReadOnly _messages As New Dictionary(Of String, String)()
        Private _selectedId As String = ""
        Private _paintCount As Integer

        Public Event TemplateSelected(sender As Object, template As DocmeeTemplateInfo)

        Public Sub New()
            Me.AutoScroll = True
            Me.BackColor = Color.White
            Me.Cursor = Cursors.Default
            SetStyle(ControlStyles.UserPaint Or
                     ControlStyles.AllPaintingInWmPaint Or
                     ControlStyles.OptimizedDoubleBuffer Or
                     ControlStyles.ResizeRedraw Or
                     ControlStyles.Selectable, True)
            UpdateStyles()
        End Sub

        Public Sub SetData(templates As IEnumerable(Of DocmeeTemplateInfo),
                           images As IDictionary(Of String, Image),
                           messages As IDictionary(Of String, String),
                           selectedId As String)
            _templates.Clear()
            If templates IsNot Nothing Then
                For Each template In templates
                    If template IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(template.Id) Then
                        _templates.Add(template)
                    End If
                Next
            End If

            _images.Clear()
            If images IsNot Nothing Then
                For Each pair In images
                    If Not String.IsNullOrWhiteSpace(pair.Key) AndAlso pair.Value IsNot Nothing Then
                        _images(pair.Key) = pair.Value
                    End If
                Next
            End If

            _messages.Clear()
            If messages IsNot Nothing Then
                For Each pair In messages
                    If Not String.IsNullOrWhiteSpace(pair.Key) Then
                        _messages(pair.Key) = If(pair.Value, "")
                    End If
                Next
            End If

            _selectedId = If(selectedId, "")
            UpdateScrollRange()
            Invalidate()
        End Sub

        Public Sub RenderNow()
            ThemePptTaskPane.AppendThemePptLog("PaintGallery RenderNow skipped: direct rendering disabled.")
            Invalidate()
        End Sub

        Protected Overrides Sub OnCreateControl()
            MyBase.OnCreateControl()
            ThemePptTaskPane.AppendThemePptLog("PaintGallery created. handle=" & Me.Handle.ToString() & ", size=" & Me.Width.ToString() & "x" & Me.Height.ToString())
        End Sub

        Protected Overrides Sub OnResize(e As EventArgs)
            MyBase.OnResize(e)
            UpdateScrollRange()
            ThemePptTaskPane.AppendThemePptLog("PaintGallery resized: " & Me.Width.ToString() & "x" & Me.Height.ToString())
            Invalidate()
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            _paintCount += 1
            If _paintCount <= 5 OrElse _paintCount Mod 20 = 0 Then
                ThemePptTaskPane.AppendThemePptLog("PaintGallery OnPaint: paint=" & _paintCount.ToString() &
                                                  ", templates=" & _templates.Count.ToString() &
                                                  ", images=" & _images.Count.ToString() &
                                                  ", size=" & Me.Width.ToString() & "x" & Me.Height.ToString() &
                                                  ", scroll=" & Me.AutoScrollPosition.X.ToString() & "," & Me.AutoScrollPosition.Y.ToString())
            End If

            PaintGalleryContent(e.Graphics)
        End Sub

        Private Sub PaintGalleryContent(graphics As Graphics)
            graphics.Clear(Me.BackColor)
            graphics.SmoothingMode = SmoothingMode.AntiAlias
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality

            Dim state = graphics.Save()
            graphics.TranslateTransform(Me.AutoScrollPosition.X, Me.AutoScrollPosition.Y)
            Try
                If _templates.Count = 0 Then
                    DrawEmptyState(graphics)
                    Return
                End If

                For index As Integer = 0 To _templates.Count - 1
                    DrawCard(graphics, _templates(index), GetCardBounds(index))
                Next
            Finally
                graphics.Restore(state)
            End Try
        End Sub

        Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
            MyBase.OnMouseDown(e)
            If e.Button <> MouseButtons.Left Then Return

            Dim template = HitTestTemplate(e.Location)
            If template IsNot Nothing Then
                Me.Focus()
                RaiseEvent TemplateSelected(Me, template)
            End If
        End Sub

        Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
            MyBase.OnMouseMove(e)
            Me.Cursor = If(HitTestTemplate(e.Location) Is Nothing, Cursors.Default, Cursors.Hand)
        End Sub

        Private Sub UpdateScrollRange()
            Dim totalHeight = OuterPadding
            For index As Integer = 0 To _templates.Count - 1
                totalHeight += GetCardHeight() + CardGap
            Next

            If _templates.Count > 0 Then totalHeight += OuterPadding
            Me.AutoScrollMinSize = New Size(0, Math.Max(0, totalHeight))
        End Sub

        Private Function GetCardWidth() As Integer
            Dim scrollbarWidth = If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, SystemInformation.VerticalScrollBarWidth)
            Return Math.Max(220, Me.ClientSize.Width - scrollbarWidth - OuterPadding * 2)
        End Function

        Private Function GetCoverHeight() As Integer
            Dim coverWidth = Math.Max(180, GetCardWidth() - CardPadding * 2)
            Return Math.Max(120, Math.Min(360, CInt(Math.Round(coverWidth * 9.0R / 16.0R))))
        End Function

        Private Function GetCardHeight() As Integer
            Return GetCoverHeight() + CardExtraHeight
        End Function

        Private Function GetCardBounds(index As Integer) As Rectangle
            Dim top = OuterPadding + index * (GetCardHeight() + CardGap)
            Return New Rectangle(OuterPadding, top, GetCardWidth(), GetCardHeight())
        End Function

        Private Function HitTestTemplate(point As Point) As DocmeeTemplateInfo
            Dim contentPoint = New Point(point.X - Me.AutoScrollPosition.X, point.Y - Me.AutoScrollPosition.Y)
            For index As Integer = 0 To _templates.Count - 1
                If GetCardBounds(index).Contains(contentPoint) Then
                    Return _templates(index)
                End If
            Next

            Return Nothing
        End Function

        Private Sub DrawEmptyState(graphics As Graphics)
            Dim bounds = New Rectangle(OuterPadding, OuterPadding, Math.Max(180, GetCardWidth()), 140)
            Using backgroundBrush As New SolidBrush(Color.FromArgb(248, 250, 252)),
                  borderPen As New Pen(Color.FromArgb(203, 213, 225)),
                  titleFont As New Font(Me.Font.FontFamily, 10.0F, FontStyle.Bold),
                  textFont As New Font(Me.Font.FontFamily, 9.0F, FontStyle.Regular)
                graphics.FillRectangle(backgroundBrush, bounds)
                graphics.DrawRectangle(borderPen, bounds)
                TextRenderer.DrawText(graphics, "暂无模板", titleFont, New Rectangle(bounds.Left + 14, bounds.Top + 34, bounds.Width - 28, 26), Color.FromArgb(39, 45, 55), TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                TextRenderer.DrawText(graphics, "点击刷新后会在这里显示模板封面。", textFont, New Rectangle(bounds.Left + 14, bounds.Top + 68, bounds.Width - 28, 44), Color.FromArgb(86, 94, 108), TextFormatFlags.WordBreak Or TextFormatFlags.EndEllipsis)
            End Using
        End Sub

        Private Sub DrawCard(graphics As Graphics, template As DocmeeTemplateInfo, cardBounds As Rectangle)
            If template Is Nothing Then Return

            Dim isSelected = String.Equals(template.Id, _selectedId, StringComparison.Ordinal)
            Dim borderColor = If(isSelected, Color.FromArgb(234, 88, 12), Color.FromArgb(203, 213, 225))
            Dim backgroundColor = If(isSelected, Color.FromArgb(255, 247, 237), Color.White)

            Using backgroundBrush As New SolidBrush(backgroundColor),
                  borderPen As New Pen(borderColor, If(isSelected, 2.0F, 1.0F))
                graphics.FillRectangle(backgroundBrush, cardBounds)
                graphics.DrawRectangle(borderPen, cardBounds)
            End Using

            Dim coverBounds = New Rectangle(cardBounds.Left + CardPadding,
                                            cardBounds.Top + CardPadding,
                                            Math.Max(1, cardBounds.Width - CardPadding * 2),
                                            GetCoverHeight())
            DrawCover(graphics, template, coverBounds)

            Dim textLeft = cardBounds.Left + CardPadding
            Dim textWidth = Math.Max(1, cardBounds.Width - CardPadding * 2)
            Dim textTop = coverBounds.Bottom + 8
            Dim titleBounds = New Rectangle(textLeft, textTop, textWidth, 30)
            Dim metaBounds = New Rectangle(textLeft, titleBounds.Bottom + 2, textWidth, 22)
            Dim statusBounds = New Rectangle(textLeft, metaBounds.Bottom + 2, textWidth, 26)
            Dim buttonBounds = New Rectangle(textLeft, cardBounds.Bottom - CardPadding - 28, Math.Min(126, textWidth), 28)

            Dim title = If(String.IsNullOrWhiteSpace(template.Name), template.Id, template.Name.Trim())
            Dim meta = BuildMetaText(template)
            Dim status = GetStatusText(template)

            Using titleFont As New Font(Me.Font.FontFamily, 9.5F, FontStyle.Bold),
                  metaFont As New Font(Me.Font.FontFamily, 8.5F, FontStyle.Regular),
                  statusFont As New Font(Me.Font.FontFamily, 8.0F, FontStyle.Regular)
                TextRenderer.DrawText(graphics, title, titleFont, titleBounds, Color.FromArgb(39, 45, 55), TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
                TextRenderer.DrawText(graphics, meta, metaFont, metaBounds, Color.FromArgb(86, 94, 108), TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
                If Not String.IsNullOrWhiteSpace(status) AndAlso Not _images.ContainsKey(template.Id) Then
                    Dim statusColor = If(IsFailureStatus(status), Color.FromArgb(185, 28, 28), Color.FromArgb(86, 94, 108))
                    TextRenderer.DrawText(graphics, status, statusFont, statusBounds, statusColor, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
                End If
            End Using

            Using buttonBrush As New SolidBrush(If(isSelected, Color.FromArgb(234, 88, 12), Color.FromArgb(241, 245, 249))),
                  buttonFont As New Font(Me.Font.FontFamily, 8.5F, FontStyle.Bold)
                graphics.FillRectangle(buttonBrush, buttonBounds)
                TextRenderer.DrawText(graphics, If(isSelected, "已选择", "选择模板"), buttonFont, buttonBounds, If(isSelected, Color.White, Color.FromArgb(39, 45, 55)), TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            End Using
        End Sub

        Private Sub DrawCover(graphics As Graphics, template As DocmeeTemplateInfo, bounds As Rectangle)
            Using backgroundBrush As New SolidBrush(Color.FromArgb(255, 248, 241)),
                  borderPen As New Pen(Color.FromArgb(226, 232, 240))
                graphics.FillRectangle(backgroundBrush, bounds)
                graphics.DrawRectangle(borderPen, bounds)
            End Using

            Dim image As Image = Nothing
            If template IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(template.Id) AndAlso _images.TryGetValue(template.Id, image) AndAlso image IsNot Nothing Then
                Try
                    DrawImageContain(graphics, image, Rectangle.Inflate(bounds, -1, -1))
                    Return
                Catch ex As Exception
                    ThemePptTaskPane.AppendThemePptLog("PaintGallery draw image failed: id=" & template.Id & ", error=" & ex.Message)
                End Try
            End If

            Dim title = If(template Is Nothing OrElse String.IsNullOrWhiteSpace(template.Name), "模板预览", template.Name.Trim())
            Dim status = GetStatusText(template)
            If String.IsNullOrWhiteSpace(status) Then status = "封面加载中..."

            Using accentBrush As New SolidBrush(Color.FromArgb(234, 88, 12)),
                  titleFont As New Font(Me.Font.FontFamily, 9.5F, FontStyle.Bold),
                  statusFont As New Font(Me.Font.FontFamily, 8.5F, FontStyle.Regular)
                graphics.FillRectangle(accentBrush, bounds.Left, bounds.Top, 5, bounds.Height)
                Dim inner = Rectangle.Inflate(bounds, -14, -12)
                TextRenderer.DrawText(graphics, title, titleFont, New Rectangle(inner.Left, inner.Top + 20, inner.Width, 48), Color.FromArgb(39, 45, 55), TextFormatFlags.WordBreak Or TextFormatFlags.EndEllipsis)
                Dim statusColor = If(IsFailureStatus(status), Color.FromArgb(185, 28, 28), Color.FromArgb(86, 94, 108))
                TextRenderer.DrawText(graphics, status, statusFont, New Rectangle(inner.Left, inner.Bottom - 44, inner.Width, 40), statusColor, TextFormatFlags.WordBreak Or TextFormatFlags.EndEllipsis)
            End Using
        End Sub

        Private Shared Sub DrawImageContain(graphics As Graphics, image As Image, bounds As Rectangle)
            If image Is Nothing OrElse bounds.Width <= 0 OrElse bounds.Height <= 0 Then Return

            Dim scale = Math.Min(bounds.Width / CDbl(image.Width), bounds.Height / CDbl(image.Height))
            Dim width = Math.Max(1, CInt(Math.Round(image.Width * scale)))
            Dim height = Math.Max(1, CInt(Math.Round(image.Height * scale)))
            Dim left = bounds.Left + (bounds.Width - width) \ 2
            Dim top = bounds.Top + (bounds.Height - height) \ 2
            graphics.DrawImage(image, New Rectangle(left, top, width, height))
        End Sub

        Private Function GetStatusText(template As DocmeeTemplateInfo) As String
            If template Is Nothing OrElse String.IsNullOrWhiteSpace(template.Id) Then Return ""
            If _messages.ContainsKey(template.Id) Then Return If(_messages(template.Id), "")
            Return ""
        End Function

        Private Shared Function IsFailureStatus(status As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(status) AndAlso status.Contains("失败")
        End Function

        Private Shared Function BuildMetaText(template As DocmeeTemplateInfo) As String
            Dim parts As New List(Of String)()
            If template IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(template.Category) Then parts.Add(template.Category.Trim())
                If Not String.IsNullOrWhiteSpace(template.Style) Then parts.Add(template.Style.Trim())
            End If

            If parts.Count = 0 Then Return "Docmee 模板"
            Return String.Join(" / ", parts)
        End Function
    End Class

    Private Sub SetStatus(text As String)
        If Not IsOnPaneUiThread() Then
            BeginInvokeIfAlive(CType(Sub() SetStatus(text), MethodInvoker))
            Return
        End If

        AppendThemePptLog("Status: " & If(text, ""))
        _statusLabel.Text = text
    End Sub
End Class
