Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Newtonsoft.Json.Linq
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class ThemePptTaskPane
    Inherits UserControl

    Private ReadOnly _pptApp As PowerPoint.Application
    Private ReadOnly _client As New DocmeePptClient()
    Private _outline As JObject
    Private _taskId As String

    Private ReadOnly _topicBox As New TextBox()
    Private ReadOnly _generateButton As New Button()
    Private ReadOnly _insertButton As New Button()
    Private ReadOnly _templateCombo As New ComboBox()
    Private ReadOnly _refreshTemplatesButton As New Button()
    Private ReadOnly _outputBox As New TextBox()
    Private ReadOnly _statusLabel As New Label()

    Public Sub New(pptApp As PowerPoint.Application)
        _pptApp = pptApp
        BuildLayout()
    End Sub

    Private Sub BuildLayout()
        Me.BackColor = Color.White
        Me.Padding = New Padding(14)

        Dim layout As New TableLayoutPanel()
        layout.Dock = DockStyle.Fill
        layout.ColumnCount = 1
        layout.RowCount = 8
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 96.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim titleLabel As New Label()
        titleLabel.AutoSize = True
        titleLabel.Font = New Font(Me.Font.FontFamily, 12.0F, FontStyle.Bold)
        titleLabel.Text = "主题生成PPT"
        titleLabel.Margin = New Padding(0, 0, 0, 8)

        _topicBox.Dock = DockStyle.Fill
        _topicBox.Multiline = True
        _topicBox.ScrollBars = ScrollBars.Vertical
        _topicBox.Text = "AI 办公趋势"
        _topicBox.Margin = New Padding(0, 0, 0, 10)

        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.AutoSize = True
        buttonPanel.Dock = DockStyle.Fill
        buttonPanel.FlowDirection = FlowDirection.LeftToRight
        buttonPanel.WrapContents = False
        buttonPanel.Margin = New Padding(0, 0, 0, 10)

        _generateButton.Text = "生成大纲"
        _generateButton.Width = 104
        _generateButton.Height = 32
        AddHandler _generateButton.Click, AddressOf GenerateButton_Click

        _insertButton.Text = "生成并导入"
        _insertButton.Width = 118
        _insertButton.Height = 32
        _insertButton.Enabled = False
        AddHandler _insertButton.Click, AddressOf InsertButton_Click

        buttonPanel.Controls.Add(_generateButton)
        buttonPanel.Controls.Add(_insertButton)

        Dim templateLabel As New Label()
        templateLabel.AutoSize = True
        templateLabel.Text = "模板"
        templateLabel.Margin = New Padding(0, 0, 0, 4)

        Dim templatePanel As New TableLayoutPanel()
        templatePanel.Dock = DockStyle.Fill
        templatePanel.ColumnCount = 2
        templatePanel.RowCount = 1
        templatePanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        templatePanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        templatePanel.Margin = New Padding(0, 0, 0, 10)

        _templateCombo.Dock = DockStyle.Fill
        _templateCombo.DropDownStyle = ComboBoxStyle.DropDownList
        _templateCombo.Enabled = False

        _refreshTemplatesButton.Text = "刷新"
        _refreshTemplatesButton.Width = 66
        _refreshTemplatesButton.Height = 28
        AddHandler _refreshTemplatesButton.Click, AddressOf RefreshTemplatesButton_Click

        templatePanel.Controls.Add(_templateCombo, 0, 0)
        templatePanel.Controls.Add(_refreshTemplatesButton, 1, 0)

        _statusLabel.AutoSize = True
        _statusLabel.ForeColor = Color.FromArgb(86, 94, 108)
        _statusLabel.Text = "输入主题，生成大纲后选择模板，再生成并导入 PPT。"
        _statusLabel.Margin = New Padding(0, 0, 0, 8)

        _outputBox.Dock = DockStyle.Fill
        _outputBox.Multiline = True
        _outputBox.ReadOnly = True
        _outputBox.ScrollBars = ScrollBars.Vertical
        _outputBox.BackColor = Color.FromArgb(248, 250, 252)
        _outputBox.BorderStyle = BorderStyle.FixedSingle
        _outputBox.Margin = New Padding(0, 0, 0, 10)

        Dim hintLabel As New Label()
        hintLabel.AutoSize = False
        hintLabel.Dock = DockStyle.Fill
        hintLabel.Height = 42
        hintLabel.ForeColor = Color.FromArgb(86, 94, 108)
        hintLabel.Text = "演示版使用 test.docmee.cn 和 ak_demo，生成的 PPTX 会下载到临时目录并导入当前演示文稿。"

        layout.Controls.Add(titleLabel, 0, 0)
        layout.Controls.Add(_topicBox, 0, 1)
        layout.Controls.Add(buttonPanel, 0, 2)
        layout.Controls.Add(templateLabel, 0, 3)
        layout.Controls.Add(templatePanel, 0, 4)
        layout.Controls.Add(_statusLabel, 0, 5)
        layout.Controls.Add(_outputBox, 0, 6)
        layout.Controls.Add(hintLabel, 0, 7)

        Me.Controls.Add(layout)
    End Sub

    Private Async Sub GenerateButton_Click(sender As Object, e As EventArgs)
        Await GenerateOutlineAsync()
    End Sub

    Private Async Sub RefreshTemplatesButton_Click(sender As Object, e As EventArgs)
        Await LoadTemplatesAsync()
    End Sub

    Private Async Function GenerateOutlineAsync() As Task
        Dim topic = _topicBox.Text.Trim()
        If String.IsNullOrWhiteSpace(topic) Then
            MessageBox.Show("请输入 PPT 主题。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        _generateButton.Enabled = False
        _insertButton.Enabled = False
        _outputBox.Clear()

        Try
            Dim requestContent = BuildRequestContent(topic)
            SetStatus("正在创建 Docmee 任务...")
            _taskId = Await _client.CreateTaskAsync(requestContent)

            SetStatus("正在生成 PPT 大纲...")
            _outputBox.Clear()
            _outline = Await _client.GenerateContentAsync(_taskId, AddressOf AppendOutlineStreamText)

            _outputBox.Text = RenderOutlineText(_outline)
            Await LoadTemplatesAsync()
            _insertButton.Enabled = True
            SetStatus("大纲已生成，请选择模板后生成并导入。")
        Catch ex As Exception
            SetStatus("生成失败。")
            MessageBox.Show("主题生成PPT失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _generateButton.Enabled = True
        End Try
    End Function

    Private Sub AppendOutlineStreamText(chunkText As String)
        If String.IsNullOrEmpty(chunkText) Then Return

        If _outputBox.InvokeRequired Then
            _outputBox.BeginInvoke(CType(Sub() AppendOutlineStreamText(chunkText), MethodInvoker))
            Return
        End If

        _outputBox.AppendText(chunkText)
        _outputBox.SelectionStart = _outputBox.TextLength
        _outputBox.ScrollToCaret()
    End Sub

    Private Async Sub InsertButton_Click(sender As Object, e As EventArgs)
        Await GenerateAndImportPptxAsync()
    End Sub

    Private Async Function LoadTemplatesAsync() As Task
        _refreshTemplatesButton.Enabled = False

        Try
            SetStatus("正在加载 Docmee 模板...")
            Dim templates = Await _client.ListTemplatesAsync(1, 20)
            _templateCombo.Items.Clear()

            For Each template In templates
                _templateCombo.Items.Add(template)
            Next

            _templateCombo.Enabled = _templateCombo.Items.Count > 0
            If _templateCombo.Items.Count > 0 AndAlso _templateCombo.SelectedIndex < 0 Then
                _templateCombo.SelectedIndex = 0
            End If

            If _templateCombo.Items.Count = 0 Then
                SetStatus("未获取到可用模板。")
            Else
                SetStatus($"已加载 {_templateCombo.Items.Count} 个模板。")
            End If
        Catch ex As Exception
            _templateCombo.Enabled = False
            SetStatus("模板加载失败。")
            MessageBox.Show("加载模板失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            _refreshTemplatesButton.Enabled = True
        End Try
    End Function

    Private Async Function GenerateAndImportPptxAsync() As Task
        If String.IsNullOrWhiteSpace(_taskId) OrElse _outline Is Nothing Then
            MessageBox.Show("请先生成大纲。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim selectedTemplate = TryCast(_templateCombo.SelectedItem, DocmeeTemplateInfo)
        If selectedTemplate Is Nothing OrElse String.IsNullOrWhiteSpace(selectedTemplate.Id) Then
            MessageBox.Show("请先选择模板。", "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        _generateButton.Enabled = False
        _insertButton.Enabled = False
        _refreshTemplatesButton.Enabled = False

        Try
            Dim markdown = ConvertOutlineToMarkdown(_outline)

            SetStatus("正在按所选模板生成 PPTX...")
            Dim pptId = Await _client.GeneratePptxAsync(_taskId, selectedTemplate.Id, markdown)

            SetStatus("正在获取 PPTX 下载地址...")
            Dim fileUrl = Await _client.DownloadPptxAsync(pptId)

            SetStatus("正在下载 PPTX...")
            Dim localPath = Path.Combine(Path.GetTempPath(), $"wenduoduoAI_{pptId}.pptx")
            Await _client.DownloadPptxFileAsync(fileUrl, localPath)

            SetStatus("正在导入当前演示文稿...")
            ImportPptxIntoPresentation(localPath)
            SetStatus("已生成并导入当前演示文稿。")
        Catch ex As Exception
            SetStatus("生成或导入失败。")
            MessageBox.Show("生成并导入 PPT 失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _generateButton.Enabled = True
            _insertButton.Enabled = True
            _refreshTemplatesButton.Enabled = True
        End Try
    End Function

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

    Private Sub ImportPptxIntoPresentation(downloadPath As String)
        If String.IsNullOrWhiteSpace(downloadPath) OrElse Not File.Exists(downloadPath) Then
            Throw New FileNotFoundException("未找到下载后的 PPTX 文件。", downloadPath)
        End If

        Dim target = GetOrCreatePresentation()
        Dim insertAfter = target.Slides.Count
        Dim insertedCount = target.Slides.InsertFromFile(downloadPath, insertAfter)

        If insertedCount > 0 Then
            FixInsertedSlideReadability(target, insertAfter + 1, insertAfter + insertedCount)
        End If
    End Sub

    Private Sub FixInsertedSlideReadability(presentation As PowerPoint.Presentation, startIndex As Integer, endIndex As Integer)
        Dim safeStart = Math.Max(1, startIndex)
        Dim safeEnd = Math.Min(endIndex, presentation.Slides.Count)

        For slideIndex As Integer = safeStart To safeEnd
            Dim slide = presentation.Slides(slideIndex)
            Dim slideBackgroundLuminance = GetSlideBackgroundLuminance(slide)

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
            Return GetColorLuminance(slide.Background.Fill.ForeColor.RGB)
        Catch
            Return 245.0R
        End Try
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

        Return $"请生成一份关于 {topic} 的产品介绍 PPT"
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

    Private Sub SetStatus(text As String)
        _statusLabel.Text = text
    End Sub
End Class
