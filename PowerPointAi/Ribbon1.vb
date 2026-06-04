' PowerPointAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon  ' 添加此引用
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Private Const DemoAccentShapeName As String = "wenduoduoAI_BeautifyAccent"

    Private Class PptTextTarget
        Public Property TextRange As PowerPoint.TextRange
        Public Property Shape As PowerPoint.Shape
        Public Property OriginalText As String
        Public Property SlideContextText As String
    End Class

    Protected Overrides Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub

    Protected Overrides Sub WebResearchButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub

    Protected Overrides Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs)
        'Globals.ThisAddIn.ShowChatTaskPane()
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
            ' 在需要时可以集成到ChatControl调用MCP服务
        End If
    End Sub

    Protected Overrides Async Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim requirement = ShowReplaceSlideDialog()
            If String.IsNullOrWhiteSpace(requirement) Then Return

            Await ReplaceCurrentSlideWithGeneratedTextAsync(requirement)
        Catch ex As Exception
            MessageBox.Show("替换单页出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 在聊天面板显示提示信息
    ''' </summary>
    Private Async Sub buildHtmlHint(chatCtrl As ChatControl, displayContent As String)
        Try
            Dim responseUuid As String = Guid.NewGuid().ToString()
            Dim aiName As String = ShareRibbon.ConfigSettings.platform & " " & ShareRibbon.ConfigSettings.ModelName
            Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
            Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCreate)
            Dim js = $"appendRenderer('{responseUuid}','{displayContent}');"
            Await chatCtrl.ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine("ExecuteJavaScriptAsyncJS 调用失败: " & ex.Message)
        End Try
    End Sub

    ' 美化当前页 - 用本地规则统一背景、字体、标题和文本框适配。
    Protected Overrides Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim changedCount = ApplySimpleBeautifyToCurrentSlide()
            If changedCount <= 0 Then
                MessageBox.Show("当前页没有可美化的文本框。", "美化单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            MessageBox.Show($"已美化当前页，处理 {changedCount} 个文本框。", "美化单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("美化单页出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Function ShowTextOptimizeDialog() As String
        Using dialog As New Form()
            dialog.Text = "文本优化"
            dialog.Size = New Size(330, 170)
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MaximizeBox = False
            dialog.MinimizeBox = False

            Dim label As New Label() With {
                .Text = "选择优化方式：",
                .Location = New Point(18, 18),
                .AutoSize = True
            }
            dialog.Controls.Add(label)

            Dim modeCombo As New ComboBox() With {
                .Location = New Point(18, 44),
                .Size = New Size(278, 24),
                .DropDownStyle = ComboBoxStyle.DropDownList
            }
            modeCombo.Items.AddRange(New Object() {"润色", "扩写", "精简", "填充", "补全文案"})
            modeCombo.SelectedIndex = 0
            dialog.Controls.Add(modeCombo)

            Dim okButton As New Button() With {
                .Text = "开始",
                .Location = New Point(126, 86),
                .Size = New Size(80, 30),
                .DialogResult = DialogResult.OK
            }
            dialog.Controls.Add(okButton)
            dialog.AcceptButton = okButton

            Dim cancelButton As New Button() With {
                .Text = "取消",
                .Location = New Point(216, 86),
                .Size = New Size(80, 30),
                .DialogResult = DialogResult.Cancel
            }
            dialog.Controls.Add(cancelButton)
            dialog.CancelButton = cancelButton

            If dialog.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return modeCombo.SelectedItem.ToString()
        End Using
    End Function

    Private Function ShowReplaceSlideDialog() As String
        Using dialog As New Form()
            dialog.Text = "替换单页"
            dialog.Size = New Size(430, 250)
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MaximizeBox = False
            dialog.MinimizeBox = False

            Dim label As New Label() With {
                .Text = "输入新单页要求：",
                .Location = New Point(18, 16),
                .AutoSize = True
            }
            dialog.Controls.Add(label)

            Dim inputBox As New TextBox() With {
                .Location = New Point(18, 42),
                .Size = New Size(374, 118),
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .Text = "生成一页关于核心结论的汇报页，包含标题和 3 个要点。"
            }
            dialog.Controls.Add(inputBox)

            Dim okButton As New Button() With {
                .Text = "替换",
                .Location = New Point(222, 174),
                .Size = New Size(80, 30),
                .DialogResult = DialogResult.OK
            }
            dialog.Controls.Add(okButton)
            dialog.AcceptButton = okButton

            Dim cancelButton As New Button() With {
                .Text = "取消",
                .Location = New Point(312, 174),
                .Size = New Size(80, 30),
                .DialogResult = DialogResult.Cancel
            }
            dialog.Controls.Add(cancelButton)
            dialog.CancelButton = cancelButton

            If dialog.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return inputBox.Text.Trim()
        End Using
    End Function

    Private Async Function ReplaceCurrentSlideWithGeneratedTextAsync(requirement As String) As Task
        Dim originalSlide = GetCurrentSlide()
        If originalSlide Is Nothing Then
            MessageBox.Show("请先选中要替换的幻灯片。", "替换单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ShareRibbon.GlobalStatusStripAll.ShowProgress("正在生成替换单页内容...")
        Dim generatedText = Await GenerateReplacementSlideTextAsync(requirement)
        If String.IsNullOrWhiteSpace(generatedText) Then
            MessageBox.Show("没有生成可用的替换内容。", "替换单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim replacementSlide = InsertReplacementSlideAfter(originalSlide, generatedText)
        DeleteOriginalSlideAfterReplacement(originalSlide, replacementSlide)
        ApplySimpleBeautifyToSlide(replacementSlide)

        ShareRibbon.GlobalStatusStripAll.ShowProgress("替换单页完成")
        MessageBox.Show("当前页已替换。", "替换单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Function

    Private Async Function GenerateReplacementSlideTextAsync(requirement As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) OrElse
           String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) OrElse
           String.IsNullOrWhiteSpace(ConfigSettings.ModelName) Then
            Throw New InvalidOperationException("请先在模型配置里填写 API 地址、API Key 和模型名称。")
        End If

        Dim systemPrompt = "你是一个专业的 PowerPoint 单页生成助手。只返回纯文本：第一行是标题，后续每行一个要点。不要解释，不要 Markdown。"
        Dim prompt = "请根据下面要求生成一页 PPT 文案：" & vbCrLf & requirement.Trim()
        Dim requestBody = LLMUtil.CreateLlmRequestBody(prompt, ConfigSettings.ModelName, systemPrompt, 0.35, 1000)
        Dim response = Await LLMUtil.SendHttpRequest(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)

        If response.StartsWith("错误:") Then Throw New Exception(response)

        Dim jObj = JObject.Parse(response)
        Dim content = jObj("choices")(0)("message")("content")?.ToString()
        If String.IsNullOrWhiteSpace(content) Then
            Throw New InvalidOperationException("模型没有返回可用的替换单页内容。")
        End If

        Return content.Trim()
    End Function

    Private Function InsertReplacementSlideAfter(originalSlide As PowerPoint.Slide, generatedText As String) As PowerPoint.Slide
        Dim presentation = Globals.ThisAddIn.Application.ActivePresentation
        Dim originalIndex = originalSlide.SlideIndex
        Dim replacementSlide = presentation.Slides.Add(originalIndex + 1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly)

        Dim normalized = generatedText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim lines = normalized.Split(New String() {vbLf}, StringSplitOptions.None).
            Select(Function(line) line.Trim()).
            Where(Function(line) Not String.IsNullOrWhiteSpace(line)).
            ToList()

        Dim title = If(lines.Count > 0, lines(0), "新单页")
        Dim body = If(lines.Count > 1, String.Join(vbCrLf, lines.Skip(1)), generatedText.Trim())

        SetReplacementSlideText(replacementSlide, title, body)
        Return replacementSlide
    End Function

    Private Sub SetReplacementSlideText(slide As PowerPoint.Slide, titleText As String, bodyText As String)
        Dim presentation = Globals.ThisAddIn.Application.ActivePresentation
        Dim slideWidth As Single = CSng(presentation.PageSetup.SlideWidth)
        Dim slideHeight As Single = CSng(presentation.PageSetup.SlideHeight)

        Dim titleShape As PowerPoint.Shape = Nothing
        Try
            If slide.Shapes.HasTitle = Microsoft.Office.Core.MsoTriState.msoTrue Then
                titleShape = slide.Shapes.Title
            End If
        Catch
        End Try

        If titleShape Is Nothing Then
            titleShape = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 46, 38, slideWidth - 92, 64)
        End If

        titleShape.TextFrame.TextRange.Text = titleText
        titleShape.TextFrame.TextRange.Font.Name = "Microsoft YaHei UI"
        titleShape.TextFrame.TextRange.Font.Size = 30
        titleShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue

        Dim bodyShape = slide.Shapes.AddTextbox(
            Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
            54,
            118,
            slideWidth - 108,
            slideHeight - 168)
        bodyShape.TextFrame.TextRange.Text = bodyText
        bodyShape.TextFrame.TextRange.Font.Name = "Microsoft YaHei UI"
        bodyShape.TextFrame.TextRange.Font.Size = 18
        bodyShape.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue
    End Sub

    Private Sub DeleteOriginalSlideAfterReplacement(originalSlide As PowerPoint.Slide, replacementSlide As PowerPoint.Slide)
        Dim targetIndex = originalSlide.SlideIndex
        originalSlide.Delete()

        Try
            replacementSlide.Select()
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(targetIndex)
        Catch
        End Try
    End Sub

    Private Async Function OptimizeSelectedTextAsync(modeName As String) As Task(Of Integer)
        Dim targets = GetSelectedPptTextTargets(modeName)
        If targets.Count = 0 Then
            MessageBox.Show("请先选中 PPT 里的文字或文本框。", "文本优化", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return 0
        End If

        If String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) OrElse
           String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) OrElse
           String.IsNullOrWhiteSpace(ConfigSettings.ModelName) Then
            MessageBox.Show("请先在模型配置里填写 API 地址、API Key 和模型名称。", "文本优化", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return 0
        End If

        ShareRibbon.GlobalStatusStripAll.ShowProgress($"正在{modeName}选中文本...")
        Dim changedCount = 0

        For Each target In targets
            Dim optimizedText = Await RequestOptimizedTextAsync(modeName, target.OriginalText, target.SlideContextText)
            If Not String.IsNullOrWhiteSpace(optimizedText) Then
                target.TextRange.Text = optimizedText
                If target.Shape IsNot Nothing Then AutoFitPptTextShape(target.Shape)
                changedCount += 1
            End If
        Next

        ShareRibbon.GlobalStatusStripAll.ShowProgress($"文本优化完成，共处理 {changedCount} 处")
        Return changedCount
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

    Private Async Function RequestOptimizedTextAsync(modeName As String, originalText As String, slideContextText As String) As Task(Of String)
        Dim systemPrompt = "你是一个专业的 PowerPoint 文案优化助手。你只返回处理后的文本，不要解释，不要添加标题，不要输出 Markdown。"
        Dim prompt = BuildTextOptimizationPrompt(modeName, originalText, slideContextText)
        Dim requestBody = LLMUtil.CreateLlmRequestBody(prompt, ConfigSettings.ModelName, systemPrompt, 0.35, 1200)
        Dim response = Await LLMUtil.SendHttpRequest(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)

        If response.StartsWith("错误:") Then
            Throw New Exception(response)
        End If

        Dim jObj = JObject.Parse(response)
        Dim content = jObj("choices")(0)("message")("content")?.ToString()
        If String.IsNullOrWhiteSpace(content) Then Throw New Exception("模型返回内容为空")

        Return content.Trim()
    End Function

    Private Function BuildTextOptimizationPrompt(modeName As String, originalText As String, slideContextText As String) As String
        Dim instruction As String
        Dim sourceText = If(originalText, "").Trim()
        Dim contextText = If(slideContextText, "").Trim()
        Select Case modeName
            Case "扩写"
                instruction = "在不偏离原意的前提下扩写为更适合 PPT 展示的内容，语言自然、有条理。"
            Case "精简"
                instruction = "精简为更适合 PPT 的短句，保留核心信息，删除重复和口语化表达。"
            Case "填充"
                instruction = "填充当前 PPT 文本框，使内容更完整、更充实，并保持适合幻灯片展示的长度。"
            Case "补全文案"
                instruction = "补全为一段完整、清晰、适合放在 PPT 文本框中的文案。"
            Case Else
                instruction = "润色为更专业、更清晰、更适合 PPT 展示的表达。"
        End Select

        Dim contextSection = If(String.IsNullOrWhiteSpace(contextText), "", vbCrLf & vbCrLf & "当前页其他内容：" & vbCrLf & contextText)

        Return $"请执行：{instruction}

原文：
{If(String.IsNullOrWhiteSpace(sourceText), "当前文本框为空，请根据当前幻灯片语境生成适合填入该文本框的内容。", sourceText)}{contextSection}"
    End Function

    Private Function ApplySimpleBeautifyToCurrentSlide() As Integer
        Dim slide = GetCurrentSlide()
        Return ApplySimpleBeautifyToSlide(slide)
    End Function

    Private Function ApplySimpleBeautifyToSlide(slide As PowerPoint.Slide) As Integer
        If slide Is Nothing Then Return 0

        Dim presentation = Globals.ThisAddIn.Application.ActivePresentation
        Dim slideWidth As Single = CSng(presentation.PageSetup.SlideWidth)
        Dim slideHeight As Single = CSng(presentation.PageSetup.SlideHeight)

        RemoveDemoAccentShapes(slide)

        Try
            slide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse
            slide.Background.Fill.Solid()
            slide.Background.Fill.ForeColor.RGB = RGB(248, 250, 252)
        Catch ex As Exception
            Debug.WriteLine("设置幻灯片背景失败: " & ex.Message)
        End Try

        Dim accent = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth, 8)
        accent.Name = DemoAccentShapeName
        accent.Fill.ForeColor.RGB = RGB(37, 99, 235)
        accent.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse

        Dim changedCount = 0
        For shapeIdx = 1 To slide.Shapes.Count
            Dim shape = slide.Shapes(shapeIdx)
            If Not shape.Name.StartsWith(DemoAccentShapeName, StringComparison.OrdinalIgnoreCase) Then
                changedCount += BeautifyShapeText(shape, slideWidth, slideHeight)
            End If
        Next

        Return changedCount
    End Function

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

    Private Sub RemoveDemoAccentShapes(slide As PowerPoint.Slide)
        For i = slide.Shapes.Count To 1 Step -1
            Try
                If slide.Shapes(i).Name.StartsWith(DemoAccentShapeName, StringComparison.OrdinalIgnoreCase) Then
                    slide.Shapes(i).Delete()
                End If
            Catch
            End Try
        Next
    End Sub

    Private Function BeautifyShapeText(shape As PowerPoint.Shape, slideWidth As Single, slideHeight As Single) As Integer
        Try
            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoGroup Then
                Dim count = 0
                For i = 1 To shape.GroupItems.Count
                    count += BeautifyShapeText(shape.GroupItems(i), slideWidth, slideHeight)
                Next
                Return count
            End If

            If shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                Dim count = 0
                For row = 1 To shape.Table.Rows.Count
                    For col = 1 To shape.Table.Columns.Count
                        count += BeautifyShapeText(shape.Table.Cell(row, col).Shape, slideWidth, slideHeight)
                    Next
                Next
                Return count
            End If

            If shape.HasTextFrame <> Microsoft.Office.Core.MsoTriState.msoTrue OrElse
               shape.TextFrame.HasText <> Microsoft.Office.Core.MsoTriState.msoTrue Then Return 0

            Dim textRange = shape.TextFrame.TextRange
            textRange.Text = CleanPptText(textRange.Text)

            Dim isTitle = IsLikelyTitleShape(shape, slideHeight)
            textRange.Font.Name = "Microsoft YaHei UI"
            textRange.Font.Size = If(isTitle, 30, 18)
            textRange.Font.Bold = If(isTitle, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse)
            textRange.Font.Color.RGB = If(isTitle, RGB(15, 23, 42), RGB(51, 65, 85))
            textRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft

            shape.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue
            shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone

            If isTitle Then
                shape.Left = 36
                shape.Top = Math.Max(24, shape.Top)
                shape.Width = Math.Max(120, slideWidth - 72)
            Else
                Try
                    shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    shape.Line.ForeColor.RGB = RGB(226, 232, 240)
                    shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    shape.Fill.Transparency = 0.08F
                Catch
                End Try
            End If

            AutoFitPptTextShape(shape)
            Return 1
        Catch ex As Exception
            Debug.WriteLine("美化文本框失败: " & ex.Message)
            Return 0
        End Try
    End Function

    Private Function CleanPptText(text As String) As String
        Dim normalized = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim lines = normalized.Split(New String() {vbLf}, StringSplitOptions.None).
            Select(Function(line) Regex.Replace(line.Trim(), "\s+", " ")).
            Where(Function(line) Not String.IsNullOrWhiteSpace(line)).
            ToList()

        Return String.Join(vbCrLf, lines)
    End Function

    Private Function IsLikelyTitleShape(shape As PowerPoint.Shape, slideHeight As Single) As Boolean
        Try
            If shape.PlaceholderFormat IsNot Nothing Then
                Select Case shape.PlaceholderFormat.Type
                    Case PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                         PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                        Return True
                End Select
            End If
        Catch
        End Try

        Return shape.Top < slideHeight * 0.24F
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

    ' 一键翻译功能 - PowerPoint实现
    Protected Overrides Async Sub TranslateButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim pptApp = Globals.ThisAddIn.Application

            ' 检查是否有选中内容
            Dim hasSelection As Boolean = False
            Try
                Dim sel = pptApp.ActiveWindow.Selection
                hasSelection = (sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText OrElse
                               sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes OrElse
                               sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides)
            Catch
                hasSelection = False
            End Try

            ' 显示翻译操作对话框
            Dim actionForm As New ShareRibbon.TranslateActionForm(hasSelection, "PowerPoint")
            If actionForm.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            ' 创建翻译服务
            Dim translateService As New PowerPointDocumentTranslateService(pptApp)

            ' 更新设置
            Dim settings = ShareRibbon.TranslateSettings.Load()
            settings.SourceLanguage = actionForm.SourceLanguage
            settings.TargetLanguage = actionForm.TargetLanguage
            settings.CurrentDomain = actionForm.SelectedDomain
            settings.OutputMode = actionForm.OutputMode
            settings.Save()

            ' 显示进度
            ShareRibbon.GlobalStatusStripAll.ShowProgress("正在准备翻译... " & translateService.GetStatistics())

            ' 绑定进度事件 - 使用ShowProgress避免翻译过程中频繁弹窗
            AddHandler translateService.ProgressChanged, Sub(s, args)
                                                             ShareRibbon.GlobalStatusStripAll.ShowProgress(args.Message)
                                                         End Sub

            ' 执行翻译
            Dim results As List(Of ShareRibbon.TranslateParagraphResult)
            If actionForm.TranslateCurrentSlide Then
                results = Await translateService.TranslateCurrentSlideAsync()
            ElseIf actionForm.TranslateAll Then
                results = Await translateService.TranslateAllAsync()
            Else
                results = Await translateService.TranslateSelectionAsync()
            End If

            ' 应用翻译结果
            If actionForm.OutputMode = ShareRibbon.TranslateOutputMode.SidePanel Then
                ' 在侧栏显示
                Globals.ThisAddIn.ShowChatTaskPane()
                Await Task.Delay(250)

                Dim chatCtrl = ThisAddIn.chatControl
                If chatCtrl IsNot Nothing Then
                    Dim displayText = translateService.FormatResultsForDisplay(results, True)
                    Dim responseUuid As String = Guid.NewGuid().ToString()
                    Dim aiName As String = "AI翻译助手"
                    Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
                    Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCreate)

                    ' 转义特殊字符
                    Dim escapedText = displayText.Replace("\", "\\").Replace("'", "\'").Replace("</script>", "<\/script>").Replace(vbCr, "").Replace(vbLf, "\n")
                    Dim js = $"appendRenderer('{responseUuid}','{escapedText}');"
                    Await chatCtrl.ExecuteJavaScriptAsyncJS(js)
                End If
            Else
                ' 应用到演示文稿
                If actionForm.TranslateCurrentSlide Then
                    translateService.ApplyTranslationToSelection(results, actionForm.OutputMode)
                ElseIf actionForm.TranslateAll Then
                    translateService.ApplyTranslation(results, actionForm.OutputMode)
                Else
                    translateService.ApplyTranslationToSelection(results, actionForm.OutputMode)
                End If
            End If

            ShareRibbon.GlobalStatusStripAll.ShowProgress($"翻译完成，共处理 {results.Count} 个文本块")

        Catch ex As Exception
            MessageBox.Show("翻译过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 文本优化 - 复用当前模型配置。
    Protected Overrides Async Sub ContinuationButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim modeName = ShowTextOptimizeDialog()
            If String.IsNullOrWhiteSpace(modeName) Then Return

            Dim changedCount = Await OptimizeSelectedTextAsync(modeName)
            If changedCount > 0 Then
                MessageBox.Show($"文本优化完成，共处理 {changedCount} 处。", "文本优化", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("文本优化出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 接受补全功能 - PowerPoint实现
    Protected Sub AcceptCompletionButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim completionManager = PowerPointCompletionManager.Instance
            If completionManager IsNot Nothing AndAlso completionManager.HasGhostText Then
                completionManager.AcceptCurrentCompletion()
            Else
                ' 没有可接受的补全时，显示提示
                MessageBox.Show("当前没有可接受的补全建议。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("接受补全时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
