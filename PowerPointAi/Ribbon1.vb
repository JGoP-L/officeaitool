' PowerPointAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
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
            ShareRibbon.GlobalStatusStripAll.ShowProgress("AI内容提效：正在启动替换单页...")
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
            MessageBox.Show("替换单页出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        Using dialog As New Form()
            OfficeAIStyleHelper.StyleFormDialog(dialog)
            dialog.Text = "替换单页"
            dialog.ClientSize = New Size(420, 220)
            dialog.BackColor = OfficeAIStyleHelper.BgPage

            ' 品牌色标题栏
            Dim header = OfficeAIStyleHelper.CreateFormHeader("替换单页", 420)
            dialog.Controls.Add(header)

            Dim contentY As Integer = header.Bottom + OfficeAIStyleHelper.SpacingMd

            Dim label As New Label() With {
                .Text = "输入新单页要求：",
                .Location = New Point(OfficeAIStyleHelper.SpacingLg, contentY),
                .AutoSize = True
            }
            OfficeAIStyleHelper.StyleLabelBody(label)
            dialog.Controls.Add(label)

            Dim inputBox As New TextBox() With {
                .Location = New Point(OfficeAIStyleHelper.SpacingLg, contentY + 26),
                .Size = New Size(386, 80),
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .Text = "工作与生活失衡的现状与代价"
            }
            OfficeAIStyleHelper.StyleTextBoxMultiline(inputBox)
            dialog.Controls.Add(inputBox)

            Dim okButton As New Button() With {
                .Text = "替换",
                .Location = New Point(224, inputBox.Bottom + OfficeAIStyleHelper.SpacingMd),
                .Size = New Size(86, OfficeAIStyleHelper.ButtonHeight),
                .DialogResult = DialogResult.OK
            }
            OfficeAIStyleHelper.StyleButtonPrimary(okButton)
            dialog.Controls.Add(okButton)
            dialog.AcceptButton = okButton

            Dim cancelButton As New Button() With {
                .Text = "取消",
                .Location = New Point(318, inputBox.Bottom + OfficeAIStyleHelper.SpacingMd),
                .Size = New Size(86, OfficeAIStyleHelper.ButtonHeight),
                .DialogResult = DialogResult.Cancel
            }
            OfficeAIStyleHelper.StyleButtonSecondary(cancelButton)
            dialog.Controls.Add(cancelButton)
            dialog.CancelButton = cancelButton

            If dialog.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return New ReplaceSlideOptions() With {
                .Content = inputBox.Text.Trim()
            }
        End Using
    End Function

    Private Async Function ReplaceCurrentSlideWithDocmeeAsync(request As ReplaceSlideOptions) As Task
        LogInfo("[ReplaceSlide] Start Docmee replace. contentLength=" & If(request?.Content, "").Length.ToString())
        Dim originalSlide = GetCurrentSlide()
        If originalSlide Is Nothing Then
            LogInfo("[ReplaceSlide] No current slide.")
            MessageBox.Show("请先选中要替换的幻灯片。", "替换单页", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
        UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在分析当前页内容...")

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
                ShareRibbon.GlobalStatusStripAll.ShowProgress("正在获取替换单页 PPTX...")
                UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在整理页面结构...")
                LogInfo("[ReplaceSlide] Download URL missing, calling downloadPptx.")
                fileUrl = Await client.DownloadPptxAsync(pptxId, True)
            End If

            Dim localPath = Path.Combine(Path.GetTempPath(), $"wenduoduoAI_newpage_{pptxId}_{Guid.NewGuid():N}.pptx")
            ShareRibbon.GlobalStatusStripAll.ShowProgress("正在下载替换单页 PPTX...")
            UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在下载候选页面...")
            LogInfo("[ReplaceSlide] Downloading PPTX to " & localPath)
            Await client.DownloadPptxFileAsync(fileUrl, localPath)

            ShareRibbon.GlobalStatusStripAll.ShowProgress("请选择要替换的单页...")
            UpdateReplaceProgressWindow(progressWindow.Item1, progressWindow.Item2, "正在准备选择界面...")
            LogInfo("[ReplaceSlide] Selecting replacement slide from downloaded PPTX.")
            CloseReplaceProgressWindow(progressWindow.Item1)
            If Not ReplaceCurrentSlideWithSelectedSlideFromPptx(originalSlide, localPath) Then Return

            _lastReplaceSlidePptxId = pptxId
            SaveDocmeePptxId(pptxId)
            ShareRibbon.GlobalStatusStripAll.ShowProgress("替换单页完成")
            LogInfo("[ReplaceSlide] Completed.")
        Finally
            CloseReplaceProgressWindow(progressWindow.Item1)
        End Try
    End Function

    Private Function CreateReplaceProgressWindow() As Tuple(Of Form, Label)
        Dim dialog As New Form()
        OfficeAIStyleHelper.StyleFormDialog(dialog)
        dialog.Text = "正在生成"
        dialog.ClientSize = New Size(380, 130)
        dialog.ControlBox = False
        dialog.BackColor = OfficeAIStyleHelper.BgPage

        ' 品牌色标题栏
        Dim header = OfficeAIStyleHelper.CreateFormHeader("正在生成替换页", 380)
        dialog.Controls.Add(header)

        Dim contentY As Integer = header.Bottom + OfficeAIStyleHelper.SpacingMd

        Dim label As New Label() With {
            .Text = "正在分析当前页内容...",
            .Location = New Point(OfficeAIStyleHelper.SpacingLg, contentY),
            .Size = New Size(350, 20),
            .AutoEllipsis = True
        }
        OfficeAIStyleHelper.StyleLabelBody(label)
        dialog.Controls.Add(label)

        Dim progressBar As New ProgressBar() With {
            .Location = New Point(OfficeAIStyleHelper.SpacingLg, contentY + 28),
            .Size = New Size(350, 8),
            .Style = ProgressBarStyle.Marquee,
            .MarqueeAnimationSpeed = 30
        }
        dialog.Controls.Add(progressBar)

        Return Tuple.Create(dialog, label)
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
        Using dialog As New Form()
            OfficeAIStyleHelper.StyleFormDialog(dialog)
            dialog.Text = "选择替换页"
            dialog.ClientSize = New Size(960, 580)
            dialog.BackColor = OfficeAIStyleHelper.BgPage

            ' 品牌色标题栏
            Dim header = OfficeAIStyleHelper.CreateFormHeader("选择替换页 - Docmee 生成了多页，请选择用于替换的页面", 960)
            dialog.Controls.Add(header)

            Dim contentY As Integer = header.Bottom + OfficeAIStyleHelper.SpacingSm

            Dim listBox As New ListBox() With {
                .Location = New Point(OfficeAIStyleHelper.SpacingSm, contentY),
                .Size = New Size(280, 420)
            }
            listBox.BackColor = OfficeAIStyleHelper.BgSurface
            listBox.BorderStyle = BorderStyle.FixedSingle
            listBox.Font = OfficeAIStyleHelper.FontUi
            For Each choice In choices
                listBox.Items.Add(choice)
            Next
            dialog.Controls.Add(listBox)

            Dim previewBox As New PictureBox() With {
                .Location = New Point(306, contentY),
                .Size = New Size(630, 354),
                .BackColor = OfficeAIStyleHelper.BgSurface,
                .BorderStyle = BorderStyle.FixedSingle,
                .SizeMode = PictureBoxSizeMode.Zoom
            }
            dialog.Controls.Add(previewBox)

            Dim titleLabel As New Label() With {
                .Location = New Point(306, previewBox.Bottom + OfficeAIStyleHelper.SpacingSm),
                .Size = New Size(630, 40),
                .AutoEllipsis = True
            }
            OfficeAIStyleHelper.StyleLabelBody(titleLabel)
            dialog.Controls.Add(titleLabel)

            Dim setPreview As MethodInvoker =
                Sub()
                    Dim choice = TryCast(listBox.SelectedItem, PptxSlideChoice)
                    If previewBox.Image IsNot Nothing Then
                        Dim oldImage = previewBox.Image
                        previewBox.Image = Nothing
                        oldImage.Dispose()
                    End If

                    If choice Is Nothing Then
                        titleLabel.Text = ""
                        Return
                    End If

                    titleLabel.Text = choice.ToString() & "    正在生成预览..."
                    dialog.Cursor = Cursors.WaitCursor
                    previewBox.Refresh()
                    titleLabel.Refresh()
                    Application.DoEvents()

                    Dim previewPath = EnsureSlidePreview(choice)
                    titleLabel.Text = choice.ToString()
                    dialog.Cursor = Cursors.Default
                    If Not String.IsNullOrWhiteSpace(previewPath) AndAlso File.Exists(previewPath) Then
                        Using previewImage = Image.FromFile(previewPath)
                            previewBox.Image = New Bitmap(previewImage)
                        End Using
                    Else
                        titleLabel.Text = choice.ToString() & "    预览图生成失败"
                    End If
                End Sub

            AddHandler listBox.SelectedIndexChanged, Sub() setPreview.Invoke()
            AddHandler dialog.FormClosed,
                Sub()
                    If previewBox.Image IsNot Nothing Then
                        Dim oldImage = previewBox.Image
                        previewBox.Image = Nothing
                        oldImage.Dispose()
                    End If
                End Sub

            listBox.SelectedIndex = Math.Max(0, choices.Count - 1)

            Dim okButton As New Button() With {
                .Text = "使用此页",
                .Location = New Point(750, titleLabel.Bottom + OfficeAIStyleHelper.SpacingSm),
                .Size = New Size(96, OfficeAIStyleHelper.ButtonHeight),
                .DialogResult = DialogResult.OK
            }
            OfficeAIStyleHelper.StyleButtonPrimary(okButton)
            dialog.Controls.Add(okButton)
            dialog.AcceptButton = okButton

            Dim cancelButton As New Button() With {
                .Text = "取消",
                .Location = New Point(854, titleLabel.Bottom + OfficeAIStyleHelper.SpacingSm),
                .Size = New Size(86, OfficeAIStyleHelper.ButtonHeight),
                .DialogResult = DialogResult.Cancel
            }
            OfficeAIStyleHelper.StyleButtonSecondary(cancelButton)
            dialog.Controls.Add(cancelButton)
            dialog.CancelButton = cancelButton

            If dialog.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return TryCast(listBox.SelectedItem, PptxSlideChoice)
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
        Dim progressWindow = CreateReplaceProgressWindow()
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
                fileUrl = Await client.DownloadPptxAsync(resultPptxId, True)
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
            dialog.ClientSize = New Size(420, 380)
            dialog.BackColor = OfficeAIStyleHelper.BgPage

            ' 品牌色标题栏
            Dim header = OfficeAIStyleHelper.CreateFormHeader("AI 创作", 420)
            dialog.Controls.Add(header)

            Dim contentY As Integer = header.Bottom + OfficeAIStyleHelper.SpacingLg

            ' 提示标签
            Dim label As New Label() With {
                .Text = "选择 AI 创作方式：",
                .Location = New Point(OfficeAIStyleHelper.SpacingLg, contentY),
                .AutoSize = True
            }
            OfficeAIStyleHelper.StyleLabelBody(label)
            dialog.Controls.Add(label)

            ' 卡片式模式选择 (2x2 网格)
            Dim cardY As Integer = contentY + 26
            Dim cardW As Integer = 186
            Dim col1X As Integer = OfficeAIStyleHelper.SpacingLg
            Dim col2X As Integer = col1X + cardW + OfficeAIStyleHelper.SpacingSm

            Dim cardPolish = OfficeAIStyleHelper.CreateOptionCard("✨", "润色", "优化表达，提升文采", cardW)
            cardPolish.Location = New Point(col1X, cardY)
            Dim cardExpand = OfficeAIStyleHelper.CreateOptionCard("⬆", "扩写", "丰富细节，充实内容", cardW)
            cardExpand.Location = New Point(col2X, cardY)
            Dim cardShorten = OfficeAIStyleHelper.CreateOptionCard("⬇", "缩写", "精简文字，突出重点", cardW)
            cardShorten.Location = New Point(col1X, cardY + 80)
            Dim cardTranslate = OfficeAIStyleHelper.CreateOptionCard("文", "翻译", "多语言互译，跨越障碍", cardW)
            cardTranslate.Location = New Point(col2X, cardY + 80)

            ' 卡片点击事件 - 使用 Dictionary 映射卡片到模式名
            Dim allCards = {cardPolish, cardExpand, cardShorten, cardTranslate}
            Dim cardModes As New Dictionary(Of Panel, String) From {
                {cardPolish, "润色"}, {cardExpand, "扩写"}, {cardShorten, "缩写"}, {cardTranslate, "翻译"}
            }
            Dim selectedMode As String = "润色"
            OfficeAIStyleHelper.SetCardSelected(cardPolish, True)

            ' 目标语言选择行 (仅翻译时可见) - 必须在卡片点击事件之前声明
            Dim langY As Integer = cardY + 168

            Dim languageLabel As New Label() With {
                .Text = "目标语言：",
                .Location = New Point(OfficeAIStyleHelper.SpacingLg, langY + 4),
                .AutoSize = True,
                .Visible = False
            }
            OfficeAIStyleHelper.StyleLabelBody(languageLabel)
            dialog.Controls.Add(languageLabel)

            Dim languageCombo As New ComboBox() With {
                .Location = New Point(96, langY),
                .Size = New Size(300, OfficeAIStyleHelper.InputHeight),
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
            Dim sepHorizontal = OfficeAIStyleHelper.CreateSeparator(386)
            sepHorizontal.Location = New Point(OfficeAIStyleHelper.SpacingLg, btnY - OfficeAIStyleHelper.SpacingSm)
            dialog.Controls.Add(sepHorizontal)

            Dim okButton As New Button() With {
                .Text = "开始创作",
                .Location = New Point(212, btnY),
                .Size = New Size(110, OfficeAIStyleHelper.ButtonHeight),
                .DialogResult = DialogResult.OK
            }
            OfficeAIStyleHelper.StyleButtonPrimary(okButton)
            dialog.Controls.Add(okButton)
            dialog.AcceptButton = okButton

            Dim cancelButton As New Button() With {
                .Text = "取消",
                .Location = New Point(330, btnY),
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
