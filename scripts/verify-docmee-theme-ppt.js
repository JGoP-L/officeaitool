const fs = require('fs');

function read(path) {
  return fs.readFileSync(path, 'utf8');
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

const ribbonDesigner = read('PowerPointAi/Ribbon1.Designer.vb');
const ribbon = read('PowerPointAi/Ribbon1.vb');
const addIn = read('PowerPointAi/ThisAddIn.vb');
const project = read('PowerPointAi/PowerPointAi.vbproj');
const client = fs.existsSync('PowerPointAi/DocmeePptClient.vb') ? read('PowerPointAi/DocmeePptClient.vb') : '';
const pane = fs.existsSync('PowerPointAi/ThemePptTaskPane.vb') ? read('PowerPointAi/ThemePptTaskPane.vb') : '';

assert(ribbonDesigner.includes('Me.TemplateFormatButton.Visible = True'), '主题生成PPT button must be visible');
assert(ribbonDesigner.includes('Me.TemplateFormatButton.Label = "主题生成PPT"'), 'TemplateFormatButton must be relabeled 主题生成PPT');
assert(ribbon.includes('Globals.ThisAddIn.ShowThemePptTaskPane()'), 'Ribbon must open the theme PPT task pane');

assert(addIn.includes('themePptTaskPane'), 'ThisAddIn must keep a theme PPT task pane');
assert(addIn.includes('ShowThemePptTaskPane'), 'ThisAddIn must expose ShowThemePptTaskPane');
assert(addIn.includes('New ThemePptTaskPane(Me.Application)'), 'ThisAddIn must create ThemePptTaskPane');

assert(project.includes('Compile Include="DocmeePptClient.vb"'), 'PowerPoint project must compile DocmeePptClient');
assert(project.includes('Compile Include="ThemePptTaskPane.vb"'), 'PowerPoint project must compile ThemePptTaskPane');

assert(client.includes('https://test.docmee.cn'), 'Docmee client must use the test API base URL');
assert(client.includes('Private Const DemoToken As String = "ak_demo"'), 'Docmee client must use ak_demo token for demo');
assert(client.includes('/api/ppt/v2/createTask'), 'Docmee client must call createTask');
assert(client.includes('MultipartFormDataContent'), 'createTask must use multipart/form-data');
assert(client.includes('Private Async Function CreateTaskAsync(content As String, taskType As String)'), 'Docmee client must support creating typed Docmee tasks');
assert(client.includes('form.Add(New StringContent(taskType, Encoding.UTF8), "type")'), 'Docmee createTask must send the requested task type');
assert(client.includes('Public Async Function CreateMarkdownTaskAsync(markdown As String) As Task(Of String)'), 'Docmee client must create fresh markdown tasks for PPT generation');
assert(client.includes('Return Await CreateTaskAsync(markdown, "7")'), 'Docmee markdown task creation must use type=7');
assert(client.includes('/api/ppt/v2/generateContent'), 'Docmee client must call generateContent');
assert(client.includes('GenerateMarkdownContentAsync'), 'Docmee client must expose markdown outline generation');
assert(client.includes('"outlineType", "MD"'), 'generateContent must request markdown outline for the task pane');
assert(client.includes('"questionMode", False'), 'generateContent must disable questionMode');
assert(client.includes('"isNeedAsk", False'), 'generateContent must disable isNeedAsk');
assert(client.includes('Optional progressHandler As Action(Of String) = Nothing'), 'GenerateContentAsync must accept a streaming progress callback');
assert(client.includes('ReadAsStreamAsync'), 'GenerateContentAsync must read the SSE response stream');
assert(client.includes('ReadLineAsync'), 'GenerateContentAsync must parse SSE lines incrementally');
assert(client.includes('progressHandler.Invoke(chunkText)'), 'GenerateContentAsync must publish streaming outline text chunks');
assert(client.includes('TryExtractMarkdownFromEnvelope(eventPayload, eventMarkdown)'), 'markdown streaming must ignore final JSON result envelopes and keep streamed markdown chunks');
assert(!client.includes('finalMarkdown = ExtractMarkdownFromEnvelope(eventPayload)'), 'markdown streaming must not throw when the final event result is a JSON outline object');
assert(client.includes('/api/ppt/templates'), 'Docmee client must include the template list endpoint for later template selection');
assert(client.includes('/api/ppt/v2/generatePptx'), 'Docmee client must generate PPTX after outline');
assert(client.includes('/api/ppt/downloadPptx'), 'Docmee client must request a downloadable PPT file URL');
assert(client.includes('Public Class DocmeePptInfo'), 'Docmee client must expose generated PPT metadata');
assert(client.includes('GeneratePptxAsync'), 'Docmee client must expose GeneratePptxAsync');
assert(client.includes('.TemplateId = TryGetString(pptInfo("templateId"))'), 'GeneratePptxAsync must return the template id that Docmee actually used');
assert(client.includes('Public Async Function DownloadPptxAsync(pptId As String, Optional refresh As Boolean = False) As Task(Of String)'), 'Docmee client must allow refreshing the PPTX download after template changes');
assert(client.includes('{"refresh", refresh}'), 'downloadPptx must pass the requested refresh value');
assert(client.includes('DownloadPptxFileAsync'), 'Docmee client must download the PPTX file');

assert(pane.includes('Class ThemePptTaskPane'), 'ThemePptTaskPane class must exist');
assert(pane.includes('CreateTaskAsync'), 'ThemePptTaskPane must create a Docmee task');
assert(pane.includes('_outlineMarkdown'), 'ThemePptTaskPane must keep the markdown outline returned by Docmee');
assert(pane.includes('GenerateMarkdownContentAsync(_taskId, AddressOf AppendOutlineStreamText)'), 'ThemePptTaskPane must generate markdown outline content');
assert(pane.includes('AddressOf AppendOutlineStreamText'), 'ThemePptTaskPane must wire streaming outline text into the UI');
assert(pane.includes('AppendOutlineStreamText'), 'ThemePptTaskPane must append streamed outline text');
assert(pane.includes('_outputBox.AppendText'), 'ThemePptTaskPane must display outline chunks while they stream');
assert(pane.includes('BeginInvoke'), 'ThemePptTaskPane must marshal streamed UI updates onto the task pane thread');
assert(pane.includes('_outputBox.Text = _outlineMarkdown.Trim()'), 'ThemePptTaskPane must show the completed markdown outline, not raw JSON');
assert(pane.includes('Private ReadOnly _templateCardPanel As New FlowLayoutPanel()'), 'ThemePptTaskPane must provide an OfficePLUS-like template card panel');
assert(pane.includes('_templateCardPanel.AutoScroll = True'), 'template card panel must support scrolling through template covers');
assert(pane.includes('_templateCardPanel.Visible = False'), 'template card panel must be hidden until templates are ready');
assert(pane.includes('_outputBox.Visible = True'), 'outline output box must be visible while content is streaming');
assert(pane.includes('_outputBox.Visible = False'), 'outline output box must be hidden after content generation completes');
assert(pane.includes('_templateCardPanel.Visible = True'), 'template card panel must replace the outline box after content generation completes');
assert(pane.includes('Private Sub ShowOutlineOutput()'), 'ThemePptTaskPane must centralize showing the streaming outline output');
assert(pane.includes('Private Sub ShowTemplateGallery()'), 'ThemePptTaskPane must centralize showing the template gallery');
assert(pane.includes('ShowOutlineOutput()'), 'ThemePptTaskPane must explicitly show outline output during generation');
assert(pane.includes('ShowTemplateGallery()'), 'ThemePptTaskPane must explicitly switch to template gallery after outline generation');
assert(pane.includes('If _templateCardPanel.Visible Then'), 'task pane diagnostics must not bring the large outline/log box back after gallery mode');
assert(pane.includes('ShowTemplateGallery()') && pane.includes('AppendTaskPaneLine("已导入页数: " & importedCount.ToString())'), 'generate/import flow must return to gallery mode after import diagnostics');
assert(pane.includes('If _templateCombo.Items.Count > 0 Then'), 'ThemePptTaskPane must show template cards only when templates were loaded');
assert(pane.includes('CreateTemplateCard(template)'), 'ThemePptTaskPane must render selectable template cards');
assert(pane.includes('PictureBoxSizeMode.Zoom'), 'template cards must display template cover images');
assert(pane.includes('If Not String.IsNullOrWhiteSpace(template.CoverUrl) Then'), 'template cards must load template cover URLs when available');
assert(pane.includes('Private ReadOnly _templateSelectLabels As New Dictionary(Of String, Label)()'), 'template cards must keep explicit select labels for visible selected state');
assert(pane.includes('BuildTemplateMetaText(template)'), 'template cards must show template metadata even when cover URLs cannot be loaded');
assert(pane.includes('Text = "模板预览"'), 'template cards must show a metadata fallback preview instead of an empty cover area');
assert(pane.includes('selectLabel.Text = If(isSelected, "已选择", "选择模板")'), 'template card selection must update explicit select button text');
assert(pane.includes('Color.FromArgb(234, 88, 12)'), 'selected template button must use a strong selected color');
assert(pane.includes('SelectTemplate(template)'), 'clicking a template card must select that template');
assert(pane.includes('RefreshTemplateSelectionStyles()'), 'template card selection must have a visual selected state');
assert(pane.includes('LoadTemplatesAsync'), 'ThemePptTaskPane must load template choices for the user');
assert(pane.includes('ComboBox'), 'ThemePptTaskPane must provide a template selector');
assert(pane.includes('Dim markdown = _outlineMarkdown.Trim()'), 'ThemePptTaskPane must send the same markdown outline into generatePptx');
assert(pane.includes('AppendTaskPaneLine("使用模板ID: " & selectedTemplate.Id)'), 'ThemePptTaskPane must print the selected template id before generation');
assert(pane.includes('Dim pptTaskId = Await _client.CreateMarkdownTaskAsync(markdown)'), 'ThemePptTaskPane must create a fresh markdown task for the selected template');
assert(pane.includes('AppendTaskPaneLine("生成任务ID: " & pptTaskId)'), 'ThemePptTaskPane must print the final PPT generation task id');
assert(pane.includes('GeneratePptxAsync'), 'ThemePptTaskPane must generate PPT with the selected template');
assert(pane.includes('Dim pptInfo = Await _client.GeneratePptxAsync(pptTaskId, selectedTemplate.Id, markdown)'), 'ThemePptTaskPane must generate PPT from the fresh markdown task');
assert(pane.includes('AppendTaskPaneLine("PPT ID: " & pptInfo.Id)'), 'ThemePptTaskPane must print the generated PPT id');
assert(pane.includes('AppendTaskPaneLine("返回模板ID: " & pptInfo.TemplateId)'), 'ThemePptTaskPane must print the template id returned by Docmee');
assert(pane.includes('String.Equals(pptInfo.TemplateId, selectedTemplate.Id, StringComparison.Ordinal)'), 'ThemePptTaskPane must verify Docmee used the selected template id');
assert(pane.includes('DownloadPptxAsync'), 'ThemePptTaskPane must request a PPT download URL');
assert(pane.includes('DownloadPptxAsync(pptInfo.Id, True)'), 'ThemePptTaskPane must refresh the downloaded PPTX for the generated template');
assert(pane.includes('AppendTaskPaneLine("PPTX 下载地址: " & fileUrl)'), 'ThemePptTaskPane must print the returned PPTX download URL for manual comparison');
assert(pane.includes('AppendTaskPaneLine("本地保存路径: " & localPath)'), 'ThemePptTaskPane must print the local downloaded PPTX path for manual comparison');
assert(pane.includes('DownloadPptxFileAsync'), 'ThemePptTaskPane must download the generated PPT file');
assert(pane.includes('ImportPptxIntoPresentation'), 'ThemePptTaskPane must import the downloaded PPT into the active presentation');
assert(pane.includes('Dim importedCount = ImportPptxIntoPresentation(localPath)'), 'ThemePptTaskPane must know how many slides were imported');
assert(pane.includes('Dim originalSlideIndex = CaptureActiveSlideIndex(target)'), 'ThemePptTaskPane must capture the user active slide before importing generated template slides');
assert(pane.includes('RestoreActiveSlide(target, originalSlideIndex)'), 'ThemePptTaskPane must restore the original active slide after import so PowerPoint color palettes do not stay on the generated template');
assert(pane.includes('AppendTaskPaneLine("已导入页数: " & importedCount.ToString())'), 'ThemePptTaskPane must print the imported slide count');
assert(pane.includes('Throw New InvalidOperationException("PPTX 已下载，但没有成功导入任何幻灯片。")'), 'ThemePptTaskPane must fail loudly if PowerPoint imports zero slides');
assert(pane.includes('AppendTaskPaneLine("生成并导入失败: " & ex.Message)'), 'ThemePptTaskPane must print import errors in the task pane');
assert(pane.includes('Presentations.Open(downloadPath'), 'ThemePptTaskPane must open the generated PPTX before importing slides');
assert(pane.includes('sourceSlide.Copy()'), 'ThemePptTaskPane must copy full generated slides to preserve template elements');
assert(pane.includes('TryPasteSlideWithSourceFormatting(target, sourcePresentation.Slides(slideIndex))'), 'ThemePptTaskPane must use source-formatting paste for generated slides');
assert(pane.includes('target.Windows(1).Activate()'), 'ThemePptTaskPane must activate the destination window before source-formatting paste');
assert(pane.includes('target.Windows(1).View.GotoSlide(target.Slides.Count)'), 'ThemePptTaskPane must move the destination view to the end before source-formatting paste');
assert(pane.includes('_pptApp.CommandBars.ExecuteMso("PasteSourceFormatting")'), 'ThemePptTaskPane must request PowerPoint keep-source-formatting paste');
assert(pane.includes('target.Windows(1).View.GotoSlide(slideIndex)'), 'ThemePptTaskPane must navigate back to the original slide after import');
assert(pane.includes('target.Slides(slideIndex).Select()'), 'ThemePptTaskPane must reselect the original slide after import');
assert(pane.includes('target.Slides.Paste(target.Slides.Count + 1)'), 'ThemePptTaskPane must keep a normal paste fallback');
assert(!pane.includes('InsertFromFile'), 'ThemePptTaskPane must not rely on InsertFromFile for Docmee imports');
assert(pane.includes('FixInsertedSlideReadability'), 'ThemePptTaskPane must improve readability only on newly inserted slides');
assert(pane.includes('importedSlides.Add'), 'ThemePptTaskPane must track exactly pasted slides for readability adjustments');
assert(pane.includes('FixShapeTextReadability'), 'ThemePptTaskPane must adjust low-contrast text on inserted slides');
assert(pane.includes('GetSlideBackgroundLuminance'), 'ThemePptTaskPane must consider slide background luminance before adjusting text');
assert(pane.includes('RGB(45, 52, 64)'), 'ThemePptTaskPane must darken faint text on light generated slides');
assert(pane.includes('InsertOutlineIntoPresentation'), 'ThemePptTaskPane must insert generated outline into PPT');
assert(pane.includes('children'), 'ThemePptTaskPane must consume outline children');
assert(pane.includes('pages'), 'ThemePptTaskPane must consume Docmee pages responses');
assert(pane.includes('overall_theme'), 'ThemePptTaskPane must use Docmee overall_theme as a title candidate');
assert(pane.includes('ppLayoutTitleOnly'), 'ThemePptTaskPane must create title/content slides');

console.log('docmee theme ppt checks passed');
