const fs = require('fs');

function read(path) {
  return fs.readFileSync(path, 'utf8');
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function assertOrder(source, first, second, message) {
  const firstIndex = source.indexOf(first);
  const secondIndex = source.indexOf(second);
  assert(firstIndex !== -1, `${message}: missing ${first}`);
  assert(secondIndex !== -1, `${message}: missing ${second}`);
  assert(firstIndex < secondIndex, message);
}

function methodBlock(source, signature) {
  const start = source.indexOf(signature);
  assert(start !== -1, `missing method: ${signature}`);
  const nextMethod = source.indexOf('\n    Private ', start + signature.length);
  const nextPublicMethod = source.indexOf('\n    Public ', start + signature.length);
  const candidates = [nextMethod, nextPublicMethod].filter((index) => index !== -1);
  const end = candidates.length ? Math.min(...candidates) : source.length;
  return source.slice(start, end);
}

function countOccurrences(source, needle) {
  return source.split(needle).length - 1;
}

const ribbonDesigner = read('PowerPointAi/Ribbon1.Designer.vb');
const ribbon = read('PowerPointAi/Ribbon1.vb');
const addIn = read('PowerPointAi/ThisAddIn.vb');
const project = read('PowerPointAi/PowerPointAi.vbproj');
const powerPointAppConfig = read('PowerPointAi/app.config');
const configSettings = read('ShareRibbon/Config/ConfigSettings.vb');
const installer = fs.existsSync('OfficeAgent/OfficeAgent.vdproj') ? read('OfficeAgent/OfficeAgent.vdproj') : '';
const client = fs.existsSync('PowerPointAi/DocmeePptClient.vb') ? read('PowerPointAi/DocmeePptClient.vb') : '';
const pane = fs.existsSync('PowerPointAi/ThemePptTaskPane.vb') ? read('PowerPointAi/ThemePptTaskPane.vb') : '';
const templateDialog = fs.existsSync('PowerPointAi/TemplateSelectionForm.vb') ? read('PowerPointAi/TemplateSelectionForm.vb') : '';
const apiSmoke = fs.existsSync('scripts/verify-docmee-api-smoke.js') ? read('scripts/verify-docmee-api-smoke.js') : '';

assert(ribbonDesigner.includes('Me.TemplateFormatButton.Visible = True'), '主题生成PPT button must be visible');
assert(ribbonDesigner.includes('Me.TemplateFormatButton.Label = "AI生成PPT"'), 'TemplateFormatButton must be relabeled AI生成PPT');
assert(ribbonDesigner.includes('Me.ProofreadButton.Visible = True'), 'PowerPoint ribbon must expose the replace-single-slide button');
assert(ribbonDesigner.includes('Me.ProofreadButton.Label = "替换单页"'), 'ProofreadButton must be repurposed as 替换单页 for PowerPoint');
assert(ribbonDesigner.includes('Me.ReformatButton.Visible = True'), 'PowerPoint ribbon must expose the beautify-single-slide button');
assert(ribbonDesigner.includes('Me.ReformatButton.Label = "美化单页"'), 'ReformatButton must be labeled 美化单页');
assert(ribbonDesigner.includes('Me.TranslateButton.Label = "文本翻译"'), 'PowerPoint ribbon must expose text translation');
assert(ribbonDesigner.includes('Me.ContinuationButton.Label = "文本优化"'), 'PowerPoint ribbon must expose text optimization');
assert(ribbon.includes('Globals.ThisAddIn.ShowThemePptTaskPane()'), 'Ribbon must open the theme PPT task pane');
assert(ribbon.includes('ReplaceCurrentSlideWithGeneratedTextAsync'), 'Ribbon must implement a replace-single-slide workflow');
assert(ribbon.includes('DeleteOriginalSlideAfterReplacement'), 'replace-single-slide workflow must delete the old page after inserting the replacement');
assert(ribbon.includes('ApplySimpleBeautifyToCurrentSlide'), 'Ribbon must implement a beautify-current-slide workflow');
assert(ribbon.includes('TranslateButton_Click'), 'Ribbon must implement text translation');
assert(ribbon.includes('ContinuationButton_Click'), 'Ribbon must implement text optimization');
assert(ribbon.includes('"扩写"') && ribbon.includes('"精简"') && ribbon.includes('"填充"') && ribbon.includes('"补全文案"'), 'text optimization must support expand, shorten, fill, and complete modes');
assert(ribbon.includes('Case "填充"'), 'text optimization must implement a dedicated fill mode');

assert(addIn.includes('themePptTaskPane'), 'ThisAddIn must keep a theme PPT task pane');
assert(addIn.includes('ShowThemePptTaskPane'), 'ThisAddIn must expose ShowThemePptTaskPane');
assert(addIn.includes('EnsureCoreServicesLoaded()'), 'Theme PPT task pane must load core services before WebView2 use');
assert(addIn.includes('New ThemePptTaskPane(Me.Application)'), 'ThisAddIn must create ThemePptTaskPane');

assert(project.includes('Compile Include="DocmeePptClient.vb"'), 'PowerPoint project must compile DocmeePptClient');
assert(project.includes('Compile Include="ThemePptTaskPane.vb"'), 'PowerPoint project must compile ThemePptTaskPane');
assert(project.includes('Compile Include="TemplateSelectionForm.vb"'), 'PowerPoint project must compile TemplateSelectionForm');
assert(project.includes('<Reference Include="Markdig'), 'PowerPoint project must reference Markdig for real markdown preview rendering');
assert(project.includes('Copy SourceFiles="..\\packages\\Markdig.0.41.1\\lib\\net462\\Markdig.dll"'), 'PowerPoint build must copy Markdig.dll for VSTO runtime');
assert(
  countOccurrences(project, 'ProjectReference Include="..\\ShareRibbon\\ShareRibbon.vbproj"') === 1,
  'PowerPoint project must reference ShareRibbon exactly once to avoid duplicate MSBuild project references'
);

if (installer) {
  assert(installer.includes('"RemovePreviousVersions" = "11:TRUE"'), 'installer must remove previous versions during upgrade');
}

assert(configSettings.includes('DefaultDocmeeApiBaseUrl As String = "https://test.docmee.cn"'), 'Docmee config must retain the demo API fallback URL');
assert(configSettings.includes('DefaultDocmeeToken As String = "ak_demo"'), 'Docmee config must retain the demo token fallback');
assert(configSettings.includes('OfficeAi.DocmeeApiBaseUrl'), 'Docmee API base URL must be configurable through app settings');
assert(configSettings.includes('OfficeAi.DocmeeToken'), 'Docmee token must be configurable through app settings');
assert(configSettings.includes('OFFICE_AI_DOCMEE_API_BASE_URL'), 'Docmee API base URL must be configurable through environment variables');
assert(configSettings.includes('OFFICE_AI_DOCMEE_TOKEN'), 'Docmee token must be configurable through environment variables');
assert(configSettings.includes('DocmeeSettingsFileName As String = "docmee_settings.json"'), 'Docmee settings must have a per-user persisted settings file');
assert(configSettings.includes('Public Shared Sub SaveDocmeeSettings'), 'Docmee settings must be saveable from the plugin UI');
assert(configSettings.includes('Private Shared Function LoadDocmeeUserSettings'), 'Docmee settings must load persisted per-user values');
assert(configSettings.includes('Newtonsoft.Json.JsonConvert'), 'Docmee settings must use structured JSON serialization');
assert(!configSettings.includes('FirstNonEmpty(apiBaseUrl, DefaultDocmeeApiBaseUrl)'), 'Docmee settings save must not persist demo defaults when the plugin dialog leaves the base URL blank');
assert(configSettings.includes('Dim normalizedBaseUrl = If(apiBaseUrl, "").Trim().TrimEnd("/"c)'), 'Docmee settings save must preserve a blank base URL so fallback sources can still apply');
assertOrder(
  configSettings,
  'Environment.GetEnvironmentVariable(DocmeeApiBaseUrlEnvironmentVariable)',
  'SafeAppSetting(DocmeeApiBaseUrlAppSettingKey)',
  'Docmee API base URL environment variable must override app.config demo defaults'
);
assertOrder(
  configSettings,
  'Environment.GetEnvironmentVariable(DocmeeTokenEnvironmentVariable)',
  'SafeAppSetting(DocmeeTokenAppSettingKey)',
  'Docmee token environment variable must override app.config demo defaults'
);
assert(powerPointAppConfig.includes('key="OfficeAi.DocmeeApiBaseUrl"'), 'PowerPoint app.config must expose the Docmee API base URL setting');
assert(powerPointAppConfig.includes('key="OfficeAi.DocmeeToken"'), 'PowerPoint app.config must expose the Docmee token setting');
assert(client.includes('GetConfiguredApiBaseUrl'), 'Docmee client must read its API base URL from configuration');
assert(client.includes('GetConfiguredToken'), 'Docmee client must read its token from configuration');
assert(client.includes('BuildEndpoint("/api/ppt/v2/createTask")'), 'Docmee createTask endpoint must be built from the configured base URL');
assert(client.includes('BuildEndpoint("/api/ppt/v2/generateContent")'), 'Docmee generateContent endpoint must be built from the configured base URL');
assert(client.includes('BuildEndpoint("/api/ppt/v2/generatePptx")'), 'Docmee generatePptx endpoint must be built from the configured base URL');
assert(client.includes('BuildEndpoint("/api/ppt/updatePptTemplate")'), 'Docmee updatePptTemplate endpoint must be built from the configured base URL');
assert(client.includes('AddDocmeeTokenHeader'), 'Docmee client must add the token through a shared helper');
assert(!client.includes('DemoToken'), 'Docmee client must not hard-code the demo token in request code');
assert(client.includes('Public Async Function CreateMarkdownTaskAsync(markdown As String) As Task(Of String)'), 'Docmee client must create fresh markdown tasks for PPT generation');
assert(client.includes('Return Await CreateTaskAsync(markdown, "7")'), 'Docmee markdown task creation must use type=7');
assert(client.includes('Public Async Function CreateFileTaskAsync(filePath As String) As Task(Of String)'), 'Docmee client must create upload-file tasks for document-to-PPT');
assert(client.includes('GetTaskTypeForFile(filePath)'), 'Docmee client must choose the Docmee task type from the uploaded document');
assert(methodBlock(client, 'Private Shared Function GetTaskTypeForFile(filePath As String) As String').includes('Case ".xmind", ".mm"'), 'Docmee mind map uploads must be detected separately');
assert(methodBlock(client, 'Private Shared Function GetTaskTypeForFile(filePath As String) As String').includes('Return "3"'), 'Docmee mind map uploads must use type=3');
assert(methodBlock(client, 'Private Shared Function GetTaskTypeForFile(filePath As String) As String').includes('Return "2"'), 'Docmee document-to-editable-markdown flow must use type=2 upload-file tasks');
assert(!client.includes('Return "4"'), 'Docmee document-to-editable-markdown flow must not use type=4 because generateContent prompt is ignored for type=4');
assert(client.includes('Private Shared Function IsSupportedUploadFile(filePath As String) As Boolean'), 'Docmee upload flow must validate supported document extensions before upload');
assert(client.includes('NotSupportedException'), 'Docmee upload flow must show a local unsupported-format error');
assert(client.includes('".ppt", ".pptx"'), 'Docmee upload validation must support PPT/PPTX files');
assert(client.includes('".xls", ".xlsx", ".csv"'), 'Docmee upload validation must support spreadsheet inputs');
assert(client.includes('".html", ".epub", ".mobi"'), 'Docmee upload validation must support HTML and ebook inputs');
assert(client.includes('".xmind", ".mm"'), 'Docmee upload validation must support mind map inputs');
assert(client.includes('Using fileStream As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)'), 'Docmee upload tasks must stream the selected file');
assert(client.includes('Private Shared ReadOnly Property UpdatePptTemplateEndpoint As String'), 'Docmee client must define updatePptTemplate endpoint');
assert(client.includes('Public Async Function UpdatePptTemplateAsync(pptId As String, templateId As String'), 'Docmee client must expose template replacement for generated PPTs');
assert(client.includes('{"sync", sync}'), 'Docmee template replacement must send the sync option');
assert(client.includes('GenerateMarkdownContentAsync'), 'Docmee client must expose markdown outline generation');
assert(client.includes('"outlineType", "MD"'), 'generateContent must request markdown outline for the task pane');
assert(client.includes('NormalizeDocmeePrompt'), 'Docmee generateContent prompt must be normalized through a shared helper');
assert(client.includes('If normalized.Length > 49 Then normalized = normalized.Substring(0, 49)'), 'Docmee prompt must be capped below the documented 50-character limit');
assert(client.includes('ReadAsStreamAsync'), 'GenerateContentAsync must read the SSE response stream');
assert(client.includes('ReadLineAsync'), 'GenerateContentAsync must parse SSE lines incrementally');
assert(client.includes('progressHandler.Invoke(chunkText)'), 'GenerateContentAsync must publish streaming outline text chunks');
assert(client.includes('TryExtractMarkdownFromEnvelope(eventPayload, eventMarkdown)'), 'markdown streaming must keep streamed markdown when final result is JSON');
assert(methodBlock(client, 'Private Async Function CreateTaskAsync(content As String, taskType As String) As Task(Of String)').includes('Await client.PostAsync(CreateTaskEndpoint, form).ConfigureAwait(False)'), 'Docmee text task creation must avoid resuming HTTP work on the Office UI synchronization context');
assert(methodBlock(client, 'Private Async Function CreateFileTaskAsync(filePath As String, taskType As String) As Task(Of String)').includes('Await client.PostAsync(CreateTaskEndpoint, form).ConfigureAwait(False)'), 'Docmee file upload task creation must avoid resuming HTTP work on the Office UI synchronization context');
assert(methodBlock(client, 'Public Async Function GenerateContentAsync(taskId As String').includes('Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)'), 'Docmee JSON outline generation must not capture the Office UI synchronization context during SSE connect');
assert(methodBlock(client, 'Public Async Function GenerateMarkdownContentAsync(taskId As String').includes('Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)'), 'Docmee Markdown outline generation must not capture the Office UI synchronization context during SSE connect');
assert(methodBlock(client, 'Private Shared Async Function ReadGeneratedMarkdownStreamAsync').includes('Await reader.ReadLineAsync().ConfigureAwait(False)'), 'Docmee Markdown SSE line reading must stay off the Office UI synchronization context');
assert(methodBlock(client, 'Private Shared Async Function ReadGeneratedOutlineStreamAsync').includes('Await reader.ReadLineAsync().ConfigureAwait(False)'), 'Docmee JSON SSE line reading must stay off the Office UI synchronization context');
assert(methodBlock(client, 'Public Async Function GeneratePptxAsync(taskId As String').includes('Await client.SendAsync(request).ConfigureAwait(False)'), 'Docmee PPTX generation must avoid resuming HTTP work on the Office UI synchronization context');
assert(methodBlock(client, 'Public Async Function DownloadPptxAsync(pptId As String').includes('Await client.SendAsync(request).ConfigureAwait(False)'), 'Docmee download URL lookup must avoid resuming HTTP work on the Office UI synchronization context');
assert(methodBlock(client, 'Public Async Function UpdatePptTemplateAsync(pptId As String').includes('Await client.SendAsync(request).ConfigureAwait(False)'), 'Docmee template replacement must avoid resuming HTTP work on the Office UI synchronization context');
assert(methodBlock(client, 'Public Async Function DownloadPptxFileAsync(fileUrl As String').includes('Await client.GetByteArrayAsync(fileUrl).ConfigureAwait(False)'), 'Docmee PPTX byte download must avoid resuming HTTP work on the Office UI synchronization context');
assert(client.includes('Optional cancellationToken As Threading.CancellationToken = Nothing'), 'template and cover calls must accept cancellation tokens');
assert(client.includes('Await client.SendAsync(request, cancellationToken).ConfigureAwait(False)'), 'template list HTTP calls must not capture the Office UI synchronization context');
assert(client.includes('DownloadTemplateCoverAsync(coverUrl As String, Optional cancellationToken As Threading.CancellationToken = Nothing)'), 'template cover downloads must accept cancellation');
assert(client.includes('HttpClientHandler() With {.AllowAutoRedirect = False}'), 'template cover downloads must handle redirects explicitly');
assert(client.includes('AddDocmeeTokenHeader(request.Headers)'), 'Docmee HTTP requests must send the configured token');
assert(client.includes('GeneratePptxAsync'), 'Docmee client must expose GeneratePptxAsync');
assert(client.includes('DownloadPptxAsync(pptId As String, Optional refresh As Boolean = False)'), 'Docmee client must allow refreshing PPTX downloads');
assert(client.includes('ValidatePptxBytes(bytes)'), 'Docmee PPTX downloads must validate bytes before writing them to disk');
assert(client.includes('bytes(0) <> &H50') && client.includes('bytes(1) <> &H4B'), 'Docmee PPTX validation must reject non-ZIP/PPTX responses');
assert(apiSmoke.includes('files = []'), 'Docmee API smoke must support multipart file parts for document-to-PPT verification');
assert(apiSmoke.includes('filename="${file.filename}"'), 'Docmee API smoke must send uploaded document filenames');
assert(apiSmoke.includes('verifyTitleOutline'), 'Docmee API smoke must exercise the title-to-editable-outline path');
assert(apiSmoke.includes("type: '1'"), 'Docmee API smoke must create a type=1 title task');
assert(apiSmoke.includes("type: '2'"), 'Docmee API smoke must exercise the type=2 uploaded document task path');
assert(apiSmoke.includes('/api/ppt/v2/generateContent'), 'Docmee API smoke must generate an editable outline from uploaded documents');
assert(apiSmoke.includes("outlineType: 'MD'"), 'Docmee API smoke must request a Markdown outline from document uploads');
assert(apiSmoke.includes('titleGeneration'), 'Docmee API smoke output must report title generation verification');
assert(apiSmoke.includes('documentUpload'), 'Docmee API smoke output must report document upload verification');

assert(templateDialog.includes('Public Class TemplateSelectionForm'), 'template preview dialog must exist');
assert(templateDialog.includes('WebView2'), 'template preview dialog must render with WebView2');
assert(templateDialog.includes('window.chrome.webview.postMessage({type:\'select\''), 'template preview dialog must post selected template id');

assert(pane.includes('Imports Markdig'), 'ThemePptTaskPane must use Markdig for markdown-to-HTML rendering');
assert(pane.includes('Class ThemePptTaskPane'), 'ThemePptTaskPane class must exist');
assert(pane.includes('Private ReadOnly _outlineEditor As New TextBox()'), 'ThemePptTaskPane must provide a markdown editor');
assert(pane.includes('Private ReadOnly _outlinePreviewWebView As New WebView2()'), 'ThemePptTaskPane must provide a markdown preview WebView2');
assert(pane.includes('Private ReadOnly _generationModeCombo As New ComboBox()'), 'ThemePptTaskPane must allow title/document/markdown generation modes');
assert(pane.includes('GenerationModeTitle As String = "标题生成"'), 'ThemePptTaskPane must support title-to-PPT mode');
assert(pane.includes('GenerationModeDocument As String = "文档生成"'), 'ThemePptTaskPane must support document-to-PPT mode');
assert(pane.includes('GenerationModeMarkdown As String = "Markdown大纲"'), 'ThemePptTaskPane must support pasted markdown outline mode');
assert(pane.includes('Private ReadOnly _documentPathBox As New TextBox()'), 'ThemePptTaskPane must display the selected document path');
assert(pane.includes('Private ReadOnly _chooseDocumentButton As New Button()'), 'ThemePptTaskPane must let users choose a source document');
assert(pane.includes('Private ReadOnly _changeThemeButton As New Button()'), 'ThemePptTaskPane must expose a one-click theme replacement action');
assert(pane.includes('Private ReadOnly _applyLocalThemeButton As New Button()'), 'ThemePptTaskPane must expose a one-click local current-presentation theme action');
assert(pane.includes('Private ReadOnly _configureDocmeeButton As New Button()'), 'ThemePptTaskPane must expose Docmee configuration in the plugin UI');
assert(pane.includes('Private _lastGeneratedPptId As String'), 'ThemePptTaskPane must remember the latest Docmee PPT id for theme replacement');
assert(pane.includes('Private _lastImportedSlideStartIndex As Integer'), 'ThemePptTaskPane must remember where generated slides were imported');
assert(pane.includes('Private _lastImportedSlideCount As Integer'), 'ThemePptTaskPane must remember how many generated slides were imported');
assert(pane.includes('Private ReadOnly _outlinePreviewDebounceTimer As New System.Windows.Forms.Timer()'), 'markdown preview debounce timer must explicitly use WinForms Timer to avoid System.Threading.Timer ambiguity');
assert(pane.includes('Private _markdownPreviewRenderGeneration As Integer'), 'markdown preview must ignore stale background renders');
assert(pane.includes('Private ReadOnly _finishOutlineEditButton As New Button()'), 'ThemePptTaskPane must provide an explicit finish-edit action');
assert(pane.includes('Private _isOutlineEditCompleted As Boolean'), 'ThemePptTaskPane must track whether the user finished editing');
assert(pane.includes('_outlineEditor.ReadOnly = False'), 'markdown editor must be editable by the user');
assert(pane.includes('AddHandler _outlineEditor.TextChanged, AddressOf OutlineEditor_TextChanged'), 'markdown preview must react to editor changes');
assert(pane.includes('AddHandler _finishOutlineEditButton.Click, AddressOf FinishOutlineEditButton_Click'), 'finish-edit button must be wired');
assert(pane.includes('AddHandler _chooseDocumentButton.Click, AddressOf ChooseDocumentButton_Click'), 'document chooser button must be wired');
assert(pane.includes('AddHandler _changeThemeButton.Click, AddressOf ChangeThemeButton_Click'), 'change-theme button must be wired');
assert(pane.includes('AddHandler _applyLocalThemeButton.Click, AddressOf ApplyLocalThemeButton_Click'), 'local theme button must be wired');
assert(pane.includes('AddHandler _configureDocmeeButton.Click, AddressOf ConfigureDocmeeButton_Click'), 'Docmee configuration button must be wired');
assert(pane.includes('Private Function GetSelectedGenerationMode() As String'), 'ThemePptTaskPane must centralize selected generation mode');
assert(pane.includes('Select Case mode'), 'ThemePptTaskPane must branch generation by selected source mode');
assert(pane.includes('Private Async Function GenerateOutlineFromDocumentAsync() As Task(Of String)'), 'ThemePptTaskPane must generate editable markdown from uploaded documents');
assert(pane.includes('*.ppt;*.pptx'), 'document chooser must expose PPT/PPTX files supported by Docmee');
assert(pane.includes('*.xls;*.xlsx;*.csv'), 'document chooser must expose spreadsheet files supported by Docmee');
assert(pane.includes('*.html;*.epub;*.mobi'), 'document chooser must expose HTML and ebook files supported by Docmee');
assert(pane.includes('*.xmind;*.mm'), 'document chooser must expose mind map files supported by Docmee');
assert(pane.includes('Private Function PrepareMarkdownOutlineFromInput() As String'), 'ThemePptTaskPane must support pasted markdown outlines');
assert(pane.includes('GetDocumentPrompt()'), 'document generation must allow an optional user prompt');
assert(pane.includes('Private Sub UpdateMarkdownPreview()'), 'ThemePptTaskPane must centralize markdown preview rendering');
assert(pane.includes('Private Sub ScheduleMarkdownPreviewUpdate'), 'ThemePptTaskPane must schedule markdown preview updates instead of rendering on every keystroke');
assert(pane.includes('Private Function BeginInvokeIfAlive(action As MethodInvoker) As Boolean'), 'ThemePptTaskPane must centralize safe UI marshaling for background callbacks');
assert(countOccurrences(pane, '.BeginInvoke(') === 1, 'ThemePptTaskPane must route raw BeginInvoke calls through BeginInvokeIfAlive only');
assert(pane.includes('Private Async Sub OutlinePreviewDebounceTimer_Tick'), 'ThemePptTaskPane must render markdown after a debounce interval');
assert(pane.includes('Await Task.Run(Function() BuildMarkdownPreviewHtml(markdownSnapshot))'), 'markdown-to-HTML rendering must run off the Office UI thread');
assert(pane.includes('_markdownPreviewRenderGeneration += 1'), 'scheduling a new markdown preview must invalidate any in-flight render immediately');
assert(pane.includes('If renderGeneration <> _markdownPreviewRenderGeneration Then Return'), 'stale markdown preview renders must be ignored');
assert(pane.includes('Markdig.Markdown.ToHtml'), 'ThemePptTaskPane must render markdown using Markdig, not a plain TextBox');
assert(pane.includes('UseAdvancedExtensions().DisableHtml().Build()'), 'markdown preview must support rich markdown while disabling raw HTML');
assert(pane.includes('_outlinePreviewWebView.NavigateToString'), 'ThemePptTaskPane must display rendered markdown in WebView2');
assert(pane.includes('Private Function GetEditedMarkdown() As String'), 'ThemePptTaskPane must read the current edited markdown');
assert(pane.includes('SetOutlineEditorText(_outlineMarkdown.Trim())'), 'generated markdown must be placed into the editable markdown editor');
assert(pane.includes('MarkOutlineEditingRequired()'), 'after generation, the user must edit/confirm before template selection');
const appendOutlineStreamBlock = methodBlock(pane, 'Private Sub AppendOutlineStreamText(chunkText As String)');
assert(appendOutlineStreamBlock.includes('If Me.IsDisposed OrElse _outputBox.IsDisposed Then Return'), 'streaming outline UI callback must ignore late chunks after the task pane is disposed');
assert(appendOutlineStreamBlock.includes('If _outputBox.InvokeRequired Then'), 'streaming outline UI callback must marshal background chunks to the UI thread');
assert(appendOutlineStreamBlock.includes('BeginInvokeIfAlive(CType(Sub() AppendOutlineStreamText(chunkText), MethodInvoker))'), 'streaming outline UI callback must use the safe UI marshaling helper');
assert(pane.includes('Private Sub MarkOutlineEditingRequired()'), 'ThemePptTaskPane must centralize the editing-required state');
assert(pane.includes('FinishOutlineEditButton_Click(sender As Object, e As EventArgs)'), 'ThemePptTaskPane must implement finish-edit flow');
assert(pane.includes('Private Sub ApplyOutlineEditCompletion()'), 'ThemePptTaskPane must centralize completed-edit state changes');
assert(pane.includes('Private Sub ClearGeneratedPptState()'), 'ThemePptTaskPane must centralize clearing generated PPT state');
assert(pane.includes('ClearGeneratedPptState()'), 'ThemePptTaskPane must clear stale generated PPT state when the outline changes or generation fails');
assert(pane.includes('Private Sub ValidateEditedMarkdownForDocmee(markdown As String)'), 'ThemePptTaskPane must validate edited markdown against Docmee outline requirements');
assert(pane.includes('ValidateEditedMarkdownForDocmee(GetEditedMarkdown())'), 'finish-edit flow must validate the edited markdown before template selection');
assert(pane.includes('ValidateEditedMarkdownForDocmee(markdown)'), 'PPT generation must revalidate edited markdown before sending it to Docmee');
assert(pane.includes('headingLevel = 1') && pane.includes('headingLevel = 2'), 'Docmee markdown validation must check required level-1 and level-2 headings');
assert(pane.includes('If Not _isOutlineEditCompleted Then'), 'template selection and generation must guard against unfinished edits');
assert(pane.includes('请先完成 Markdown 大纲编辑'), 'unfinished edits must show a clear user-facing message');
assert(pane.includes('_selectTemplateButton.Enabled = CanChooseTemplate()'), 'template preview must only be enabled when editing is complete and templates exist');
assert(pane.includes('_insertButton.Enabled = CanGenerateFromTemplate()'), 'generation must only be enabled after editing is complete and a template is selected');
assert(pane.includes('Dim markdown = GetEditedMarkdown()'), 'PPT generation must use the user-edited markdown');
assert(!pane.includes('Dim markdown = _outlineMarkdown.Trim()'), 'PPT generation must not use the stale original markdown');
assert(pane.includes('CreateMarkdownTaskAsync(markdown)'), 'ThemePptTaskPane must create a fresh markdown task from edited markdown');
assert(pane.includes('GeneratePptxAsync(pptTaskId, selectedTemplate.Id, markdown)'), 'ThemePptTaskPane must generate PPT from edited markdown and selected template');
assert(pane.includes('_lastGeneratedPptId = pptInfo.Id'), 'ThemePptTaskPane must retain generated PPT id after generation');
assert(pane.includes('CaptureImportedSlideRange(importedCount)'), 'ThemePptTaskPane must record the imported slide range');
assertOrder(
  methodBlock(pane, 'Private Async Function GenerateAndImportPptxAsync() As Task'),
  'ClearGeneratedPptState()',
  'Dim pptTaskId = Await _client.CreateMarkdownTaskAsync(markdown)',
  'starting a fresh template generation must clear stale PPT/theme state before creating the new Docmee markdown task'
);
assert(pane.includes('Private Async Function ChangeThemeForLatestPptAsync() As Task'), 'ThemePptTaskPane must implement the change-theme action');
assert(pane.includes('Await _client.UpdatePptTemplateAsync(_lastGeneratedPptId, selectedTemplate.Id'), 'change-theme action must call Docmee updatePptTemplate');
assert(pane.includes('ReplaceImportedSlideRange(localPath)'), 'change-theme action must replace the previous imported slide range in Office');
assert(pane.includes('Uri.EscapeDataString(DocmeePptClient.GetConfiguredToken())'), 'template cover URLs must use the configured Docmee token');
assert(pane.includes('Private Function ApplyLocalThemeToCurrentPresentation() As Integer'), 'ThemePptTaskPane must implement one-click local theming for the current presentation');
assert(pane.includes('Private Function ApplyLocalThemeToSlide(slide As PowerPoint.Slide'), 'local theming must apply to each slide');
assert(pane.includes('wenduoduoAI_LocalThemeAccent'), 'local theming must mark reusable accent shapes to avoid duplicates');
assert(pane.includes('For slideIndex As Integer = 1 To presentation.Slides.Count'), 'local theming must iterate every slide in the active presentation');
assert(pane.includes('Private Sub ShowDocmeeSettingsDialog()'), 'ThemePptTaskPane must provide a Docmee settings dialog');
assert(pane.includes('ShareRibbon.ConfigSettings.SaveDocmeeSettings'), 'Docmee settings dialog must persist user-provided settings');
assert(pane.includes('DocmeePptClient.GetConfiguredApiBaseUrl()'), 'Docmee settings dialog must show the currently configured API base URL');
assert(pane.includes('DocmeePptClient.GetConfiguredToken()'), 'Docmee settings dialog must show the currently configured token');

console.log('docmee theme ppt checks passed');
