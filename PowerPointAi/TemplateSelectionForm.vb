Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

Public Class TemplateSelectionForm
    Inherits Form

    Private ReadOnly _templates As New List(Of DocmeeTemplateInfo)()
    Private ReadOnly _selectedTemplateId As String
    Private ReadOnly _coverUrlBuilder As Func(Of String, String)
    Private ReadOnly _browser As New WebView2()
    Private _selectedTemplate As DocmeeTemplateInfo
    Private _initialized As Boolean

    Public Sub New(templates As IEnumerable(Of DocmeeTemplateInfo),
                   selectedTemplateId As String,
                   coverUrlBuilder As Func(Of String, String))
        If templates IsNot Nothing Then
            For Each template In templates
                If template IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(template.Id) Then
                    _templates.Add(template)
                End If
            Next
        End If

        _selectedTemplateId = If(selectedTemplateId, "")
        _coverUrlBuilder = coverUrlBuilder

        Me.Text = "选择模板"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Size = New Size(1040, 760)
        Me.MinimumSize = New Size(760, 540)
        Me.ShowInTaskbar = False

        _browser.Dock = DockStyle.Fill
        _browser.DefaultBackgroundColor = Color.White
        Me.Controls.Add(_browser)

        AddHandler Me.Shown, AddressOf TemplateSelectionForm_Shown
    End Sub

    Public ReadOnly Property SelectedTemplate As DocmeeTemplateInfo
        Get
            Return _selectedTemplate
        End Get
    End Property

    Private Async Sub TemplateSelectionForm_Shown(sender As Object, e As EventArgs)
        If _initialized Then Return
        _initialized = True

        Try
            Await InitializeBrowserAsync()
        Catch ex As Exception
            ThemePptTaskPane.AppendThemePptLog("Template dialog initialize failed: " & ex.ToString())
            MessageBox.Show("模板预览窗口初始化失败：" & ex.Message, "主题生成PPT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Async Function InitializeBrowserAsync() As Task
        ThemePptTaskPane.AppendThemePptLog("Template dialog WebView2 initialize start. count=" & _templates.Count.ToString())
        WebView2Loader.EnsureWebView2Loader()

        Dim userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                          "OfficeAiThemeTemplateDialogWebView2")
        Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)
        Await _browser.EnsureCoreWebView2Async(env)

        If _browser.CoreWebView2 Is Nothing Then
            Throw New InvalidOperationException("CoreWebView2 不可用")
        End If

        _browser.CoreWebView2.Settings.IsScriptEnabled = True
        _browser.CoreWebView2.Settings.IsWebMessageEnabled = True
        _browser.CoreWebView2.Settings.AreDevToolsEnabled = True
        AddHandler _browser.CoreWebView2.WebMessageReceived, AddressOf Browser_WebMessageReceived

        _browser.NavigateToString(BuildHtml())
        ThemePptTaskPane.AppendThemePptLog("Template dialog WebView2 navigate string.")
    End Function

    Private Sub Browser_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim payload = JObject.Parse(e.WebMessageAsJson)
            Dim messageType = TryGetString(payload("type"))

            If String.Equals(messageType, "cancel", StringComparison.Ordinal) Then
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
                Return
            End If

            If Not String.Equals(messageType, "select", StringComparison.Ordinal) Then Return

            Dim templateId = TryGetString(payload("id"))
            If String.IsNullOrWhiteSpace(templateId) Then Return

            For Each template In _templates
                If String.Equals(template.Id, templateId, StringComparison.Ordinal) Then
                    _selectedTemplate = template
                    ThemePptTaskPane.AppendThemePptLog("Template dialog selected: " & template.Id)
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                    Return
                End If
            Next
        Catch ex As Exception
            ThemePptTaskPane.AppendThemePptLog("Template dialog message failed: " & ex.ToString())
        End Try
    End Sub

    Private Function BuildHtml() As String
        Dim builder As New StringBuilder()
        Dim selectedId = _selectedTemplateId

        builder.AppendLine("<!doctype html>")
        builder.AppendLine("<html lang=""zh-CN""><head><meta charset=""utf-8"">")
        builder.AppendLine("<meta name=""viewport"" content=""width=device-width,initial-scale=1"">")
        builder.AppendLine("<style>")
        builder.AppendLine(":root{font-family:'Microsoft YaHei UI','Segoe UI',Arial,sans-serif;color:#1f2937;background:#f7f8fa;}")
        builder.AppendLine("*{box-sizing:border-box}body{margin:0;background:#f7f8fa}.top{position:sticky;top:0;z-index:2;background:#fff;border-bottom:1px solid #e5e7eb;padding:14px 18px;display:flex;align-items:center;gap:12px}.title{font-size:18px;font-weight:700}.sub{font-size:12px;color:#6b7280}.spacer{flex:1}.cancel{border:1px solid #d1d5db;background:#fff;border-radius:4px;height:30px;padding:0 12px;cursor:pointer}")
        builder.AppendLine(".grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px;padding:16px}.card{background:#fff;border:1px solid #dfe3ea;border-radius:6px;overflow:hidden;cursor:pointer;transition:border-color .12s,box-shadow .12s}.card:hover{border-color:#fb923c;box-shadow:0 8px 22px rgba(31,41,55,.12)}.card.selected{border:2px solid #f97316}.thumb{aspect-ratio:16/9;background:#fff7ed;border-bottom:1px solid #edf0f5;display:flex;align-items:center;justify-content:center;position:relative}.thumb img{width:100%;height:100%;object-fit:contain;display:block;background:#fff}.fallback{padding:14px;text-align:center;color:#6b7280;font-size:13px}.body{padding:10px 12px 12px}.name{font-weight:700;font-size:14px;line-height:20px;min-height:40px}.meta{font-size:12px;color:#6b7280;margin-top:4px;min-height:18px}.action{margin-top:10px;height:28px;border-radius:4px;background:#f97316;color:#fff;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700}.card.selected .action{background:#16a34a}.error{position:absolute;inset:0;padding:16px;display:none;align-items:center;justify-content:center;text-align:center;color:#9a3412;background:#fff7ed;font-size:12px}.thumb.failed .error{display:flex}.thumb.failed img{display:none}")
        builder.AppendLine("</style></head><body>")
        builder.Append("<div class=""top""><div><div class=""title"">模板预览</div><div id=""status"" class=""sub"">正在预加载 ")
        builder.Append(_templates.Count.ToString())
        builder.AppendLine(" 个模板封面...</div></div><div class=""spacer""></div><button class=""cancel"" data-action=""cancel"">关闭</button></div>")
        builder.AppendLine("<main class=""grid"">")

        For Each template In _templates
            Dim id = If(template.Id, "")
            Dim title = If(String.IsNullOrWhiteSpace(template.Name), id, template.Name.Trim())
            Dim meta = BuildTemplateMetaText(template)
            Dim coverUrl = BuildCoverUrl(template.CoverUrl)
            Dim isSelected = String.Equals(id, selectedId, StringComparison.Ordinal)

            builder.Append("<article class=""card")
            If isSelected Then builder.Append(" selected")
            builder.Append(""" data-id=""").Append(EscapeHtmlAttribute(id)).AppendLine(""">")
            builder.Append("<div class=""thumb"">")
            If String.IsNullOrWhiteSpace(coverUrl) Then
                builder.Append("<div class=""fallback"">模板没有返回封面地址</div>")
            Else
                builder.Append("<img alt=""").Append(EscapeHtmlAttribute(title)).Append(""" data-src=""").Append(EscapeHtmlAttribute(coverUrl)).Append(""">")
                builder.Append("<div class=""error"">封面加载失败</div>")
            End If
            builder.AppendLine("</div>")
            builder.Append("<div class=""body""><div class=""name"">").Append(EscapeHtml(title)).AppendLine("</div>")
            builder.Append("<div class=""meta"">").Append(EscapeHtml(meta)).AppendLine("</div>")
            builder.Append("<div class=""action"">").Append(If(isSelected, "已选择", "选择模板")).AppendLine("</div></div>")
            builder.AppendLine("</article>")
        Next

        builder.AppendLine("</main>")
        builder.AppendLine("<script>")
        builder.AppendLine("(function(){")
        builder.AppendLine("const status=document.getElementById('status');const imgs=[...document.querySelectorAll('img[data-src]')];let done=0,fail=0;function paint(){if(!imgs.length){status.textContent='没有可加载的封面图';return;}status.textContent='封面预加载 '+done+'/'+imgs.length+(fail?'，失败 '+fail:'');}")
        builder.AppendLine("imgs.forEach(img=>{const thumb=img.closest('.thumb');img.onload=()=>{done++;paint();};img.onerror=()=>{done++;fail++;if(thumb)thumb.classList.add('failed');paint();};img.src=img.dataset.src;});paint();")
        builder.AppendLine("document.addEventListener('click',e=>{if(e.target.closest('[data-action=""cancel""]')){window.chrome.webview.postMessage({type:'cancel'});return;}const card=e.target.closest('.card');if(!card)return;window.chrome.webview.postMessage({type:'select',id:card.dataset.id});});")
        builder.AppendLine("})();")
        builder.AppendLine("</script></body></html>")
        Return builder.ToString()
    End Function

    Private Function BuildCoverUrl(coverUrl As String) As String
        If String.IsNullOrWhiteSpace(coverUrl) Then Return ""
        If _coverUrlBuilder Is Nothing Then Return coverUrl.Trim()
        Return _coverUrlBuilder(coverUrl)
    End Function

    Private Shared Function BuildTemplateMetaText(template As DocmeeTemplateInfo) As String
        Dim parts As New List(Of String)()
        If template IsNot Nothing Then
            If Not String.IsNullOrWhiteSpace(template.Category) Then parts.Add(template.Category.Trim())
            If Not String.IsNullOrWhiteSpace(template.Style) Then parts.Add(template.Style.Trim())
        End If

        If parts.Count = 0 Then Return "Docmee 模板"
        Return String.Join(" / ", parts)
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
End Class
