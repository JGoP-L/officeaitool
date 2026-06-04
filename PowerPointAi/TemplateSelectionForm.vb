Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Threading
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
    Private ReadOnly _uiThreadId As Integer
    Private _selectedTemplate As DocmeeTemplateInfo
    Private _initialized As Boolean

    Public Sub New(templates As IEnumerable(Of DocmeeTemplateInfo),
                   selectedTemplateId As String,
                   coverUrlBuilder As Func(Of String, String))
        _uiThreadId = Thread.CurrentThread.ManagedThreadId

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
        If Not IsOnFormUiThread() Then
            BeginInvokeIfAlive(CType(Sub() TemplateSelectionForm_Shown(sender, e), MethodInvoker))
            Return
        End If

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
        Await EnsureBrowserCoreWebViewOnUiThreadAsync(env)
        Await RunOnUiThreadAsync(
            Sub()
                If _browser.CoreWebView2 Is Nothing Then
                    Throw New InvalidOperationException("CoreWebView2 不可用")
                End If

                _browser.CoreWebView2.Settings.IsScriptEnabled = True
                _browser.CoreWebView2.Settings.IsWebMessageEnabled = True
                _browser.CoreWebView2.Settings.AreDevToolsEnabled = True
                AddHandler _browser.CoreWebView2.WebMessageReceived, AddressOf Browser_WebMessageReceived

                _browser.NavigateToString(BuildHtml())
            End Sub)
        ThemePptTaskPane.AppendThemePptLog("Template dialog WebView2 navigate string.")
    End Function

    Private Sub Browser_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Dim messageJson = e.WebMessageAsJson
        If Not IsOnFormUiThread() Then
            BeginInvokeIfAlive(CType(Sub() HandleBrowserMessage(messageJson), MethodInvoker))
            Return
        End If

        HandleBrowserMessage(messageJson)
    End Sub

    Private Sub HandleBrowserMessage(messageJson As String)
        Try
            Dim payload = JObject.Parse(messageJson)
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

    Private Function IsOnFormUiThread() As Boolean
        Return Thread.CurrentThread.ManagedThreadId = _uiThreadId AndAlso Not Me.InvokeRequired
    End Function

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

    Private Function RunOnUiThreadAsync(action As Action) As Task
        If IsOnFormUiThread() Then
            action()
            Return Task.FromResult(True)
        End If

        Dim tcs As New TaskCompletionSource(Of Boolean)()
        If Not BeginInvokeIfAlive(CType(Sub()
                                            Try
                                                action()
                                                tcs.TrySetResult(True)
                                            Catch ex As Exception
                                                tcs.TrySetException(ex)
                                            End Try
                                        End Sub, MethodInvoker)) Then
            tcs.TrySetException(New ObjectDisposedException(Me.GetType().Name))
        End If

        Return tcs.Task
    End Function

    Private Function EnsureBrowserCoreWebViewOnUiThreadAsync(env As CoreWebView2Environment) As Task
        If IsOnFormUiThread() Then
            Return _browser.EnsureCoreWebView2Async(env)
        End If

        Dim tcs As New TaskCompletionSource(Of Boolean)()
        If Not BeginInvokeIfAlive(CType(Sub()
                                            Try
                                                Dim initTask = _browser.EnsureCoreWebView2Async(env)
                                                initTask.ContinueWith(
                                                    Sub(task)
                                                        If task.IsFaulted AndAlso task.Exception IsNot Nothing Then
                                                            tcs.TrySetException(task.Exception.InnerExceptions)
                                                        ElseIf task.IsCanceled Then
                                                            tcs.TrySetCanceled()
                                                        Else
                                                            tcs.TrySetResult(True)
                                                        End If
                                                    End Sub)
                                            Catch ex As Exception
                                                tcs.TrySetException(ex)
                                            End Try
                                        End Sub, MethodInvoker)) Then
            tcs.TrySetException(New ObjectDisposedException(Me.GetType().Name))
        End If

        Return tcs.Task
    End Function

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
                builder.Append("<img alt=""").Append(EscapeHtmlAttribute(title)).Append(""" data-lazy-src=""").Append(EscapeHtmlAttribute(coverUrl)).Append(""" loading=""lazy"" decoding=""async"">")
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
        builder.AppendLine("const lazyStatus=document.getElementById('status');const lazyImgs=[...document.querySelectorAll('img[data-lazy-src]')];let lazyDone=0,lazyFail=0,lazyActive=0;const lazyMaxActive=4;const lazyQueue=[];const lazyQueued=new WeakSet();")
        builder.AppendLine("function lazyPaint(){if(!lazyImgs.length){lazyStatus.textContent='\u6ca1\u6709\u53ef\u52a0\u8f7d\u7684\u5c01\u9762\u56fe';return;}lazyStatus.textContent='\u5c01\u9762\u52a0\u8f7d '+lazyDone+'/'+lazyImgs.length+(lazyFail?'\uff0c\u5931\u8d25 '+lazyFail:'');}")
        builder.AppendLine("function lazyEnqueue(img){if(!img||lazyQueued.has(img)||img.dataset.loading==='1'||img.dataset.loaded==='1')return;lazyQueued.add(img);lazyQueue.push(img);lazyPump();}")
        builder.AppendLine("function lazyFinish(img,ok){lazyActive=Math.max(0,lazyActive-1);lazyDone++;img.dataset.loaded='1';if(!ok){lazyFail++;const thumb=img.closest('.thumb');if(thumb)thumb.classList.add('failed');}lazyPaint();lazyPump();}")
        builder.AppendLine("function lazyPump(){while(lazyActive<lazyMaxActive&&lazyQueue.length){const img=lazyQueue.shift();lazyActive++;img.dataset.loading='1';img.onload=()=>lazyFinish(img,true);img.onerror=()=>lazyFinish(img,false);img.src=img.dataset.lazySrc;}}")
        builder.AppendLine("lazyImgs.slice(0,6).forEach(lazyEnqueue);if('IntersectionObserver'in window){const io=new IntersectionObserver(entries=>{entries.forEach(entry=>{if(entry.isIntersecting){io.unobserve(entry.target);lazyEnqueue(entry.target);}});},{rootMargin:'420px 0px'});lazyImgs.slice(6).forEach(img=>io.observe(img));}else{lazyImgs.slice(6).forEach(lazyEnqueue);}lazyPaint();")
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
