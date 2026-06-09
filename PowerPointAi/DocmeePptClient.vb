Imports System.Collections.Generic
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class DocmeeTemplateInfo
    Public Property Id As String
    Public Property Name As String
    Public Property Category As String
    Public Property Style As String
    Public Property CoverUrl As String

    Public Overrides Function ToString() As String
        Dim displayName = If(String.IsNullOrWhiteSpace(Name), Id, Name)
        If String.IsNullOrWhiteSpace(Category) AndAlso String.IsNullOrWhiteSpace(Style) Then Return displayName
        Return $"{displayName}（{Category} {Style}）".Trim()
    End Function
End Class

Public Class DocmeePptInfo
    Public Property Id As String
    Public Property TemplateId As String
    Public Property Subject As String
    Public Property CoverUrl As String
End Class

Public Class DocmeeNewPageResult
    Public Property PptxId As String
    Public Property FileUrl As String
End Class

Public Class DocmeePptClient
    Private Shared ReadOnly NewPageTemplateIds As String() = {
        "1804885538940116992",
        "1804889500284084224",
        "1804893649646116864",
        "1804898801006403584",
        "1804901831135191040",
        "1804902366068334592",
        "1804904403862544384",
        "1804905770857521152",
        "1804906177663066112",
        "1805081814809960448",
        "1806268304982269952",
        "1806271593098502144",
        "1806279004165234688",
        "1806280891782389760",
        "1806285286083387392",
        "1806287200204349440",
        "1806290661188820992",
        "1806291734771261440",
        "1806297030256222208",
        "1806297660265848832",
        "1806299845762473984",
        "1806301544875024384",
        "1806303058985213952",
        "1806304212762746880",
        "1806506174552727552",
        "1807601742553276416",
        "1807617227806203904",
        "1807655586138152960",
        "1807658348435464192",
        "1807659875694796800"
    }
    Private Shared ReadOnly NewPageTemplateRandom As New Random()
    Private Shared ReadOnly NewPageTemplateRandomLock As New Object()

    Private Shared ReadOnly Property CreateTaskEndpoint As String
        Get
            Return BuildEndpoint("/api/ppt/v2/createTask")
        End Get
    End Property

    Private Shared ReadOnly Property GenerateContentEndpoint As String
        Get
            Return BuildEndpoint("/api/ppt/v2/generateContent")
        End Get
    End Property

    Private Shared ReadOnly Property GeneratePptxEndpoint As String
        Get
            Return BuildEndpoint("/api/ppt/v2/generatePptx")
        End Get
    End Property

    Private Shared ReadOnly Property TemplateListEndpoint As String
        Get
            Return BuildEndpoint("/api/ppt/templates?lang=zh-CN")
        End Get
    End Property

    Private Shared ReadOnly Property DownloadPptxEndpoint As String
        Get
            Return BuildEndpoint("/api/ppt/downloadPptx")
        End Get
    End Property

    Private Shared ReadOnly Property TextOptimizeEndpoint As String
        Get
            Return BuildEndpoint("/api/ppt/textOptimize")
        End Get
    End Property

    Private Shared ReadOnly Property NewPageWithAiV2Endpoint As String
        Get
            Return BuildEndpoint("/api/ppt/v2/newPageWithAiV2")
        End Get
    End Property

    Public Shared Function GetConfiguredApiBaseUrl() As String
        Return ShareRibbon.ConfigSettings.GetDocmeeApiBaseUrl()
    End Function

    Public Shared Function GetConfiguredToken() As String
        Return ShareRibbon.ConfigSettings.GetDocmeeToken()
    End Function

    Public Shared Function GetRandomNewPageTemplateId() As String
        SyncLock NewPageTemplateRandomLock
            Return NewPageTemplateIds(NewPageTemplateRandom.Next(NewPageTemplateIds.Length))
        End SyncLock
    End Function

    Private Shared Function BuildEndpoint(path As String) As String
        Return GetConfiguredApiBaseUrl() & path
    End Function

    Private Shared Sub AddDocmeeTokenHeader(headers As HttpHeaders)
        Dim token = GetConfiguredToken()
        If String.IsNullOrWhiteSpace(token) Then
            Throw New InvalidOperationException("缺少 Docmee token，请配置 " &
                                                ShareRibbon.ConfigSettings.DocmeeTokenAppSettingKey &
                                                " 或环境变量 " &
                                                ShareRibbon.ConfigSettings.DocmeeTokenEnvironmentVariable & "。")
        End If

        headers.Remove("token")
        headers.TryAddWithoutValidation("token", token)
    End Sub

    Public Shared Function GetFallbackTemplates() As List(Of DocmeeTemplateInfo)
        Return New List(Of DocmeeTemplateInfo) From {
            New DocmeeTemplateInfo With {
                .Id = "1940698099655794688",
                .Name = "橙蓝商务办公通用模板",
                .Category = "办公报告",
                .Style = "扁平简约",
                .CoverUrl = "https://test.chatmee.cn/api/common/oss/meta-doc/ppt_template/1940698099655794688.png"
            },
            New DocmeeTemplateInfo With {
                .Id = "1940697631068151808",
                .Name = "办公简约灰色文职年度工作总结",
                .Category = "办公报告",
                .Style = "扁平简约",
                .CoverUrl = "https://test.chatmee.cn/api/common/oss/meta-doc/ppt_template/1940697631068151808.png"
            },
            New DocmeeTemplateInfo With {
                .Id = "1940698176554164224",
                .Name = "橙色简约风工作总结模板",
                .Category = "教育培训",
                .Style = "扁平简约",
                .CoverUrl = "https://test.chatmee.cn/api/common/oss/meta-doc/ppt_template/1940698176554164224.png"
            },
            New DocmeeTemplateInfo With {
                .Id = "1940698140785139712",
                .Name = "橙色餐饮行业工作报告演示模板",
                .Category = "办公报告",
                .Style = "商务科技",
                .CoverUrl = "https://test.chatmee.cn/api/common/oss/meta-doc/ppt_template/1940698140785139712.png"
            },
            New DocmeeTemplateInfo With {
                .Id = "1940698053686222848",
                .Name = "扁平蓝色员工年终总结模板",
                .Category = "办公报告",
                .Style = "创意趣味",
                .CoverUrl = "https://test.chatmee.cn/api/common/oss/meta-doc/ppt_template/1940698053686222848.png"
            }
        }
    End Function

    Public Async Function CreateTaskAsync(content As String) As Task(Of String)
        Return Await CreateTaskAsync(content, "1")
    End Function

    Public Async Function CreateMarkdownTaskAsync(markdown As String) As Task(Of String)
        Return Await CreateTaskAsync(markdown, "7")
    End Function

    Public Async Function CreateFileTaskAsync(filePath As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(filePath) Then
            Throw New ArgumentException("请选择要生成 PPT 的文档。", NameOf(filePath))
        End If
        If Not File.Exists(filePath) Then
            Throw New FileNotFoundException("未找到要生成 PPT 的文档。", filePath)
        End If
        If Not IsSupportedUploadFile(filePath) Then
            Throw New NotSupportedException("Docmee 文档生成暂不支持该文件格式。支持格式：doc/docx/pdf/ppt/pptx/txt/md/xls/xlsx/csv/html/epub/mobi/xmind/mm。")
        End If

        Return Await CreateFileTaskAsync(filePath, GetTaskTypeForFile(filePath))
    End Function

    Private Async Function CreateTaskAsync(content As String, taskType As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(content) Then
            Throw New ArgumentException("请输入主题或生成要求。", NameOf(content))
        End If
        If String.IsNullOrWhiteSpace(taskType) Then
            Throw New ArgumentException("缺少 Docmee 任务类型。", NameOf(taskType))
        End If

        Using client = CreateHttpClient()
            AddDocmeeTokenHeader(client.DefaultRequestHeaders)
            Using form As New MultipartFormDataContent()
                form.Add(New StringContent(taskType, Encoding.UTF8), "type")
                form.Add(New StringContent(content.Trim(), Encoding.UTF8), "content")

                Using response = Await client.PostAsync(CreateTaskEndpoint, form).ConfigureAwait(False)
                    Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                    EnsureSuccess(response, responseText)

                    Dim payload = JObject.Parse(responseText)
                    EnsureDocmeeSuccess(payload)

                    Dim taskId = TryGetString(payload.SelectToken("data.id"))
                    If String.IsNullOrWhiteSpace(taskId) Then
                        Throw New InvalidOperationException("Docmee 创建任务成功，但未返回任务 ID。")
                    End If

                    Return taskId
                End Using
            End Using
        End Using
    End Function

    Private Async Function CreateFileTaskAsync(filePath As String, taskType As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(taskType) Then
            Throw New ArgumentException("缺少 Docmee 文件任务类型。", NameOf(taskType))
        End If

        Using client = CreateHttpClient()
            AddDocmeeTokenHeader(client.DefaultRequestHeaders)
            Using form As New MultipartFormDataContent()
                form.Add(New StringContent(taskType, Encoding.UTF8), "type")

                Using fileStream As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                    Using fileContent As New StreamContent(fileStream)
                        fileContent.Headers.ContentType = New MediaTypeHeaderValue("application/octet-stream")
                        form.Add(fileContent, "file", Path.GetFileName(filePath))

                        Using response = Await client.PostAsync(CreateTaskEndpoint, form).ConfigureAwait(False)
                            Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                            EnsureSuccess(response, responseText)

                            Dim payload = JObject.Parse(responseText)
                            EnsureDocmeeSuccess(payload)

                            Dim taskId = TryGetString(payload.SelectToken("data.id"))
                            If String.IsNullOrWhiteSpace(taskId) Then
                                Throw New InvalidOperationException("Docmee 创建文档任务成功，但未返回任务 ID。")
                            End If

                            Return taskId
                        End Using
                    End Using
                End Using
            End Using
        End Using
    End Function

    Private Shared Function GetTaskTypeForFile(filePath As String) As String
        ' The task pane needs generateContent to produce editable Markdown.
        ' Docmee type=4 is Word precise conversion and ignores prompt; type=2 supports uploaded documents.
        Dim extension = Path.GetExtension(filePath)
        If String.IsNullOrWhiteSpace(extension) Then Return "2"

        Select Case extension.Trim().ToLowerInvariant()
            Case ".xmind", ".mm"
                Return "3"
            Case Else
                Return "2"
        End Select
    End Function

    Private Shared Function IsSupportedUploadFile(filePath As String) As Boolean
        Dim extension = Path.GetExtension(filePath)
        If String.IsNullOrWhiteSpace(extension) Then Return False

        Select Case extension.Trim().ToLowerInvariant()
            Case ".doc", ".docx", ".pdf", ".ppt", ".pptx", ".txt", ".md", ".xls", ".xlsx", ".csv", ".html", ".epub", ".mobi", ".xmind", ".mm"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public Async Function GenerateContentAsync(taskId As String, Optional progressHandler As Action(Of String) = Nothing) As Task(Of JObject)
        If String.IsNullOrWhiteSpace(taskId) Then
            Throw New ArgumentException("缺少 Docmee 任务 ID。", NameOf(taskId))
        End If

        Dim payload As New JObject From {
            {"id", taskId.Trim()},
            {"stream", True},
            {"outlineType", "JSON"},
            {"questionMode", False},
            {"isNeedAsk", False},
            {"length", "long"},
            {"scene", "产品介绍"},
            {"audience", "客户"},
            {"lang", "zh"},
            {"prompt", "语气专业，适合演示"},
            {"aiSearch", False},
            {"isGenImg", False}
        }

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, GenerateContentEndpoint)
                AddDocmeeTokenHeader(request.Headers)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                    EnsureSuccess(response, "")
                    Using responseStream = Await response.Content.ReadAsStreamAsync().ConfigureAwait(False)
                        Return Await ReadGeneratedOutlineStreamAsync(responseStream, progressHandler).ConfigureAwait(False)
                    End Using
                End Using
            End Using
        End Using
    End Function

    Public Async Function GenerateMarkdownContentAsync(taskId As String, Optional progressHandler As Action(Of String) = Nothing, Optional promptOverride As String = Nothing) As Task(Of String)
        If String.IsNullOrWhiteSpace(taskId) Then
            Throw New ArgumentException("缺少 Docmee 任务 ID。", NameOf(taskId))
        End If

        Dim promptText = NormalizeDocmeePrompt(promptOverride)

        Dim payload As New JObject From {
            {"id", taskId.Trim()},
            {"stream", True},
            {"outlineType", "MD"},
            {"questionMode", False},
            {"isNeedAsk", False},
            {"length", "long"},
            {"scene", "产品介绍"},
            {"audience", "客户"},
            {"lang", "zh"},
            {"prompt", promptText},
            {"aiSearch", False},
            {"isGenImg", False}
        }

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, GenerateContentEndpoint)
                AddDocmeeTokenHeader(request.Headers)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                    EnsureSuccess(response, "")
                    Using responseStream = Await response.Content.ReadAsStreamAsync().ConfigureAwait(False)
                        Return Await ReadGeneratedMarkdownStreamAsync(responseStream, progressHandler).ConfigureAwait(False)
                    End Using
                End Using
            End Using
        End Using
    End Function

    Private Shared Function NormalizeDocmeePrompt(promptOverride As String) As String
        Dim normalized = If(String.IsNullOrWhiteSpace(promptOverride), "语气专业，适合演示", promptOverride.Trim())
        If normalized.Length > 49 Then normalized = normalized.Substring(0, 49)
        Return normalized
    End Function

    Public Async Function ListTemplatesAsync(Optional page As Integer = 1, Optional size As Integer = 10, Optional cancellationToken As Threading.CancellationToken = Nothing) As Task(Of List(Of DocmeeTemplateInfo))
        Dim payload As New JObject From {
            {"page", Math.Max(page, 1)},
            {"size", Math.Max(size, 1)},
            {"filters", New JObject From {
                {"type", 1},
                {"category", JValue.CreateNull()},
                {"style", JValue.CreateNull()},
                {"themeColor", JValue.CreateNull()}
            }}
        }

        Using client = CreateHttpClient(TimeSpan.FromSeconds(12))
            Using request As New HttpRequestMessage(HttpMethod.Post, TemplateListEndpoint)
                AddDocmeeTokenHeader(request.Headers)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request, cancellationToken).ConfigureAwait(False)
                    Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                    EnsureSuccess(response, responseText)

                    Dim result = JObject.Parse(responseText)
                    EnsureDocmeeSuccess(result)

                    Dim templates = ExtractTemplateArray(result)
                    Dim items As New List(Of DocmeeTemplateInfo)()
                    For Each item In templates
                        Dim templateObj = TryCast(item, JObject)
                        If templateObj Is Nothing Then Continue For

                        Dim templateId = TryGetString(templateObj("id"))
                        If String.IsNullOrWhiteSpace(templateId) Then Continue For

                        items.Add(New DocmeeTemplateInfo With {
                            .Id = templateId,
                            .Name = FirstNonEmpty(templateObj, "name", "subject"),
                            .Category = TryGetString(templateObj("category")),
                            .Style = TryGetString(templateObj("style")),
                            .CoverUrl = TryGetString(templateObj("coverUrl"))
                        })
                    Next

                    Return items
                End Using
            End Using
        End Using
    End Function

    Public Async Function DownloadTemplateCoverAsync(coverUrl As String, Optional cancellationToken As Threading.CancellationToken = Nothing) As Task(Of Byte())
        If String.IsNullOrWhiteSpace(coverUrl) Then
            Throw New ArgumentException("缺少模板封面地址。", NameOf(coverUrl))
        End If

        Using handler As New HttpClientHandler() With {.AllowAutoRedirect = False}
            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromSeconds(8)

                Dim currentUrl = coverUrl.Trim()
                For redirectCount As Integer = 0 To 4
                    Using request As New HttpRequestMessage(HttpMethod.Get, currentUrl)
                        AddDocmeeTokenHeader(request.Headers)

                        Using response = Await client.SendAsync(request, cancellationToken).ConfigureAwait(False)
                            If IsRedirectStatusCode(response.StatusCode) Then
                                Dim location = response.Headers.Location
                                If location Is Nothing Then
                                    Throw New HttpRequestException($"Docmee 模板封面重定向缺少 Location: {(CInt(response.StatusCode))}")
                                End If

                                currentUrl = ResolveRedirectUrl(response.RequestMessage.RequestUri, location)
                                Continue For
                            End If

                            If Not response.IsSuccessStatusCode Then
                                Throw New HttpRequestException($"Docmee 模板封面请求失败: {(CInt(response.StatusCode))} {response.ReasonPhrase}")
                            End If

                            Dim bytes = Await response.Content.ReadAsByteArrayAsync().ConfigureAwait(False)
                            Dim contentType = If(response.Content.Headers.ContentType Is Nothing, "", response.Content.Headers.ContentType.MediaType)
                            If bytes Is Nothing OrElse bytes.Length = 0 Then
                                Throw New InvalidOperationException("Docmee 模板封面返回为空。")
                            End If

                            If Not String.IsNullOrWhiteSpace(contentType) AndAlso
                               Not contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase) Then
                                Throw New InvalidOperationException($"Docmee 模板封面返回的不是图片: {contentType}")
                            End If

                            Return bytes
                        End Using
                    End Using
                Next
            End Using
        End Using

        Throw New HttpRequestException("Docmee 模板封面重定向次数过多。")
    End Function

    Private Shared Function IsRedirectStatusCode(statusCode As HttpStatusCode) As Boolean
        Dim code = CInt(statusCode)
        Return code = 301 OrElse code = 302 OrElse code = 303 OrElse code = 307 OrElse code = 308
    End Function

    Private Shared Function ResolveRedirectUrl(baseUri As Uri, location As Uri) As String
        If location.IsAbsoluteUri Then Return location.AbsoluteUri
        Return New Uri(baseUri, location).AbsoluteUri
    End Function

    Public Async Function GeneratePptxAsync(taskId As String, templateId As String, markdown As String) As Task(Of DocmeePptInfo)
        If String.IsNullOrWhiteSpace(taskId) Then
            Throw New ArgumentException("缺少 Docmee 任务 ID。", NameOf(taskId))
        End If
        If String.IsNullOrWhiteSpace(templateId) Then
            Throw New ArgumentException("请选择模板。", NameOf(templateId))
        End If
        If String.IsNullOrWhiteSpace(markdown) Then
            Throw New ArgumentException("缺少 PPT 大纲 Markdown。", NameOf(markdown))
        End If

        Dim payload As New JObject From {
            {"id", taskId.Trim()},
            {"templateId", templateId.Trim()},
            {"markdown", markdown.Trim()}
        }

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, GeneratePptxEndpoint)
                AddDocmeeTokenHeader(request.Headers)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request).ConfigureAwait(False)
                    Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                    EnsureSuccess(response, responseText)

                    Dim result = JObject.Parse(responseText)
                    EnsureDocmeeSuccess(result)

                    Dim pptInfo = TryCast(result.SelectToken("data.pptInfo"), JObject)
                    If pptInfo Is Nothing Then
                        Throw New InvalidOperationException("Docmee 已生成 PPT，但未返回 PPT 信息。")
                    End If

                    Dim generatedInfo As New DocmeePptInfo With {
                        .Id = TryGetString(pptInfo("id")),
                        .TemplateId = TryGetString(pptInfo("templateId")),
                        .Subject = TryGetString(pptInfo("subject")),
                        .CoverUrl = TryGetString(pptInfo("coverUrl"))
                    }

                    If String.IsNullOrWhiteSpace(generatedInfo.Id) Then
                        Throw New InvalidOperationException("Docmee 已生成 PPT，但未返回 PPT ID。")
                    End If

                    Return generatedInfo
                End Using
            End Using
        End Using
    End Function

    Public Async Function DownloadPptxAsync(pptId As String,
                                            Optional refresh As Boolean = False,
                                            Optional maxAttempts As Integer = 1,
                                            Optional retryDelayMs As Integer = 1000) As Task(Of String)
        If String.IsNullOrWhiteSpace(pptId) Then
            Throw New ArgumentException("缺少 PPT ID。", NameOf(pptId))
        End If

        Dim payload As New JObject From {
            {"id", pptId.Trim()},
            {"refresh", refresh}
        }

        Dim attempts = Math.Max(1, maxAttempts)
        Using client = CreateHttpClient()
            For attempt As Integer = 1 To attempts
                Using request As New HttpRequestMessage(HttpMethod.Post, DownloadPptxEndpoint)
                    AddDocmeeTokenHeader(request.Headers)
                    request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                    Using response = Await client.SendAsync(request).ConfigureAwait(False)
                        Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                        EnsureSuccess(response, responseText)

                        Dim result = JObject.Parse(responseText)
                        EnsureDocmeeSuccess(result)

                        Dim fileUrl = TryGetString(result.SelectToken("data.fileUrl"))
                        If Not String.IsNullOrWhiteSpace(fileUrl) Then Return fileUrl

                        If attempt >= attempts Then
                            Throw New InvalidOperationException("Docmee 返回成功，但没有 PPT 下载地址。")
                        End If

                        ShareRibbon.LogInfo("[DocmeeDownload] Download URL is not ready. pptId=" & pptId.Trim() & ", attempt=" & attempt.ToString() & "/" & attempts.ToString())
                    End Using
                End Using

                Await Task.Delay(Math.Max(200, retryDelayMs)).ConfigureAwait(False)
            Next
        End Using

        Throw New InvalidOperationException("Docmee 返回成功，但没有 PPT 下载地址。")
    End Function

    Public Async Function NewPageWithAiV2Async(content As String, pptxId As String, Optional progressHandler As Action(Of String) = Nothing, Optional templateIdOverride As String = "") As Task(Of DocmeeNewPageResult)
        If String.IsNullOrWhiteSpace(content) Then
            Throw New ArgumentException("缺少新单页内容。", NameOf(content))
        End If
        If String.IsNullOrWhiteSpace(pptxId) Then
            Throw New ArgumentException("缺少 Docmee PPT ID。", NameOf(pptxId))
        End If

        Dim templateId = If(String.IsNullOrWhiteSpace(templateIdOverride), GetRandomNewPageTemplateId(), templateIdOverride.Trim())
        Dim payload As New JObject From {
            {"content", content.Trim()},
            {"pptxId", pptxId.Trim()},
            {"templateId", templateId},
            {"stream", True},
            {"lang", "zh"}
        }
        ShareRibbon.LogInfo("[DocmeeNewPage] Request. pptxId=" & pptxId.Trim() & ", templateId=" & templateId & ", contentLength=" & content.Trim().Length.ToString())

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, NewPageWithAiV2Endpoint)
                AddDocmeeTokenHeader(request.Headers)
                request.Headers.TryAddWithoutValidation("lang", "zh")
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                    EnsureSuccess(response, "")
                    Using responseStream = Await response.Content.ReadAsStreamAsync().ConfigureAwait(False)
                        Dim result = Await ReadNewPageWithAiStreamAsync(responseStream, progressHandler).ConfigureAwait(False)
                        If String.IsNullOrWhiteSpace(result.PptxId) Then result.PptxId = pptxId.Trim()
                        ShareRibbon.LogInfo("[DocmeeNewPage] Response parsed. pptxId=" & If(result.PptxId, "") & ", hasFileUrl=" & Not String.IsNullOrWhiteSpace(result.FileUrl))
                        Return result
                    End Using
                End Using
            End Using
        End Using
    End Function

    Public Async Function OptimizeTextAsync(text As String, optimizeType As String, Optional documentId As String = "2029735886777888768", Optional targetLanguageName As String = "", Optional targetLanguageCode As String = "", Optional progressHandler As Action(Of String) = Nothing) As Task(Of String)
        If String.IsNullOrWhiteSpace(text) Then
            Throw New ArgumentException("缺少要处理的文本。", NameOf(text))
        End If
        If String.IsNullOrWhiteSpace(optimizeType) Then
            Throw New ArgumentException("缺少 AI 创作类型。", NameOf(optimizeType))
        End If

        Dim requestLang = "zh"
        If String.Equals(optimizeType.Trim(), "fy", StringComparison.OrdinalIgnoreCase) AndAlso
           Not String.IsNullOrWhiteSpace(targetLanguageCode) Then
            requestLang = targetLanguageCode.Trim()
        End If

        Dim payload As New JObject From {
            {"id", If(String.IsNullOrWhiteSpace(documentId), "2029735886777888768", documentId.Trim())},
            {"type", optimizeType.Trim()},
            {"text", text},
            {"lang", requestLang},
            {"stream", True}
        }

        If String.Equals(optimizeType.Trim(), "fy", StringComparison.OrdinalIgnoreCase) Then
            If Not String.IsNullOrWhiteSpace(targetLanguageCode) Then
                payload("targetLang") = targetLanguageCode.Trim()
                payload("language") = targetLanguageCode.Trim()
            End If
            If Not String.IsNullOrWhiteSpace(targetLanguageName) Then
                payload("targetLanguage") = targetLanguageName.Trim()
            End If
        End If

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, TextOptimizeEndpoint)
                AddDocmeeTokenHeader(request.Headers)
                request.Headers.TryAddWithoutValidation("lang", requestLang)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                    EnsureSuccess(response, "")
                    Using responseStream = Await response.Content.ReadAsStreamAsync().ConfigureAwait(False)
                        Return Await ReadTextOptimizeStreamAsync(responseStream, progressHandler).ConfigureAwait(False)
                    End Using
                End Using
            End Using
        End Using
    End Function

    Public Async Function DownloadPptxFileAsync(fileUrl As String, destinationPath As String) As Task
        If String.IsNullOrWhiteSpace(fileUrl) Then
            Throw New ArgumentException("缺少 PPT 下载地址。", NameOf(fileUrl))
        End If
        If String.IsNullOrWhiteSpace(destinationPath) Then
            Throw New ArgumentException("缺少 PPT 保存路径。", NameOf(destinationPath))
        End If

        Using client = CreateHttpClient()
            Dim bytes = Await client.GetByteArrayAsync(fileUrl).ConfigureAwait(False)
            ValidatePptxBytes(bytes)
            File.WriteAllBytes(destinationPath, bytes)
        End Using
    End Function

    Private Shared Sub ValidatePptxBytes(bytes As Byte())
        If bytes Is Nothing OrElse bytes.Length < 4 Then
            Throw New InvalidOperationException("Docmee PPTX 下载结果为空。")
        End If

        If bytes(0) <> &H50 OrElse bytes(1) <> &H4B Then
            Throw New InvalidOperationException("Docmee PPTX 下载结果不是有效的 PPTX 文件。")
        End If
    End Sub

    Private Shared Function CreateHttpClient() As HttpClient
        Return CreateHttpClient(TimeSpan.FromMinutes(5))
    End Function

    Private Shared Function CreateHttpClient(timeout As TimeSpan) As HttpClient
        Dim client As New HttpClient()
        client.Timeout = timeout
        Return client
    End Function

    Private Shared Sub EnsureSuccess(response As HttpResponseMessage, responseText As String)
        If response.IsSuccessStatusCode Then Return

        Dim detail = If(String.IsNullOrWhiteSpace(responseText), response.ReasonPhrase, responseText)
        Throw New HttpRequestException($"Docmee 接口请求失败: {(CInt(response.StatusCode))} {detail}")
    End Sub

    Private Shared Sub EnsureDocmeeSuccess(payload As JObject)
        Dim codeToken = payload("code")
        If codeToken Is Nothing OrElse codeToken.Type = JTokenType.Null Then Return

        Dim codeValue As Integer
        If Integer.TryParse(codeToken.ToString(), codeValue) AndAlso codeValue = 0 Then Return

        Dim message = TryGetString(payload("message"))
        If String.IsNullOrWhiteSpace(message) Then message = payload.ToString(Formatting.None)
        Throw New InvalidOperationException($"Docmee 返回失败: {message}")
    End Sub

    Private Shared Async Function ReadGeneratedOutlineStreamAsync(responseStream As Stream, progressHandler As Action(Of String)) As Task(Of JObject)
        Dim rawResponse As New StringBuilder()
        Dim finalOutline As JObject = Nothing

        Using reader As New StreamReader(responseStream, Encoding.UTF8)
            Do
                Dim line = Await reader.ReadLineAsync().ConfigureAwait(False)
                If line Is Nothing Then Exit Do

                rawResponse.AppendLine(line)

                Dim trimmed = line.Trim()
                If Not trimmed.StartsWith("data:") Then Continue Do

                Dim dataText = trimmed.Substring(5).Trim()
                If String.IsNullOrWhiteSpace(dataText) OrElse dataText = "[DONE]" Then Continue Do

                Dim eventPayload = TryParseObject(dataText)
                If eventPayload Is Nothing Then Continue Do

                Dim chunkText = TryGetString(eventPayload("text"))
                If Not String.IsNullOrEmpty(chunkText) AndAlso progressHandler IsNot Nothing Then
                    progressHandler.Invoke(chunkText)
                End If

                Dim statusValue As Integer = 0
                Dim statusToken = eventPayload("status")
                If statusToken IsNot Nothing Then Integer.TryParse(statusToken.ToString(), statusValue)

                If statusValue = 4 Then
                    finalOutline = ExtractOutlineFromEnvelope(eventPayload)
                End If
            Loop
        End Using

        If finalOutline IsNot Nothing Then Return finalOutline
        Return ExtractGeneratedOutline(rawResponse.ToString())
    End Function

    Private Shared Async Function ReadGeneratedMarkdownStreamAsync(responseStream As Stream, progressHandler As Action(Of String)) As Task(Of String)
        Dim rawResponse As New StringBuilder()
        Dim markdownBuilder As New StringBuilder()
        Dim finalMarkdown As String = ""

        Using reader As New StreamReader(responseStream, Encoding.UTF8)
            Do
                Dim line = Await reader.ReadLineAsync().ConfigureAwait(False)
                If line Is Nothing Then Exit Do

                rawResponse.AppendLine(line)

                Dim trimmed = line.Trim()
                If Not trimmed.StartsWith("data:") Then Continue Do

                Dim dataText = trimmed.Substring(5).Trim()
                If String.IsNullOrWhiteSpace(dataText) OrElse dataText = "[DONE]" Then Continue Do

                Dim eventPayload = TryParseObject(dataText)
                If eventPayload Is Nothing Then Continue Do

                Dim chunkText = TryGetString(eventPayload("text"))
                If Not String.IsNullOrEmpty(chunkText) Then
                    markdownBuilder.Append(chunkText)
                    If progressHandler IsNot Nothing Then progressHandler.Invoke(chunkText)
                End If

                Dim statusValue As Integer = 0
                Dim statusToken = eventPayload("status")
                If statusToken IsNot Nothing Then Integer.TryParse(statusToken.ToString(), statusValue)

                If statusValue = 4 Then
                    Dim eventMarkdown As String = ""
                    If TryExtractMarkdownFromEnvelope(eventPayload, eventMarkdown) Then
                        finalMarkdown = eventMarkdown
                    End If
                End If
            Loop
        End Using

        If Not String.IsNullOrWhiteSpace(finalMarkdown) Then Return finalMarkdown.Trim()
        If markdownBuilder.Length > 0 Then Return markdownBuilder.ToString().Trim()
        Return ExtractGeneratedMarkdown(rawResponse.ToString())
    End Function

    Private Shared Async Function ReadNewPageWithAiStreamAsync(responseStream As Stream, progressHandler As Action(Of String)) As Task(Of DocmeeNewPageResult)
        Dim rawResponse As New StringBuilder()
        Dim result As New DocmeeNewPageResult()
        Dim streamReadException As Exception = Nothing

        Try
            Using reader As New StreamReader(responseStream, Encoding.UTF8)
                Do
                    Dim line = Await reader.ReadLineAsync().ConfigureAwait(False)
                    If line Is Nothing Then Exit Do

                    rawResponse.AppendLine(line)

                    Dim trimmed = line.Trim()
                    If String.IsNullOrWhiteSpace(trimmed) Then Continue Do

                    Dim payload As JObject = Nothing
                    If trimmed.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                        Dim dataText = trimmed.Substring(5).Trim()
                        If String.IsNullOrWhiteSpace(dataText) OrElse dataText = "[DONE]" Then Continue Do
                        If progressHandler IsNot Nothing Then progressHandler.Invoke(dataText)
                        payload = TryParseObject(dataText)
                    Else
                        If progressHandler IsNot Nothing Then progressHandler.Invoke(trimmed)
                        payload = TryParseObject(trimmed)
                    End If

                    If payload Is Nothing Then Continue Do
                    EnsureDocmeeSuccess(payload)

                    MergeNewPageResult(result, payload)

                    Dim progressText = FirstNonEmpty(payload, "text", "content", "message", "msg")
                    If Not String.IsNullOrWhiteSpace(progressText) AndAlso progressHandler IsNot Nothing Then
                        progressHandler.Invoke("message: " & progressText)
                    End If
                Loop
            End Using
        Catch ex As IOException
            streamReadException = ex
            ShareRibbon.LogInfo("[DocmeeNewPage] Stream closed while reading. parsedPptxId=" & If(result.PptxId, "") & ", rawLength=" & rawResponse.Length.ToString() & ", message=" & ex.Message)
        End Try

        If Not String.IsNullOrWhiteSpace(result.PptxId) OrElse Not String.IsNullOrWhiteSpace(result.FileUrl) Then Return result

        Dim directJson = TryParseObject(rawResponse.ToString())
        If directJson IsNot Nothing Then
            EnsureDocmeeSuccess(directJson)
            MergeNewPageResult(result, directJson)
        End If

        If streamReadException IsNot Nothing AndAlso
           String.IsNullOrWhiteSpace(result.PptxId) AndAlso
           String.IsNullOrWhiteSpace(result.FileUrl) Then
            ShareRibbon.LogInfo("[DocmeeNewPage] Stream closed without download info; caller will use request pptxId as download fallback.")
        End If

        Return result
    End Function

    Private Shared Sub MergeNewPageResult(result As DocmeeNewPageResult, payload As JObject)
        If result Is Nothing OrElse payload Is Nothing Then Return

        Dim pptxId = FirstNonEmpty(payload, "pptxId", "pptId", "id")
        If String.IsNullOrWhiteSpace(pptxId) Then pptxId = TryGetString(payload.SelectToken("data.pptxId"))
        If String.IsNullOrWhiteSpace(pptxId) Then pptxId = TryGetString(payload.SelectToken("data.pptId"))
        If String.IsNullOrWhiteSpace(pptxId) Then pptxId = TryGetString(payload.SelectToken("data.id"))
        If String.IsNullOrWhiteSpace(pptxId) Then pptxId = TryGetString(payload.SelectToken("data.pptInfo.id"))
        If Not String.IsNullOrWhiteSpace(pptxId) Then result.PptxId = pptxId

        Dim fileUrl = FirstNonEmpty(payload, "fileUrl", "url", "downloadUrl")
        If String.IsNullOrWhiteSpace(fileUrl) Then fileUrl = TryGetString(payload.SelectToken("data.fileUrl"))
        If String.IsNullOrWhiteSpace(fileUrl) Then fileUrl = TryGetString(payload.SelectToken("data.url"))
        If String.IsNullOrWhiteSpace(fileUrl) Then fileUrl = TryGetString(payload.SelectToken("data.downloadUrl"))
        If Not String.IsNullOrWhiteSpace(fileUrl) Then result.FileUrl = fileUrl
    End Sub

    Private Shared Async Function ReadTextOptimizeStreamAsync(responseStream As Stream, progressHandler As Action(Of String)) As Task(Of String)
        Dim rawResponse As New StringBuilder()
        Dim resultBuilder As New StringBuilder()
        Dim finalText As String = ""

        Using reader As New StreamReader(responseStream, Encoding.UTF8)
            Do
                Dim line = Await reader.ReadLineAsync().ConfigureAwait(False)
                If line Is Nothing Then Exit Do

                rawResponse.AppendLine(line)

                Dim trimmed = line.Trim()
                If String.IsNullOrWhiteSpace(trimmed) Then Continue Do

                If trimmed.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                    Dim dataText = trimmed.Substring(5).Trim()
                    If String.IsNullOrWhiteSpace(dataText) OrElse dataText = "[DONE]" Then Continue Do

                    Dim eventPayload = TryParseObject(dataText)
                    If eventPayload Is Nothing Then Continue Do

                    Dim chunkText = FirstNonEmpty(eventPayload, "text", "content", "result", "data")
                    If Not String.IsNullOrEmpty(chunkText) Then
                        resultBuilder.Append(chunkText)
                        If progressHandler IsNot Nothing Then progressHandler.Invoke(resultBuilder.ToString())
                    End If

                    Dim statusValue As Integer = 0
                    Dim statusToken = eventPayload("status")
                    If statusToken IsNot Nothing Then Integer.TryParse(statusToken.ToString(), statusValue)
                    If statusValue = 4 Then
                        Dim statusText = FirstNonEmpty(eventPayload, "result", "text", "content")
                        If Not String.IsNullOrWhiteSpace(statusText) Then finalText = statusText
                    End If
                Else
                    Dim directJson = TryParseObject(trimmed)
                    If directJson IsNot Nothing Then
                        EnsureDocmeeSuccess(directJson)
                        Dim directText = FirstNonEmpty(directJson, "text", "content", "result")
                        If String.IsNullOrWhiteSpace(directText) Then directText = TryGetString(directJson.SelectToken("data.text"))
                        If String.IsNullOrWhiteSpace(directText) Then directText = TryGetString(directJson.SelectToken("data.result"))
                        If Not String.IsNullOrWhiteSpace(directText) Then finalText = directText
                    End If
                End If
            Loop
        End Using

        If Not String.IsNullOrWhiteSpace(finalText) Then Return finalText.Trim()
        If resultBuilder.Length > 0 Then Return resultBuilder.ToString().Trim()
        Return ExtractTextOptimizeResult(rawResponse.ToString())
    End Function

    Private Shared Function ExtractTextOptimizeResult(responseText As String) As String
        Dim directJson = TryParseObject(responseText)
        If directJson IsNot Nothing Then
            EnsureDocmeeSuccess(directJson)
            Dim directText = FirstNonEmpty(directJson, "text", "content", "result")
            If String.IsNullOrWhiteSpace(directText) Then directText = TryGetString(directJson.SelectToken("data.text"))
            If String.IsNullOrWhiteSpace(directText) Then directText = TryGetString(directJson.SelectToken("data.result"))
            If Not String.IsNullOrWhiteSpace(directText) Then Return directText.Trim()
        End If

        Dim builder As New StringBuilder()
        Dim lines = responseText.Replace(vbCrLf, vbLf).Split({vbLf}, StringSplitOptions.None)
        For Each rawLine In lines
            Dim line = rawLine.Trim()
            If Not line.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim dataText = line.Substring(5).Trim()
            If String.IsNullOrWhiteSpace(dataText) OrElse dataText = "[DONE]" Then Continue For

            Dim eventPayload = TryParseObject(dataText)
            If eventPayload Is Nothing Then Continue For

            Dim chunkText = FirstNonEmpty(eventPayload, "text", "content", "result", "data")
            If Not String.IsNullOrEmpty(chunkText) Then builder.Append(chunkText)
        Next

        If builder.Length > 0 Then Return builder.ToString().Trim()
        Throw New InvalidOperationException("Docmee 文本创作接口未返回可用内容。")
    End Function

    Private Shared Function ExtractGeneratedOutline(responseText As String) As JObject
        Dim directJson = TryParseObject(responseText)
        If directJson IsNot Nothing Then
            EnsureDocmeeSuccess(directJson)
            Return ExtractOutlineFromEnvelope(directJson)
        End If

        Dim finalOutline As JObject = Nothing
        Dim lines = responseText.Replace(vbCrLf, vbLf).Split({vbLf}, StringSplitOptions.None)
        For Each rawLine In lines
            Dim line = rawLine.Trim()
            If Not line.StartsWith("data:") Then Continue For

            Dim dataText = line.Substring(5).Trim()
            If String.IsNullOrWhiteSpace(dataText) OrElse dataText = "[DONE]" Then Continue For

            Dim eventPayload = TryParseObject(dataText)
            If eventPayload Is Nothing Then Continue For

            Dim statusValue As Integer = 0
            Dim statusToken = eventPayload("status")
            If statusToken IsNot Nothing Then Integer.TryParse(statusToken.ToString(), statusValue)

            If statusValue = 4 Then
                finalOutline = ExtractOutlineFromEnvelope(eventPayload)
            End If
        Next

        If finalOutline IsNot Nothing Then Return finalOutline
        Throw New InvalidOperationException("Docmee 未返回完整大纲内容。")
    End Function

    Private Shared Function ExtractGeneratedMarkdown(responseText As String) As String
        Dim directJson = TryParseObject(responseText)
        If directJson IsNot Nothing Then
            EnsureDocmeeSuccess(directJson)
            Return ExtractMarkdownFromEnvelope(directJson).Trim()
        End If

        Dim finalMarkdown As String = ""
        Dim markdownBuilder As New StringBuilder()
        Dim lines = responseText.Replace(vbCrLf, vbLf).Split({vbLf}, StringSplitOptions.None)
        For Each rawLine In lines
            Dim line = rawLine.Trim()
            If Not line.StartsWith("data:") Then Continue For

            Dim dataText = line.Substring(5).Trim()
            If String.IsNullOrWhiteSpace(dataText) OrElse dataText = "[DONE]" Then Continue For

            Dim eventPayload = TryParseObject(dataText)
            If eventPayload Is Nothing Then Continue For

            Dim chunkText = TryGetString(eventPayload("text"))
            If Not String.IsNullOrEmpty(chunkText) Then markdownBuilder.Append(chunkText)

            Dim statusValue As Integer = 0
            Dim statusToken = eventPayload("status")
            If statusToken IsNot Nothing Then Integer.TryParse(statusToken.ToString(), statusValue)

            If statusValue = 4 Then
                Dim eventMarkdown As String = ""
                If TryExtractMarkdownFromEnvelope(eventPayload, eventMarkdown) Then
                    finalMarkdown = eventMarkdown
                End If
            End If
        Next

        If Not String.IsNullOrWhiteSpace(finalMarkdown) Then Return finalMarkdown.Trim()
        If markdownBuilder.Length > 0 Then Return markdownBuilder.ToString().Trim()
        Throw New InvalidOperationException("Docmee 未返回完整 Markdown 大纲内容。")
    End Function

    Private Shared Function ExtractOutlineFromEnvelope(payload As JObject) As JObject
        Dim resultToken As JToken = payload("result")
        If resultToken Is Nothing AndAlso TypeOf payload("data") Is JObject Then
            resultToken = DirectCast(payload("data"), JObject)("result")
        End If
        If resultToken Is Nothing AndAlso payload("children") IsNot Nothing Then
            resultToken = payload
        End If

        If TypeOf resultToken Is JObject Then Return DirectCast(resultToken, JObject)
        If resultToken IsNot Nothing AndAlso resultToken.Type = JTokenType.String Then
            Dim parsed = TryParseObject(resultToken.ToString())
            If parsed IsNot Nothing Then Return parsed
        End If

        Throw New InvalidOperationException("Docmee 返回内容中没有可用的大纲 result。")
    End Function

    Private Shared Function ExtractMarkdownFromEnvelope(payload As JObject) As String
        Dim markdown As String = ""
        If TryExtractMarkdownFromEnvelope(payload, markdown) Then Return markdown

        Throw New InvalidOperationException("Docmee 返回内容中没有可用的 Markdown result。")
    End Function

    Private Shared Function TryExtractMarkdownFromEnvelope(payload As JObject, ByRef markdown As String) As Boolean
        markdown = ""
        If payload Is Nothing Then Return False

        Dim resultToken As JToken = payload("result")
        If resultToken Is Nothing AndAlso TypeOf payload("data") Is JObject Then
            resultToken = DirectCast(payload("data"), JObject)("result")
        End If

        If resultToken IsNot Nothing AndAlso resultToken.Type = JTokenType.String Then
            markdown = resultToken.ToString()
            Return Not String.IsNullOrWhiteSpace(markdown)
        End If

        Dim textValue = TryGetString(payload("text"))
        If Not String.IsNullOrWhiteSpace(textValue) Then
            markdown = textValue
            Return True
        End If

        Return False
    End Function

    Private Shared Function ExtractTemplateArray(payload As JObject) As JArray
        Dim direct = TryCast(payload("data"), JArray)
        If direct IsNot Nothing Then Return direct

        Dim dataObj = TryCast(payload("data"), JObject)
        If dataObj IsNot Nothing Then
            Dim listToken = dataObj("list")
            If listToken Is Nothing Then listToken = dataObj("records")
            If listToken Is Nothing Then listToken = dataObj("items")

            Dim nested = TryCast(listToken, JArray)
            If nested IsNot Nothing Then Return nested
        End If

        Return New JArray()
    End Function

    Private Shared Function TryParseObject(text As String) As JObject
        If String.IsNullOrWhiteSpace(text) Then Return Nothing

        Try
            Return JObject.Parse(text)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function TryGetString(token As JToken) As String
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return ""
        Return token.ToString()
    End Function

    Private Shared Function FirstNonEmpty(payload As JObject, ParamArray keys As String()) As String
        For Each key In keys
            Dim value = TryGetString(payload(key))
            If Not String.IsNullOrWhiteSpace(value) Then Return value
        Next
        Return ""
    End Function
End Class
