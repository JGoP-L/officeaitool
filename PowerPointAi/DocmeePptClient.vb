Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Http
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

Public Class DocmeePptClient
    Private Const ApiBaseUrl As String = "https://test.docmee.cn"
    Private Const DemoToken As String = "ak_demo"

    Private Shared ReadOnly CreateTaskEndpoint As String = ApiBaseUrl & "/api/ppt/v2/createTask"
    Private Shared ReadOnly GenerateContentEndpoint As String = ApiBaseUrl & "/api/ppt/v2/generateContent"
    Private Shared ReadOnly GeneratePptxEndpoint As String = ApiBaseUrl & "/api/ppt/v2/generatePptx"
    Private Shared ReadOnly TemplateListEndpoint As String = ApiBaseUrl & "/api/ppt/templates?lang=zh-CN"
    Private Shared ReadOnly DownloadPptxEndpoint As String = ApiBaseUrl & "/api/ppt/downloadPptx"

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

    Private Async Function CreateTaskAsync(content As String, taskType As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(content) Then
            Throw New ArgumentException("请输入主题或生成要求。", NameOf(content))
        End If
        If String.IsNullOrWhiteSpace(taskType) Then
            Throw New ArgumentException("缺少 Docmee 任务类型。", NameOf(taskType))
        End If

        Using client = CreateHttpClient()
            client.DefaultRequestHeaders.Add("token", DemoToken)
            Using form As New MultipartFormDataContent()
                form.Add(New StringContent(taskType, Encoding.UTF8), "type")
                form.Add(New StringContent(content.Trim(), Encoding.UTF8), "content")

                Using response = Await client.PostAsync(CreateTaskEndpoint, form)
                    Dim responseText = Await response.Content.ReadAsStringAsync()
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
            {"length", "medium"},
            {"scene", "产品介绍"},
            {"audience", "客户"},
            {"lang", "zh"},
            {"prompt", "语气专业，适合演示"},
            {"aiSearch", False},
            {"isGenImg", False}
        }

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, GenerateContentEndpoint)
                request.Headers.Add("token", DemoToken)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    EnsureSuccess(response, "")
                    Using responseStream = Await response.Content.ReadAsStreamAsync()
                        Return Await ReadGeneratedOutlineStreamAsync(responseStream, progressHandler)
                    End Using
                End Using
            End Using
        End Using
    End Function

    Public Async Function GenerateMarkdownContentAsync(taskId As String, Optional progressHandler As Action(Of String) = Nothing) As Task(Of String)
        If String.IsNullOrWhiteSpace(taskId) Then
            Throw New ArgumentException("缺少 Docmee 任务 ID。", NameOf(taskId))
        End If

        Dim payload As New JObject From {
            {"id", taskId.Trim()},
            {"stream", True},
            {"outlineType", "MD"},
            {"questionMode", False},
            {"isNeedAsk", False},
            {"length", "medium"},
            {"scene", "产品介绍"},
            {"audience", "客户"},
            {"lang", "zh"},
            {"prompt", "语气专业，适合演示"},
            {"aiSearch", False},
            {"isGenImg", False}
        }

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, GenerateContentEndpoint)
                request.Headers.Add("token", DemoToken)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    EnsureSuccess(response, "")
                    Using responseStream = Await response.Content.ReadAsStreamAsync()
                        Return Await ReadGeneratedMarkdownStreamAsync(responseStream, progressHandler)
                    End Using
                End Using
            End Using
        End Using
    End Function

    Public Async Function ListTemplatesAsync(Optional page As Integer = 1, Optional size As Integer = 10) As Task(Of List(Of DocmeeTemplateInfo))
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

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, TemplateListEndpoint)
                request.Headers.Add("token", DemoToken)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request)
                    Dim responseText = Await response.Content.ReadAsStringAsync()
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
                request.Headers.Add("token", DemoToken)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request)
                    Dim responseText = Await response.Content.ReadAsStringAsync()
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

    Public Async Function DownloadPptxAsync(pptId As String, Optional refresh As Boolean = False) As Task(Of String)
        If String.IsNullOrWhiteSpace(pptId) Then
            Throw New ArgumentException("缺少 PPT ID。", NameOf(pptId))
        End If

        Dim payload As New JObject From {
            {"id", pptId.Trim()},
            {"refresh", refresh}
        }

        Using client = CreateHttpClient()
            Using request As New HttpRequestMessage(HttpMethod.Post, DownloadPptxEndpoint)
                request.Headers.Add("token", DemoToken)
                request.Content = New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request)
                    Dim responseText = Await response.Content.ReadAsStringAsync()
                    EnsureSuccess(response, responseText)

                    Dim result = JObject.Parse(responseText)
                    EnsureDocmeeSuccess(result)

                    Dim fileUrl = TryGetString(result.SelectToken("data.fileUrl"))
                    If String.IsNullOrWhiteSpace(fileUrl) Then
                        Throw New InvalidOperationException("Docmee 返回成功，但没有 PPT 下载地址。")
                    End If

                    Return fileUrl
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
            Dim bytes = Await client.GetByteArrayAsync(fileUrl)
            File.WriteAllBytes(destinationPath, bytes)
        End Using
    End Function

    Private Shared Function CreateHttpClient() As HttpClient
        Dim client As New HttpClient()
        client.Timeout = TimeSpan.FromMinutes(5)
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
                Dim line = Await reader.ReadLineAsync()
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
                Dim line = Await reader.ReadLineAsync()
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
