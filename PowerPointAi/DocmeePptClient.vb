Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class DocmeePptClient
    Private Const ApiBaseUrl As String = "https://test.docmee.cn"
    Private Const DemoToken As String = "ak_demo"

    Private Shared ReadOnly CreateTaskEndpoint As String = ApiBaseUrl & "/api/ppt/v2/createTask"
    Private Shared ReadOnly GenerateContentEndpoint As String = ApiBaseUrl & "/api/ppt/v2/generateContent"
    Private Shared ReadOnly TemplateListEndpoint As String = ApiBaseUrl & "/api/ppt/templates"

    Public Async Function CreateTaskAsync(content As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(content) Then
            Throw New ArgumentException("请输入主题或生成要求。", NameOf(content))
        End If

        Using client = CreateHttpClient()
            client.DefaultRequestHeaders.Add("token", DemoToken)
            Using form As New MultipartFormDataContent()
                form.Add(New StringContent("1", Encoding.UTF8), "type")
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

    Public Async Function GenerateContentAsync(taskId As String) As Task(Of JObject)
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
                    Dim responseText = Await response.Content.ReadAsStringAsync()
                    EnsureSuccess(response, responseText)
                    Return ExtractGeneratedOutline(responseText)
                End Using
            End Using
        End Using
    End Function

    Public Async Function ListTemplatesAsync(Optional page As Integer = 1, Optional size As Integer = 10) As Task(Of JArray)
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

                    Dim data = TryCast(result("data"), JObject)
                    If data Is Nothing Then Return New JArray()

                    Dim listToken = data("list")
                    If listToken Is Nothing Then listToken = data("records")
                    If listToken Is Nothing Then listToken = data("items")

                    Dim templates = TryCast(listToken, JArray)
                    If templates Is Nothing Then Return New JArray()
                    Return templates
                End Using
            End Using
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
End Class
