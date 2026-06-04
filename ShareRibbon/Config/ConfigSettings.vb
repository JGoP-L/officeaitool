Imports System.Configuration
Imports System.IO

' 存储配置的api大模型和api key
Public Class ConfigSettings
    Private Sub New()
    End Sub

    Public Shared Property platform As String
    Public Shared Property ApiUrl As String
    Public Shared Property ApiKey As String
    Public Shared Property ModelName As String
    Public Shared Property mcpable As Boolean

    ' Embedding 模型配置
    Public Shared Property EmbeddingModel As String = ""

    ' FIM (Fill-In-the-Middle) 补全能力
    Public Shared Property fimSupported As Boolean = False
    Public Shared Property fimUrl As String = ""

    ' 提示词相关配置
    Public Shared Property propmtName As String
    Public Shared Property propmtContent As String

    ' Agent 架构切换：true=使用新 AgentKernel，false=保留旧 RalphLoop/RalphAgent
    Public Shared Property UseNewAgentKernel As Boolean = True

    Public Const DefaultDocmeeApiBaseUrl As String = "https://test.docmee.cn"
    Public Const DefaultDocmeeToken As String = "ak_demo"
    Public Const DocmeeApiBaseUrlAppSettingKey As String = "OfficeAi.DocmeeApiBaseUrl"
    Public Const DocmeeTokenAppSettingKey As String = "OfficeAi.DocmeeToken"
    Public Const DocmeeApiBaseUrlEnvironmentVariable As String = "OFFICE_AI_DOCMEE_API_BASE_URL"
    Public Const DocmeeTokenEnvironmentVariable As String = "OFFICE_AI_DOCMEE_TOKEN"
    Public Const DocmeeSettingsFileName As String = "docmee_settings.json"

    Public Shared Property DocmeeApiBaseUrl As String = ""
    Public Shared Property DocmeeToken As String = ""

    Public Const OfficeAiAppDataFolder As String = "OfficeAiAppData"

    Public Shared Function GetDocmeeApiBaseUrl() As String
        Dim userSettings = LoadDocmeeUserSettings()
        Dim configured = FirstNonEmpty(
            DocmeeApiBaseUrl,
            If(userSettings Is Nothing, "", userSettings.ApiBaseUrl),
            Environment.GetEnvironmentVariable(DocmeeApiBaseUrlEnvironmentVariable),
            SafeAppSetting(DocmeeApiBaseUrlAppSettingKey),
            DefaultDocmeeApiBaseUrl)

        Return configured.Trim().TrimEnd("/"c)
    End Function

    Public Shared Function GetDocmeeToken() As String
        Dim userSettings = LoadDocmeeUserSettings()
        Return FirstNonEmpty(
            DocmeeToken,
            If(userSettings Is Nothing, "", userSettings.Token),
            Environment.GetEnvironmentVariable(DocmeeTokenEnvironmentVariable),
            SafeAppSetting(DocmeeTokenAppSettingKey),
            DefaultDocmeeToken).Trim()
    End Function

    Public Shared Sub SaveDocmeeSettings(apiBaseUrl As String, token As String)
        Dim normalizedBaseUrl = If(apiBaseUrl, "").Trim().TrimEnd("/"c)
        Dim normalizedToken = If(token, "").Trim()

        Dim settings As New DocmeeSettingsData With {
            .ApiBaseUrl = normalizedBaseUrl,
            .Token = normalizedToken
        }

        Dim settingsPath = GetDocmeeSettingsFilePath()
        Dim settingsDirectory = Path.GetDirectoryName(settingsPath)
        If Not Directory.Exists(settingsDirectory) Then
            Directory.CreateDirectory(settingsDirectory)
        End If

        Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(settings, Newtonsoft.Json.Formatting.Indented)
        File.WriteAllText(settingsPath, json)

        DocmeeApiBaseUrl = normalizedBaseUrl
        DocmeeToken = normalizedToken
    End Sub

    Private Shared Function LoadDocmeeUserSettings() As DocmeeSettingsData
        Try
            Dim settingsPath = GetDocmeeSettingsFilePath()
            If Not File.Exists(settingsPath) Then Return Nothing

            Dim json = File.ReadAllText(settingsPath)
            If String.IsNullOrWhiteSpace(json) Then Return Nothing

            Return Newtonsoft.Json.JsonConvert.DeserializeObject(Of DocmeeSettingsData)(json)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function GetDocmeeSettingsFilePath() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            OfficeAiAppDataFolder,
            DocmeeSettingsFileName)
    End Function

    Private Shared Function FirstNonEmpty(ParamArray values As String()) As String
        For Each value In values
            If Not String.IsNullOrWhiteSpace(value) Then Return value
        Next

        Return ""
    End Function

    Private Shared Function SafeAppSetting(key As String) As String
        Try
            Return ConfigurationManager.AppSettings(key)
        Catch
            Return ""
        End Try
    End Function

    Private Class DocmeeSettingsData
        Public Property ApiBaseUrl As String
        Public Property Token As String
    End Class
End Class
