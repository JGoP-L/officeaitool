Imports System.IO


Public Class ConfigManager
    Public Shared Property ConfigData As List(Of ConfigItem)

    ' 默认配置文件在当前用户，我的文档下
    Private Shared configFileName As String = "office_ai_config.json"
    Private Shared configFilePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder, configFileName)

    Public Sub LoadConfig()
        Dim loadedData As List(Of ConfigItem) = Nothing
        If File.Exists(configFilePath) Then
            Dim json As String = File.ReadAllText(configFilePath)
            loadedData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of ConfigItem))(json)
        End If

        ConfigData = New List(Of ConfigItem) From {
            CreateOpenAICompatibleConfig(loadedData)
        }

        ' 初始化配置，将数据初始化到 ConfigSettings，方便全局调用
        For Each item In ConfigData
            If item.selected Then
                ConfigSettings.ApiUrl = item.url
                ConfigSettings.ApiKey = item.key
                ConfigSettings.platform = item.pltform
                For Each item_m In item.model
                    If item_m.selected Then
                        If item_m.modelType = ModelType.Chat Then
                            ConfigSettings.ModelName = item_m.modelName
                            ConfigSettings.mcpable = item_m.mcpable
                            ConfigSettings.fimSupported = item_m.fimSupported
                            ConfigSettings.fimUrl = If(String.IsNullOrEmpty(item_m.fimUrl), item.url, item_m.fimUrl)
                        ElseIf item_m.modelType = ModelType.Embedding Then
                            ConfigSettings.EmbeddingModel = item_m.modelName
                        End If
                    End If
                Next
            End If
        Next
    End Sub

    Private Function CreateOpenAICompatibleConfig(loadedData As List(Of ConfigItem)) As ConfigItem
        Dim source = FindOpenAICompatibleSource(loadedData)
        Dim modelName As String = FindSelectedChatModelName(source)

        Dim config As New ConfigItem() With {
            .pltform = "wenduoduoAI",
            .url = If(source IsNot Nothing, source.url, String.Empty),
            .key = If(source IsNot Nothing, source.key, String.Empty),
            .selected = True,
            .translateSelected = True,
            .validated = source IsNot Nothing AndAlso source.validated,
            .providerType = ProviderType.Cloud,
            .isPreset = False,
            .model = New List(Of ConfigItemModel)()
        }

        config.model.Add(New ConfigItemModel() With {
            .modelName = modelName,
            .displayName = modelName,
            .selected = True,
            .translateSelected = True,
            .modelType = ModelType.Chat,
            .mcpable = False,
            .mcpValidated = False,
            .fimSupported = False,
            .fimUrl = config.url
        })

        Return config
    End Function

    Private Function FindOpenAICompatibleSource(loadedData As List(Of ConfigItem)) As ConfigItem
        If loadedData Is Nothing OrElse loadedData.Count = 0 Then Return Nothing

        Dim configured = loadedData.FirstOrDefault(Function(item) item.pltform = "wenduoduoAI")
        If configured IsNot Nothing Then Return configured

        configured = loadedData.FirstOrDefault(Function(item) item.selected AndAlso Not String.IsNullOrWhiteSpace(item.url))
        If configured IsNot Nothing Then Return configured

        Return loadedData.FirstOrDefault(Function(item) Not String.IsNullOrWhiteSpace(item.url))
    End Function

    Private Function FindSelectedChatModelName(source As ConfigItem) As String
        If source Is Nothing OrElse source.model Is Nothing Then Return String.Empty

        Dim selectedModel = source.model.FirstOrDefault(Function(item) item.modelType = ModelType.Chat AndAlso item.selected)
        If selectedModel IsNot Nothing Then Return selectedModel.modelName

        Dim firstChatModel = source.model.FirstOrDefault(Function(item) item.modelType = ModelType.Chat)
        If firstChatModel IsNot Nothing Then Return firstChatModel.modelName

        Return String.Empty
    End Function

    ''' <summary>
    ''' 合并预置配置和用户配置
    ''' 保留用户的key、selected、validated等状态
    ''' </summary>
    Private Sub MergeConfigurations(loadedData As List(Of ConfigItem))
        ' 先添加云端预置配置
        For Each preset In PresetProviders.GetCloudProviders()
            Dim existing = loadedData.FirstOrDefault(Function(x) x.pltform = preset.pltform OrElse x.url = preset.url)
            If existing IsNot Nothing Then
                ' 保留用户的key和selected状态
                preset.key = existing.key
                preset.selected = existing.selected
                preset.validated = existing.validated
                preset.translateSelected = existing.translateSelected
                ' 合并模型列表
                For Each userModel In existing.model
                    Dim presetModel = preset.model.FirstOrDefault(Function(x) x.modelName = userModel.modelName)
                    If presetModel IsNot Nothing Then
                        presetModel.selected = userModel.selected
                        presetModel.mcpValidated = userModel.mcpValidated
                        presetModel.mcpable = userModel.mcpable
                        presetModel.translateSelected = userModel.translateSelected
                        presetModel.fimSupported = userModel.fimSupported
                        presetModel.fimUrl = userModel.fimUrl
                    Else
                        ' 用户添加的自定义模型，保留
                        preset.model.Add(userModel)
                    End If
                Next
            End If
            ConfigData.Add(preset)
        Next

        ' 添加本地预置配置
        For Each preset In PresetProviders.GetLocalProviders()
            Dim existing = loadedData.FirstOrDefault(Function(x) x.pltform = preset.pltform OrElse x.url = preset.url)
            If existing IsNot Nothing Then
                preset.key = existing.key
                preset.selected = existing.selected
                preset.validated = existing.validated
                preset.translateSelected = existing.translateSelected
                ' 本地模型URL可能被用户修改
                If Not String.IsNullOrEmpty(existing.url) Then
                    preset.url = existing.url
                End If
                ' 合并模型列表
                For Each userModel In existing.model
                    Dim presetModel = preset.model.FirstOrDefault(Function(x) x.modelName = userModel.modelName)
                    If presetModel IsNot Nothing Then
                        presetModel.selected = userModel.selected
                        presetModel.mcpValidated = userModel.mcpValidated
                        presetModel.mcpable = userModel.mcpable
                    Else
                        preset.model.Add(userModel)
                    End If
                Next
            End If
            ConfigData.Add(preset)
        Next

        ' 添加用户自定义的非预置配置
        For Each userConfig In loadedData
            Dim isPresetPlatform = ConfigData.Any(Function(x) x.pltform = userConfig.pltform OrElse x.url = userConfig.url)
            If Not isPresetPlatform Then
                ' 用户自定义的服务商，直接添加
                userConfig.isPreset = False
                ConfigData.Add(userConfig)
            End If
        Next
    End Sub


    ' 保存到文件中，默认存在用户的文档目录下
    Public Shared Sub SaveConfig()
        Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(ConfigData, Newtonsoft.Json.Formatting.Indented)
        ' 如果configFilePath的目录不存在就创建
        Dim dir = Path.GetDirectoryName(configFilePath)
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
        '如果文件不存在就创建
        If Not File.Exists(configFilePath) Then
            File.Create(configFilePath).Dispose()
        End If
        File.WriteAllText(configFilePath, json)

    End Sub


    ' Api配置（每次仅可使用1格）
    Public Class ConfigItem
        Public Property pltform As String
        Public Property url As String
        Public Property model As List(Of ConfigItemModel)
        Public Property key As String
        Public Property selected As Boolean

        ' 是否被选为翻译专用平台（在 UI 中为单选，仅允许一个 true）
        Public Property translateSelected As Boolean = False

        ' 是否通过了API验证
        Public Property validated As Boolean

        ' 服务商类型: Cloud(云端) / Local(本地)
        Public Property providerType As ProviderType = ProviderType.Cloud

        ' 获取APIKey的注册链接
        Public Property registerUrl As String = ""

        ' 是否为预置配置
        Public Property isPreset As Boolean = False

        ' 本地模型默认APIKey提示
        Public Property defaultApiKey As String = ""

        Public Overrides Function ToString() As String
            Return pltform
        End Function
    End Class

    ' 具体模型，例：阿里云百炼的 qwen-coder-plus
    Public Class ConfigItemModel
        Public Property modelName As String
        Public Property selected As Boolean

        ' 是否被选为翻译专用平台（在 UI 中为单选，仅允许一个 true）
        Public Property translateSelected As Boolean = False
        Public Property mcpable As Boolean = False
        Public Property mcpValidated As Boolean = False
        
        ' FIM (Fill-In-the-Middle) 补全能力支持
        Public Property fimSupported As Boolean = False
        
        ' FIM API端点（如果与chat端点不同）
        Public Property fimUrl As String = ""

        ' 是否为推理模型
        Public Property isReasoningModel As Boolean = False

        ' 显示名称(含[推理][MCP]标签)
        Public Property displayName As String = ""
        
        ' 模型类型：Chat(对话) / Embedding(向量)
        Public Property modelType As ModelType = ModelType.Chat
        
        Public Overrides Function ToString() As String
            Return If(String.IsNullOrEmpty(displayName), modelName, displayName)
        End Function
    End Class
    
    ' 模型类型枚举
    Public Enum ModelType
        Chat = 0
        Embedding = 1
    End Enum
End Class
