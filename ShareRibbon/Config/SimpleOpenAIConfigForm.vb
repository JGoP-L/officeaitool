Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms

Public Class SimpleOpenAIConfigForm
    Inherits Form

    Private Const PlatformName As String = "wenduoduoAI"

    Private ReadOnly apiUrlTextBox As TextBox
    Private ReadOnly apiKeyTextBox As TextBox
    Private ReadOnly modelNameTextBox As TextBox
    Private ReadOnly saveButton As Button
    Private ReadOnly cancelButton As Button

    Public Sub New()
        Me.Text = "配置模型"
        Me.Size = New Size(520, 300)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White

        Dim root As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 5,
            .Padding = New Padding(18)
        }
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 44))

        apiUrlTextBox = CreateTextBox()
        apiKeyTextBox = CreateTextBox()
        apiKeyTextBox.UseSystemPasswordChar = True
        modelNameTextBox = CreateTextBox()

        root.Controls.Add(CreateLabel("API 地址"), 0, 0)
        root.Controls.Add(apiUrlTextBox, 1, 0)
        root.Controls.Add(CreateLabel("API Key"), 0, 1)
        root.Controls.Add(apiKeyTextBox, 1, 1)
        root.Controls.Add(CreateLabel("模型名称"), 0, 2)
        root.Controls.Add(modelNameTextBox, 1, 2)

        Dim hintLabel As New Label() With {
            .Text = "请填写兼容 OpenAI Chat Completions 协议的接口地址，例如 https://api.example.com/v1/chat/completions。",
            .Dock = DockStyle.Fill,
            .ForeColor = Color.FromArgb(90, 90, 90),
            .AutoSize = False
        }
        root.Controls.Add(hintLabel, 1, 3)

        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.RightToLeft
        }
        saveButton = New Button() With {.Text = "保存", .Width = 88, .Height = 30}
        cancelButton = New Button() With {.Text = "取消", .Width = 88, .Height = 30, .DialogResult = DialogResult.Cancel}
        buttonPanel.Controls.Add(saveButton)
        buttonPanel.Controls.Add(cancelButton)
        root.Controls.Add(buttonPanel, 1, 4)

        AddHandler saveButton.Click, AddressOf SaveButton_Click
        Me.AcceptButton = saveButton
        Me.CancelButton = cancelButton
        Me.Controls.Add(root)

        LoadCurrentConfig()
    End Sub

    Private Function CreateLabel(text As String) As Label
        Return New Label() With {
            .Text = text,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleLeft,
            .AutoSize = False
        }
    End Function

    Private Function CreateTextBox() As TextBox
        Return New TextBox() With {
            .Dock = DockStyle.Fill,
            .Margin = New Padding(0, 6, 0, 0)
        }
    End Function

    Private Sub LoadCurrentConfig()
        apiUrlTextBox.Text = ConfigSettings.ApiUrl
        apiKeyTextBox.Text = ConfigSettings.ApiKey
        modelNameTextBox.Text = ConfigSettings.ModelName
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs)
        Dim apiUrl = apiUrlTextBox.Text.Trim()
        Dim apiKey = apiKeyTextBox.Text.Trim()
        Dim modelName = modelNameTextBox.Text.Trim()

        If String.IsNullOrWhiteSpace(apiUrl) Then
            MessageBox.Show("请填写 API 地址。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            apiUrlTextBox.Focus()
            Return
        End If

        Dim parsedUri As Uri = Nothing
        If Not Uri.TryCreate(apiUrl, UriKind.Absolute, parsedUri) OrElse _
           (parsedUri.Scheme <> Uri.UriSchemeHttp AndAlso parsedUri.Scheme <> Uri.UriSchemeHttps) Then
            MessageBox.Show("API 地址必须是 http 或 https 开头的完整地址。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            apiUrlTextBox.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(apiKey) Then
            MessageBox.Show("请填写 API Key。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            apiKeyTextBox.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(modelName) Then
            MessageBox.Show("请填写模型名称。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            modelNameTextBox.Focus()
            Return
        End If

        SaveOpenAICompatibleConfig(apiUrl, apiKey, modelName)
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub SaveOpenAICompatibleConfig(apiUrl As String, apiKey As String, modelName As String)
        If ConfigManager.ConfigData Is Nothing Then
            ConfigManager.ConfigData = New List(Of ConfigManager.ConfigItem)()
        End If

        For Each item In ConfigManager.ConfigData
            item.selected = False
            item.translateSelected = False
            If item.model Is Nothing Then Continue For
            For Each model In item.model
                model.selected = False
                model.translateSelected = False
            Next
        Next

        Dim config = ConfigManager.ConfigData.FirstOrDefault(Function(item) item.pltform = PlatformName)
        If config Is Nothing Then
            config = New ConfigManager.ConfigItem() With {
                .pltform = PlatformName,
                .providerType = ProviderType.Cloud,
                .isPreset = False
            }
            ConfigManager.ConfigData.Add(config)
        End If

        config.pltform = PlatformName
        config.url = apiUrl
        config.key = apiKey
        config.selected = True
        config.translateSelected = True
        config.validated = True
        config.providerType = ProviderType.Cloud
        config.isPreset = False
        config.model = New List(Of ConfigManager.ConfigItemModel) From {
            New ConfigManager.ConfigItemModel() With {
                .modelName = modelName,
                .displayName = modelName,
                .selected = True,
                .translateSelected = True,
                .modelType = ConfigManager.ModelType.Chat,
                .mcpable = False,
                .mcpValidated = False,
                .fimSupported = False,
                .fimUrl = apiUrl
            }
        }

        ConfigSettings.platform = PlatformName
        ConfigSettings.ApiUrl = apiUrl
        ConfigSettings.ApiKey = apiKey
        ConfigSettings.ModelName = modelName
        ConfigSettings.mcpable = False
        ConfigSettings.EmbeddingModel = String.Empty
        ConfigSettings.fimSupported = False
        ConfigSettings.fimUrl = apiUrl

        ConfigManager.SaveConfig()
    End Sub
End Class
