Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' 添加此引用
Partial Class Ribbon1
    Inherits ShareRibbon.BaseOfficeRibbon

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Overloads Sub InitializeComponent()
        Me.TabAI.Label = "wenduoduoAI"

        ' 设置特定的图标
        Me.ConfigApiButton.Image = ShareRibbon.SharedResources.AiApiConfig
        Me.DataAnalysisButton.Image = ShareRibbon.SharedResources.Magic
        ' 提示词配置
        Me.PromptConfigButton.Image = ShareRibbon.SharedResources.promptconfig
        ' 自动补全（已禁用）
        ' Me.AutocompleteSettingsButton.Image = ShareRibbon.SharedResources.autocomplete
        Me.ChatButton.Image = ShareRibbon.SharedResources.Chat
        Me.ClearCacheButton.Image = ShareRibbon.SharedResources.Clear

        ' 设置 Excel 特定的提示
        Me.DataAnalysisButton.SuperTip = "可选中提出的问题和数据后AI帮你整理到另外一个sheet中"
        Me.PromptConfigButton.SuperTip = "优秀的提示词可以更好的帮AI确定自己的定位，让输出内容更符合你的期望"
        Me.ChatButton.SuperTip = "像使用客户端一样与AI对话，聊天更加便捷"

        ' 设置 RibbonType
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"

        Me.MCPButton.Image = ShareRibbon.SharedResources.Mcp1
        Me.BatchDataGenButton.Visible = False
        Me.SpotlightButton.Visible = False
        Me.WebCaptureButton.Visible = False

        Me.WebCaptureButton.Image = ShareRibbon.SharedResources.Send32

        Me.ContinuationButton.Image = ShareRibbon.SharedResources.Aiwrite
        Me.ContinuationButton.Label = "文本优化"
        Me.ContinuationButton.ScreenTip = "PPT文本优化"
        Me.ContinuationButton.SuperTip = "选中文字后可润色、扩写、精简、填充或补全文案，并替换回当前 PPT"

        Me.TranslateButton.Image = ShareRibbon.SharedResources.Translate
        Me.TranslateButton.Label = "文本翻译"
        Me.TranslateButton.ScreenTip = "PPT文本翻译"
        Me.TranslateButton.SuperTip = "翻译当前页文本并替换原文，支持多语言，翻译后自动适配文本框"
        Me.ProofreadButton.Visible = True
        Me.ProofreadButton.Label = "替换单页"
        Me.ProofreadButton.ScreenTip = "PPT单页替换"
        Me.ProofreadButton.SuperTip = "根据输入要求生成新单页，并替换当前幻灯片"
        Me.ReformatButton.Visible = True
        Me.ReformatButton.Label = "美化单页"
        Me.ReformatButton.ScreenTip = "PPT单页美化"
        Me.ReformatButton.SuperTip = "对当前页应用排版美化，统一字体、颜色、标题和文本框适配"
        Me.DataAnalysisButton.Visible = False
        Me.TemplateFormatButton.Visible = True
        Me.TemplateFormatButton.Image = ShareRibbon.SharedResources.Aiwrite
        Me.TemplateFormatButton.Label = "AI生成PPT"
        Me.TemplateFormatButton.ScreenTip = "AI生成PPT"
        Me.TemplateFormatButton.SuperTip = "支持标题生成、文档生成、Markdown 大纲编辑、选择模板生成和一键更换主题"
    End Sub

End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
