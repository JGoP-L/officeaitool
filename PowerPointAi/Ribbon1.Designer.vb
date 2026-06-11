Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' 添加此引用
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Text

Partial Class Ribbon1
    Inherits ShareRibbon.BaseOfficeRibbon

    Private WithEvents DocumentGeneratePptButton As RibbonButton
    Private WithEvents TextExpandButton As RibbonButton
    Private WithEvents TextShortenButton As RibbonButton
    Private WithEvents TextTranslateButton As RibbonButton
    Private WithEvents GroupSinglePage As RibbonGroup
    Private WithEvents ThemeColorButton As RibbonButton
    Private WithEvents GroupTheme As RibbonGroup

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
        Me.TabAI.Label = "文多多"

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
        Me.ContinuationButton.Visible = False

        Me.TranslateButton.Image = ShareRibbon.SharedResources.Translate
        Me.TranslateButton.Label = "AI创作"
        Me.TranslateButton.ScreenTip = "PPT AI创作"
        Me.TranslateButton.SuperTip = "选中文字后可润色、扩写、缩写或翻译，并替换回当前 PPT"
        Me.ProofreadButton.Visible = True
        Me.ProofreadButton.Image = ShareRibbon.SharedResources.Papers
        Me.ProofreadButton.Label = "AI生成单页"
        Me.ProofreadButton.ScreenTip = "PPT AI生成单页"
        Me.ProofreadButton.SuperTip = "根据输入要求生成新单页，并应用到当前幻灯片"
        Me.ReformatButton.Visible = True
        Me.ReformatButton.Image = ShareRibbon.SharedResources.Magic
        Me.ReformatButton.Label = "美化单页"
        Me.ReformatButton.ScreenTip = "PPT单页美化"
        Me.ReformatButton.SuperTip = "对当前页应用排版美化，统一字体、颜色、标题和文本框适配"
        Me.DataAnalysisButton.Visible = False
        Me.TemplateFormatButton.Visible = True
        Me.TemplateFormatButton.Image = ShareRibbon.SharedResources.Aiwrite
        Me.TemplateFormatButton.Label = "生成PPT"
        Me.TemplateFormatButton.ScreenTip = "生成PPT"
        Me.TemplateFormatButton.SuperTip = "输入主题，AI 生成大纲、选择模板并导入 PPT"

        ConfigureSplitPowerPointRibbon()
    End Sub

    Private Sub ConfigureSplitPowerPointRibbon()
        Me.DocumentGeneratePptButton = Me.Factory.CreateRibbonButton
        Me.TextExpandButton = Me.Factory.CreateRibbonButton
        Me.TextShortenButton = Me.Factory.CreateRibbonButton
        Me.TextTranslateButton = Me.Factory.CreateRibbonButton
        Me.GroupSinglePage = Me.Factory.CreateRibbonGroup
        Me.ThemeColorButton = Me.Factory.CreateRibbonButton
        Me.GroupTheme = Me.Factory.CreateRibbonGroup

        ApplyRibbonIcon(Me.TemplateFormatButton, CreateRibbonGlyphIcon("P", Color.FromArgb(79, 70, 229), Color.FromArgb(24, 144, 255)))
        SetupLargeButton(Me.DocumentGeneratePptButton,
                         "文档生成PPT",
                         "文档生成PPT",
                         "选择 Word/PDF/Markdown 等文档，生成 PPT 大纲并导入 PPT",
                         CreateRibbonGlyphIcon("文", Color.FromArgb(59, 130, 246), Color.FromArgb(20, 184, 166)))
        SetupLargeButton(Me.TranslateButton,
                         "润色",
                         "润色选中文字",
                         "优化选中文本表达，让文字更自然、专业，并替换回当前 PPT",
                         CreateRibbonGlyphIcon("润", Color.FromArgb(99, 102, 241), Color.FromArgb(168, 85, 247)))
        SetupLargeButton(Me.TextExpandButton,
                         "扩写",
                         "扩写选中文字",
                         "补充细节，将选中文本扩展为更完整表述，并替换回当前 PPT",
                         CreateRibbonGlyphIcon("扩", Color.FromArgb(14, 165, 233), Color.FromArgb(99, 102, 241)))
        SetupLargeButton(Me.TextShortenButton,
                         "缩写",
                         "缩写选中文字",
                         "精简选中文本，保留核心信息和重点，并替换回当前 PPT",
                         CreateRibbonGlyphIcon("缩", Color.FromArgb(16, 185, 129), Color.FromArgb(59, 130, 246)))
        SetupLargeButton(Me.TextTranslateButton,
                         "翻译",
                         "翻译选中文字",
                         "选择目标语言，翻译选中文本并替换回当前 PPT",
                         CreateRibbonGlyphIcon("译", Color.FromArgb(245, 158, 11), Color.FromArgb(239, 68, 68)))

        ApplyRibbonIcon(Me.ReformatButton, CreateRibbonGlyphIcon("美", Color.FromArgb(236, 72, 153), Color.FromArgb(124, 58, 237)))
        Me.ProofreadButton.Label = "生成单页"
        Me.ProofreadButton.ScreenTip = "AI生成单页"
        Me.ProofreadButton.SuperTip = "根据输入要求生成新单页，并应用到当前幻灯片"
        ApplyRibbonIcon(Me.ProofreadButton, CreateRibbonGlyphIcon("页", Color.FromArgb(45, 212, 191), Color.FromArgb(37, 99, 235)))
        SetupLargeButton(Me.ThemeColorButton,
                         "一键换主题",
                         "PPT一键换主题",
                         "选择颜色后，将当前演示文稿的主题色、形状、线条和有色文字统一切换为对应色系",
                         CreateRibbonGlyphIcon("色", Color.FromArgb(124, 58, 237), Color.FromArgb(245, 158, 11)))

        Me.GroupChat.Items.Clear()
        Me.GroupChat.Items.Add(Me.TemplateFormatButton)
        Me.GroupChat.Items.Add(Me.DocumentGeneratePptButton)
        Me.GroupChat.Label = "AI生成PPT"

        Me.GroupAIContent.Items.Clear()
        Me.GroupAIContent.Items.Add(Me.TranslateButton)
        Me.GroupAIContent.Items.Add(Me.TextExpandButton)
        Me.GroupAIContent.Items.Add(Me.TextShortenButton)
        Me.GroupAIContent.Items.Add(Me.TextTranslateButton)
        Me.GroupAIContent.Label = "AI文本创作"

        Me.GroupSinglePage.Items.Add(Me.ReformatButton)
        Me.GroupSinglePage.Items.Add(Me.ProofreadButton)
        Me.GroupSinglePage.Label = "AI单页"
        Me.GroupSinglePage.Name = "GroupSinglePage"

        Me.GroupTheme.Items.Add(Me.ThemeColorButton)
        Me.GroupTheme.Label = "AI主题"
        Me.GroupTheme.Name = "GroupTheme"

        Me.GroupTools.Visible = False
        Me.GroupMCP.Visible = False
        Me.GroupAbout.Visible = False
        Me.TabAI.Groups.Clear()
        Me.TabAI.Groups.Add(Me.GroupChat)
        Me.TabAI.Groups.Add(Me.GroupAIContent)
        Me.TabAI.Groups.Add(Me.GroupSinglePage)
        Me.TabAI.Groups.Add(Me.GroupTheme)

        AddHandler Me.DocumentGeneratePptButton.Click, AddressOf DocumentGeneratePptButton_Click
        AddHandler Me.TextExpandButton.Click, AddressOf TextExpandButton_Click
        AddHandler Me.TextShortenButton.Click, AddressOf TextShortenButton_Click
        AddHandler Me.TextTranslateButton.Click, AddressOf TextTranslateButton_Click
        AddHandler Me.ThemeColorButton.Click, AddressOf ThemeColorButton_Click
    End Sub

    Private Sub SetupLargeButton(button As RibbonButton,
                                 label As String,
                                 screenTip As String,
                                 superTip As String,
                                 image As Drawing.Image)
        button.Label = label
        button.Name = label.Replace(" ", "")
        button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        button.ShowImage = True
        button.Image = image
        button.ScreenTip = screenTip
        button.SuperTip = superTip
    End Sub

    Private Sub ApplyRibbonIcon(button As RibbonButton, image As Image)
        If button Is Nothing Then Return
        button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        button.ShowImage = True
        button.Image = image
    End Sub

    Private Shared Function CreateRibbonGlyphIcon(glyph As String, startColor As Color, endColor As Color) As Image
        Dim bitmap As New Bitmap(32, 32)
        Using g = Graphics.FromImage(bitmap)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit
            g.Clear(Color.Transparent)

            Dim rect As New Rectangle(2, 2, 28, 28)
            Using shadowPath = CreateRoundedIconPath(New Rectangle(3, 4, 28, 28), 7)
                Using shadowBrush As New SolidBrush(Color.FromArgb(32, 15, 23, 42))
                    g.FillPath(shadowBrush, shadowPath)
                End Using
            End Using

            Using path = CreateRoundedIconPath(rect, 7)
                Using brush As New LinearGradientBrush(rect, startColor, endColor, LinearGradientMode.ForwardDiagonal)
                    g.FillPath(brush, path)
                End Using

                Using shineBrush As New LinearGradientBrush(New Rectangle(2, 2, 28, 14),
                                                            Color.FromArgb(70, Color.White),
                                                            Color.FromArgb(0, Color.White),
                                                            LinearGradientMode.Vertical)
                    g.FillPath(shineBrush, path)
                End Using
            End Using

            Dim fontSize As Single = If(glyph.Length > 1, 9.0F, 14.0F)
            Using font As New Font("Microsoft YaHei UI", fontSize, FontStyle.Bold, GraphicsUnit.Pixel)
                Using textBrush As New SolidBrush(Color.White)
                    Using format As New StringFormat()
                        format.Alignment = StringAlignment.Center
                        format.LineAlignment = StringAlignment.Center
                        g.DrawString(glyph, font, textBrush, New RectangleF(2, 2, 28, 28), format)
                    End Using
                End Using
            End Using
        End Using

        Return bitmap
    End Function

    Private Shared Function CreateRoundedIconPath(rect As Rectangle, radius As Integer) As GraphicsPath
        Dim path As New GraphicsPath()
        Dim diameter = Math.Max(1, radius * 2)
        Dim arc As New Rectangle(rect.Location, New Size(diameter, diameter))

        path.AddArc(arc, 180, 90)
        arc.X = rect.Right - diameter
        path.AddArc(arc, 270, 90)
        arc.Y = rect.Bottom - diameter
        path.AddArc(arc, 0, 90)
        arc.X = rect.Left
        path.AddArc(arc, 90, 90)
        path.CloseFigure()
        Return path
    End Function

End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
