' ShareRibbon\Config\OfficeAIStyleHelper.vb
' 统一 WinForms 控件美化工具类
' 为所有 Office 插件的 WinForms 界面提供统一的设计语言

Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Module OfficeAIStyleHelper

    ' ========== 品牌色常量 ==========
    Public ReadOnly BrandPrimary As Color = Color.FromArgb(79, 70, 229)       ' #4F46E5 靛蓝主色
    Public ReadOnly BrandPrimaryDark As Color = Color.FromArgb(55, 48, 163)    ' #3730A3 深色
    Public ReadOnly BrandPrimaryLight As Color = Color.FromArgb(238, 242, 255) ' #EEF2FF 极浅蓝背景
    Public ReadOnly BrandAccent As Color = Color.FromArgb(249, 115, 22)        ' #F97316 橙色强调
    Public ReadOnly BrandSuccess As Color = Color.FromArgb(34, 197, 94)        ' #22C55E 绿色
    Public ReadOnly BrandDanger As Color = Color.FromArgb(239, 68, 68)         ' #EF4444 红色

    Public ReadOnly TextPrimary As Color = Color.FromArgb(30, 41, 59)          ' #1E293B 主文字
    Public ReadOnly TextSecondary As Color = Color.FromArgb(100, 116, 139)     ' #64748B 次文字
    Public ReadOnly TextMuted As Color = Color.FromArgb(148, 163, 184)         ' #94A3B8 弱文字

    Public ReadOnly BgPage As Color = Color.FromArgb(248, 250, 252)            ' #F8FAFC 页面背景
    Public ReadOnly BgSurface As Color = Color.White                           ' 表面/卡片
    Public ReadOnly BgInput As Color = Color.FromArgb(241, 245, 249)           ' #F1F5F9 输入框背景

    Public ReadOnly BorderLight As Color = Color.FromArgb(226, 232, 240)       ' #E2E8F0 浅边框
    Public ReadOnly BorderMedium As Color = Color.FromArgb(203, 213, 225)      ' #CBD5E1 中边框

    ' 字体
    Public ReadOnly FontUi As New Font("Microsoft YaHei UI", 9.0!, FontStyle.Regular)
    Public ReadOnly FontUiBold As New Font("Microsoft YaHei UI", 9.0!, FontStyle.Bold)
    Public ReadOnly FontUiSmall As New Font("Microsoft YaHei UI", 8.0!, FontStyle.Regular)
    Public ReadOnly FontTitle As New Font("Microsoft YaHei UI", 13.0!, FontStyle.Bold)
    Public ReadOnly FontHeading As New Font("Microsoft YaHei UI", 11.0!, FontStyle.Bold)

    ' 尺寸常量
    Public Const ButtonHeight As Integer = 32
    Public Const ButtonHeightSmall As Integer = 28
    Public Const InputHeight As Integer = 28
    Public Const CornerRadius As Integer = 6
    Public Const SpacingXs As Integer = 4
    Public Const SpacingSm As Integer = 8
    Public Const SpacingMd As Integer = 12
    Public Const SpacingLg As Integer = 16

    ' ========== 按钮样式 ==========

    ''' <summary>品牌色填充主按钮</summary>
    Public Sub StyleButtonPrimary(btn As Button)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.BackColor = BrandPrimary
        btn.ForeColor = Color.White
        btn.Font = FontUiBold
        btn.Height = ButtonHeight
        btn.Cursor = Cursors.Hand
        btn.TextAlign = ContentAlignment.MiddleCenter
        btn.UseVisualStyleBackColor = False
        AddHandler btn.MouseEnter, Sub(s, e)
                                       btn.BackColor = BrandPrimaryDark
                                   End Sub
        AddHandler btn.MouseLeave, Sub(s, e)
                                       btn.BackColor = BrandPrimary
                                   End Sub
        AddHandler btn.EnabledChanged, Sub(s, e)
                                           If btn.Enabled Then
                                               btn.BackColor = BrandPrimary
                                           Else
                                               btn.BackColor = BorderMedium
                                               btn.ForeColor = TextMuted
                                           End If
                                       End Sub
    End Sub

    ''' <summary>边框次要按钮</summary>
    Public Sub StyleButtonSecondary(btn As Button)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 1
        btn.FlatAppearance.BorderColor = BorderMedium
        btn.BackColor = BgSurface
        btn.ForeColor = TextPrimary
        btn.Font = FontUi
        btn.Height = ButtonHeight
        btn.Cursor = Cursors.Hand
        btn.TextAlign = ContentAlignment.MiddleCenter
        btn.UseVisualStyleBackColor = False
        AddHandler btn.MouseEnter, Sub(s, e)
                                       btn.BackColor = BrandPrimaryLight
                                       btn.FlatAppearance.BorderColor = BrandPrimary
                                   End Sub
        AddHandler btn.MouseLeave, Sub(s, e)
                                       btn.BackColor = BgSurface
                                       btn.FlatAppearance.BorderColor = BorderMedium
                                   End Sub
        AddHandler btn.EnabledChanged, Sub(s, e)
                                           If btn.Enabled Then
                                               btn.BackColor = BgSurface
                                               btn.ForeColor = TextPrimary
                                               btn.FlatAppearance.BorderColor = BorderMedium
                                           Else
                                               btn.BackColor = BgPage
                                               btn.ForeColor = TextMuted
                                               btn.FlatAppearance.BorderColor = BorderLight
                                           End If
                                       End Sub
    End Sub

    ''' <summary>强调色按钮（橙色）</summary>
    Public Sub StyleButtonAccent(btn As Button)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.BackColor = BrandAccent
        btn.ForeColor = Color.White
        btn.Font = FontUiBold
        btn.Height = ButtonHeight
        btn.Cursor = Cursors.Hand
        btn.TextAlign = ContentAlignment.MiddleCenter
        btn.UseVisualStyleBackColor = False
        AddHandler btn.MouseEnter, Sub(s, e)
                                       btn.BackColor = Color.FromArgb(234, 88, 12)
                                   End Sub
        AddHandler btn.MouseLeave, Sub(s, e)
                                       btn.BackColor = BrandAccent
                                   End Sub
        AddHandler btn.EnabledChanged, Sub(s, e)
                                           If btn.Enabled Then
                                               btn.BackColor = BrandAccent
                                           Else
                                               btn.BackColor = BorderMedium
                                               btn.ForeColor = TextMuted
                                           End If
                                       End Sub
    End Sub

    ''' <summary>危险按钮（红色）</summary>
    Public Sub StyleButtonDanger(btn As Button)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.BackColor = BrandDanger
        btn.ForeColor = Color.White
        btn.Font = FontUiBold
        btn.Height = ButtonHeight
        btn.Cursor = Cursors.Hand
        btn.TextAlign = ContentAlignment.MiddleCenter
        btn.UseVisualStyleBackColor = False
        AddHandler btn.MouseEnter, Sub(s, e)
                                       btn.BackColor = Color.FromArgb(220, 38, 38)
                                   End Sub
        AddHandler btn.MouseLeave, Sub(s, e)
                                       btn.BackColor = BrandDanger
                                   End Sub
    End Sub

    ''' <summary>小号图标按钮</summary>
    Public Sub StyleButtonSmall(btn As Button)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 1
        btn.FlatAppearance.BorderColor = BorderMedium
        btn.BackColor = BgSurface
        btn.ForeColor = TextSecondary
        btn.Font = FontUiSmall
        btn.Height = ButtonHeightSmall
        btn.Cursor = Cursors.Hand
        btn.TextAlign = ContentAlignment.MiddleCenter
        btn.UseVisualStyleBackColor = False
        AddHandler btn.MouseEnter, Sub(s, e)
                                       btn.BackColor = BrandPrimaryLight
                                       btn.ForeColor = BrandPrimary
                                       btn.FlatAppearance.BorderColor = BrandPrimary
                                   End Sub
        AddHandler btn.MouseLeave, Sub(s, e)
                                       btn.BackColor = BgSurface
                                       btn.ForeColor = TextSecondary
                                       btn.FlatAppearance.BorderColor = BorderMedium
                                   End Sub
    End Sub

    ' ========== 输入控件样式 ==========

    ''' <summary>文本框样式</summary>
    Public Sub StyleTextBox(tb As TextBoxBase)
        tb.BackColor = BgInput
        tb.ForeColor = TextPrimary
        tb.Font = FontUi
        tb.BorderStyle = BorderStyle.FixedSingle
    End Sub

    ''' <summary>多行文本框（带滚动条）</summary>
    Public Sub StyleTextBoxMultiline(tb As TextBoxBase)
        StyleTextBox(tb)
        tb.BackColor = BgSurface
        tb.BorderStyle = BorderStyle.FixedSingle
    End Sub

    ''' <summary>RichTextBox 样式</summary>
    Public Sub StyleRichTextBox(rtb As RichTextBox)
        rtb.BackColor = BgSurface
        rtb.ForeColor = TextPrimary
        rtb.Font = New Font("Microsoft YaHei UI", 9.5!, FontStyle.Regular)
        rtb.BorderStyle = BorderStyle.FixedSingle
        rtb.DetectUrls = False
        rtb.HideSelection = False
    End Sub

    ''' <summary>ComboBox 样式</summary>
    Public Sub StyleComboBox(cmb As ComboBox, Optional flatBorder As Boolean = False)
        cmb.BackColor = BgSurface
        cmb.ForeColor = TextPrimary
        cmb.Font = FontUi
        cmb.FlatStyle = FlatStyle.Flat
        If flatBorder Then
            ' 通过 DrawItem 可以实现更美的效果，这里先做基础设定
        End If
    End Sub

    ' ========== 标签样式 ==========

    ''' <summary>标题标签</summary>
    Public Sub StyleLabelTitle(lbl As Label)
        lbl.Font = FontTitle
        lbl.ForeColor = TextPrimary
        lbl.AutoSize = True
    End Sub

    ''' <summary>段落标题</summary>
    Public Sub StyleLabelHeading(lbl As Label)
        lbl.Font = FontHeading
        lbl.ForeColor = TextPrimary
        lbl.AutoSize = True
    End Sub

    ''' <summary>正文标签</summary>
    Public Sub StyleLabelBody(lbl As Label)
        lbl.Font = FontUi
        lbl.ForeColor = TextPrimary
        lbl.AutoSize = True
    End Sub

    ''' <summary>辅助文字标签</summary>
    Public Sub StyleLabelHint(lbl As Label)
        lbl.Font = FontUiSmall
        lbl.ForeColor = TextSecondary
        lbl.AutoSize = True
    End Sub

    ''' <summary>状态标签</summary>
    Public Sub StyleLabelStatus(lbl As Label)
        lbl.Font = FontUiSmall
        lbl.ForeColor = TextMuted
        lbl.AutoSize = False
        lbl.TextAlign = ContentAlignment.MiddleLeft
    End Sub

    ' ========== 面板/容器样式 ==========

    ''' <summary>卡片容器</summary>
    Public Sub StylePanelCard(pnl As Panel)
        pnl.BackColor = BgSurface
        pnl.Padding = New Padding(SpacingMd)
    End Sub

    ''' <summary>FlowLayoutPanel 统一</summary>
    Public Sub StyleFlowPanel(flp As FlowLayoutPanel)
        flp.BackColor = Color.Transparent
    End Sub

    ' ========== 表单/对话框样式 ==========

    ''' <summary>对话框基础样式</summary>
    Public Sub StyleFormDialog(frm As Form)
        frm.BackColor = BgPage
        frm.Font = FontUi
        frm.StartPosition = FormStartPosition.CenterParent
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False
        frm.MinimizeBox = False
        frm.ShowIcon = False
        frm.ShowInTaskbar = False
    End Sub

    ''' <summary>为 Form 创建品牌色标题栏 Panel</summary>
    Public Function CreateFormHeader(title As String, Optional width As Integer = 400) As Panel
        Dim header As New Panel()
        header.Height = 48
        header.Width = width
        header.BackColor = BrandPrimary
        header.Padding = New Padding(SpacingLg, 0, SpacingLg, 0)

        Dim lbl As New Label()
        lbl.Text = title
        lbl.Font = New Font("Microsoft YaHei UI", 12.0!, FontStyle.Bold)
        lbl.ForeColor = Color.White
        lbl.AutoSize = True
        lbl.Location = New Point(SpacingLg, 12)
        header.Controls.Add(lbl)

        Return header
    End Function

    ' ========== 工具方法 ==========

    ''' <summary>创建品牌色分隔线</summary>
    Public Function CreateSeparator(width As Integer) As Label
        Dim sep As New Label()
        sep.Height = 1
        sep.Width = width
        sep.BackColor = BorderLight
        sep.BorderStyle = BorderStyle.None
        sep.AutoSize = False
        Return sep
    End Function

    ''' <summary>创建间距占位</summary>
    Public Function CreateSpacer(height As Integer) As Label
        Dim spacer As New Label()
        spacer.Height = height
        spacer.AutoSize = False
        Return spacer
    End Function

    ''' <summary>创建选项卡片 (用于替代 ComboBox 的模式选择)</summary>
    Public Function CreateOptionCard(icon As String, title As String, desc As String, width As Integer, Optional height As Integer = 86) As Panel
        Dim card As New Panel()
        card.Width = width
        card.Height = Math.Max(72, height)
        card.BackColor = BgSurface
        card.Cursor = Cursors.Hand
        card.Tag = "unselected"

        ' 用 Panel.Tag 存储绘制所需的数据
        ' Tag 已用于 selected 状态，用其他方式存储文字数据
        ' 直接捕获局部变量到 Paint handler 闭包中

        AddHandler card.Paint, Sub(s, e)
            Dim pnl = DirectCast(s, Panel)
            Dim g = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

            Dim rect = New Rectangle(0, 0, pnl.Width - 1, pnl.Height - 1)
            Dim isSelected As Boolean = (pnl.Tag IsNot Nothing AndAlso pnl.Tag.ToString() = "selected")

            Using path = CreateRoundedRect(rect, CornerRadius)
                If isSelected Then
                    g.FillPath(New SolidBrush(BrandPrimaryLight), path)
                    Using pen = New Pen(BrandPrimary, 1.8F)
                        g.DrawPath(pen, path)
                    End Using
                Else
                    g.FillPath(New SolidBrush(BgSurface), path)
                    Using pen = New Pen(BorderLight, 1)
                        g.DrawPath(pen, path)
                    End Using
                End If
            End Using

            Dim iconTile = New Rectangle(SpacingMd + 2, Math.Max(10, (pnl.Height - 44) \ 2), 44, 44)
            Using tilePath = CreateRoundedRect(iconTile, 10)
                Using tileBrush As New SolidBrush(If(isSelected, BrandPrimary, BrandPrimaryLight))
                    g.FillPath(tileBrush, tilePath)
                End Using
            End Using

            Dim iconBounds = iconTile
            Dim textLeft = iconTile.Right + SpacingMd
            Dim badgeWidth = If(isSelected, 62, 0)
            Dim textWidth = Math.Max(80, pnl.Width - textLeft - SpacingLg - badgeWidth)
            Dim titleBounds = New Rectangle(textLeft, Math.Max(10, (pnl.Height - 48) \ 2), textWidth, 24)
            Dim descBounds = New Rectangle(textLeft, titleBounds.Bottom + 1, textWidth, Math.Max(24, pnl.Height - titleBounds.Bottom - 10))

            Using iconFont = New Font("Segoe UI Symbol", 17.0!, FontStyle.Regular)
                TextRenderer.DrawText(g,
                                      icon,
                                      iconFont,
                                      iconBounds,
                                      If(isSelected, Color.White, BrandPrimary),
                                      TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
            End Using

            Using titleFont = New Font("Microsoft YaHei UI", 9.5!, FontStyle.Bold)
                TextRenderer.DrawText(g,
                                      title,
                                      titleFont,
                                      titleBounds,
                                      TextPrimary,
                                      TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
            End Using

            Using descFont = New Font("Microsoft YaHei UI", 8.5!, FontStyle.Regular)
                TextRenderer.DrawText(g,
                                      desc,
                                      descFont,
                                      descBounds,
                                      TextSecondary,
                                      TextFormatFlags.Left Or TextFormatFlags.Top Or TextFormatFlags.WordBreak Or TextFormatFlags.EndEllipsis)
            End Using

            If isSelected Then
                Dim badgeRect = New Rectangle(pnl.Width - SpacingLg - 54, Math.Max(10, (pnl.Height - 24) \ 2), 54, 24)
                Using badgePath = CreateRoundedRect(badgeRect, 12)
                    Using badgeBrush As New SolidBrush(Color.White)
                        g.FillPath(badgeBrush, badgePath)
                    End Using
                End Using

                Using badgeFont = New Font("Microsoft YaHei UI", 8.0!, FontStyle.Bold)
                    TextRenderer.DrawText(g,
                                          "已选",
                                          badgeFont,
                                          badgeRect,
                                          BrandPrimary,
                                          TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
                End Using
            End If
        End Sub

        Return card
    End Function

    ''' <summary>标记卡片为选中/未选中</summary>
    Public Sub SetCardSelected(card As Panel, selected As Boolean)
        card.Tag = If(selected, "selected", "unselected")
        card.Refresh()
    End Sub

    ''' <summary>创建圆角矩形路径</summary>
    Private Function CreateRoundedRect(rect As Rectangle, radius As Integer) As GraphicsPath
        Dim path As New GraphicsPath()
        Dim r = Math.Min(radius, Math.Min(rect.Width, rect.Height) \ 2)
        path.AddArc(rect.X, rect.Y, r * 2, r * 2, 180, 90)
        path.AddArc(rect.Right - r * 2, rect.Y, r * 2, r * 2, 270, 90)
        path.AddArc(rect.Right - r * 2, rect.Bottom - r * 2, r * 2, r * 2, 0, 90)
        path.AddArc(rect.X, rect.Bottom - r * 2, r * 2, r * 2, 90, 90)
        path.CloseFigure()
        Return path
    End Function

End Module
