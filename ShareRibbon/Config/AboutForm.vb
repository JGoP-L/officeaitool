' ShareRibbon\Config\AboutForm.vb
Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' 关于对话框 - 显示插件信息和开源地址
''' </summary>
Public Class AboutForm
    Inherits Form

    Private lblTitle As Label
    Private lblDescription As Label
    Private lblDataPath As Label
    Private lblGithub As LinkLabel
    Private btnClose As Button

    Public Sub New()
        InitializeComponents()
    End Sub

    Private Sub InitializeComponents()
        Me.Text = "关于 Office AI 助手"
        Me.Size = New Size(450, 340)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White

        ' 标题
        lblTitle = New Label()
        lblTitle.Text = "Office AI 助手"
        lblTitle.Font = New Font("微软雅黑", 16, FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(74, 111, 165)
        lblTitle.Location = New Point(20, 20)
        lblTitle.AutoSize = True
        Me.Controls.Add(lblTitle)

        ' 描述
        lblDescription = New Label()
        lblDescription.Text = "Office AI 助手是面向 Excel、Word、PowerPoint 的" & vbCrLf &
                             "Windows 办公插件套件，用于提供 AI 聊天、文档处理、" & vbCrLf &
                             "数据分析、内容生成、翻译、排版和 MCP 工具调用能力。" & vbCrLf & vbCrLf &
                             "使用前请在设置中配置大模型 API。"
        lblDescription.Font = New Font("微软雅黑", 9)
        lblDescription.ForeColor = Color.FromArgb(80, 80, 80)
        lblDescription.Location = New Point(20, 55)
        lblDescription.Size = New Size(400, 130)
        Me.Controls.Add(lblDescription)

        ' 数据路径
        lblDataPath = New Label()
        lblDataPath.Text = "数据存放目录: 我的文档\" & ConfigSettings.OfficeAiAppDataFolder
        lblDataPath.Font = New Font("微软雅黑", 9)
        lblDataPath.ForeColor = Color.Gray
        lblDataPath.Location = New Point(20, 170)
        lblDataPath.AutoSize = True
        Me.Controls.Add(lblDataPath)

        ' 开源地址标题
        Dim lblOpenSource As New Label()
        lblOpenSource.Text = "开源地址:"
        lblOpenSource.Font = New Font("微软雅黑", 9, FontStyle.Bold)
        lblOpenSource.Location = New Point(20, 205)
        lblOpenSource.AutoSize = True
        Me.Controls.Add(lblOpenSource)

        ' Github链接
        lblGithub = New LinkLabel()
        lblGithub.Text = "Github: https://github.com/JGoP-L/officeAI"
        lblGithub.Font = New Font("微软雅黑", 9)
        lblGithub.Location = New Point(20, 230)
        lblGithub.AutoSize = True
        lblGithub.LinkColor = Color.FromArgb(74, 111, 165)
        AddHandler lblGithub.LinkClicked, AddressOf Github_LinkClicked
        Me.Controls.Add(lblGithub)

        ' 关闭按钮
        btnClose = New Button()
        btnClose.Text = "关闭"
        btnClose.Size = New Size(80, 30)
        btnClose.Location = New Point(350, 270)
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.BackColor = Color.FromArgb(74, 111, 165)
        btnClose.ForeColor = Color.White
        btnClose.Font = New Font("微软雅黑", 9)
        AddHandler btnClose.Click, AddressOf BtnClose_Click
        Me.Controls.Add(btnClose)
        Me.AcceptButton = btnClose
    End Sub

    Private Sub Github_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Try
            System.Diagnostics.Process.Start("https://github.com/JGoP-L/officeAI")
        Catch ex As Exception
            MessageBox.Show("无法打开链接: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
End Class
