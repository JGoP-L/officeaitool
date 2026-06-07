Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports ShareRibbon

Public Class TemplateSelectionForm
    Inherits Form

    Private ReadOnly _templates As New List(Of DocmeeTemplateInfo)()
    Private ReadOnly _buildCoverUrl As Func(Of String, String)
    Private ReadOnly _pageLoader As Func(Of Integer, List(Of DocmeeTemplateInfo))
    Private ReadOnly _pageSize As Integer
    Private ReadOnly _listBox As New ListBox()
    Private ReadOnly _previewBox As New PictureBox()
    Private ReadOnly _titleLabel As New Label()
    Private ReadOnly _pageLabel As New Label()
    Private ReadOnly _prevButton As New Button()
    Private ReadOnly _nextButton As New Button()
    Private ReadOnly _okButton As New Button()
    Private ReadOnly _imageCache As New Dictionary(Of String, Image)()
    Private _hasNextPage As Boolean

    Public Property SelectedTemplate As DocmeeTemplateInfo
    Public Property CurrentPage As Integer
    Public Property ShouldGenerateAndImport As Boolean
    Public ReadOnly Property HasNextPage As Boolean
        Get
            Return _hasNextPage
        End Get
    End Property
    Public ReadOnly Property CurrentTemplates As List(Of DocmeeTemplateInfo)
        Get
            Return New List(Of DocmeeTemplateInfo)(_templates)
        End Get
    End Property

    Public Sub New(templates As IEnumerable(Of DocmeeTemplateInfo),
                   selectedTemplateId As String,
                   buildCoverUrl As Func(Of String, String),
                   currentPage As Integer,
                   pageSize As Integer,
                   pageLoader As Func(Of Integer, List(Of DocmeeTemplateInfo)))
        _buildCoverUrl = buildCoverUrl
        _pageLoader = pageLoader
        _pageSize = Math.Max(1, pageSize)
        CurrentPage = Math.Max(1, currentPage)

        If templates IsNot Nothing Then
            For Each template In templates
                If template IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(template.Id) Then
                    _templates.Add(template)
                End If
            Next
        End If

        _hasNextPage = _pageLoader IsNot Nothing AndAlso _templates.Count >= _pageSize
        BuildLayout()
        PopulateList(selectedTemplateId)
    End Sub

    Private Sub BuildLayout()
        OfficeAIStyleHelper.StyleFormDialog(Me)
        Text = "选择模板"
        ClientSize = New Size(960, 580)
        BackColor = OfficeAIStyleHelper.BgPage

        ' 品牌色标题栏
        Dim header = OfficeAIStyleHelper.CreateFormHeader("🎨 请选择用于生成 PPT 的模板", 960)
        Controls.Add(header)

        Dim contentY As Integer = header.Bottom + OfficeAIStyleHelper.SpacingSm

        _listBox.Location = New Point(OfficeAIStyleHelper.SpacingSm, contentY)
        _listBox.Size = New Size(280, 420)
        _listBox.DrawMode = DrawMode.OwnerDrawFixed
        _listBox.ItemHeight = 76
        _listBox.IntegralHeight = False
        _listBox.BackColor = OfficeAIStyleHelper.BgSurface
        _listBox.BorderStyle = BorderStyle.FixedSingle
        _listBox.Font = OfficeAIStyleHelper.FontUi
        AddHandler _listBox.SelectedIndexChanged, AddressOf ListBox_SelectedIndexChanged
        AddHandler _listBox.DrawItem, AddressOf ListBox_DrawItem
        Controls.Add(_listBox)

        _previewBox.Location = New Point(306, contentY)
        _previewBox.Size = New Size(630, 354)
        _previewBox.BackColor = OfficeAIStyleHelper.BgSurface
        _previewBox.BorderStyle = BorderStyle.FixedSingle
        _previewBox.SizeMode = PictureBoxSizeMode.Zoom
        Controls.Add(_previewBox)

        _titleLabel.Location = New Point(306, _previewBox.Bottom + OfficeAIStyleHelper.SpacingSm)
        _titleLabel.Size = New Size(630, 50)
        _titleLabel.AutoEllipsis = True
        OfficeAIStyleHelper.StyleLabelBody(_titleLabel)
        Controls.Add(_titleLabel)

        _prevButton.Text = "◀ 上一页"
        _prevButton.Location = New Point(OfficeAIStyleHelper.SpacingSm, contentY + 436)
        _prevButton.Size = New Size(86, OfficeAIStyleHelper.ButtonHeightSmall)
        AddHandler _prevButton.Click, AddressOf PrevButton_Click
        OfficeAIStyleHelper.StyleButtonSmall(_prevButton)
        Controls.Add(_prevButton)

        _pageLabel.Location = New Point(108, contentY + 438)
        _pageLabel.Size = New Size(88, 22)
        _pageLabel.TextAlign = ContentAlignment.MiddleCenter
        OfficeAIStyleHelper.StyleLabelHint(_pageLabel)
        Controls.Add(_pageLabel)

        _nextButton.Text = "下一页 ▶"
        _nextButton.Location = New Point(202, contentY + 436)
        _nextButton.Size = New Size(86, OfficeAIStyleHelper.ButtonHeightSmall)
        AddHandler _nextButton.Click, AddressOf NextButton_Click
        OfficeAIStyleHelper.StyleButtonSmall(_nextButton)
        Controls.Add(_nextButton)

        _okButton.Text = "生成并导入"
        _okButton.Location = New Point(610, _titleLabel.Bottom + OfficeAIStyleHelper.SpacingSm)
        _okButton.Size = New Size(120, OfficeAIStyleHelper.ButtonHeight)
        _okButton.DialogResult = DialogResult.OK
        AddHandler _okButton.Click, AddressOf GenerateAndImportButton_Click
        OfficeAIStyleHelper.StyleButtonAccent(_okButton)
        Controls.Add(_okButton)
        AcceptButton = _okButton

        Dim cancelButton As New Button() With {
            .Text = "取消",
            .Location = New Point(742, _titleLabel.Bottom + OfficeAIStyleHelper.SpacingSm),
            .Size = New Size(86, OfficeAIStyleHelper.ButtonHeight),
            .DialogResult = DialogResult.Cancel
        }
        OfficeAIStyleHelper.StyleButtonSecondary(cancelButton)
        Controls.Add(cancelButton)
        CancelButton = cancelButton

        AddHandler FormClosed, AddressOf TemplateSelectionForm_FormClosed
    End Sub

    Private Sub PopulateList(selectedTemplateId As String)
        _listBox.BeginUpdate()
        _listBox.Items.Clear()
        For Each template In _templates
            _listBox.Items.Add(template)
        Next
        _listBox.EndUpdate()

        Dim selectedIndex = 0
        If Not String.IsNullOrWhiteSpace(selectedTemplateId) Then
            For index As Integer = 0 To _listBox.Items.Count - 1
                Dim template = TryCast(_listBox.Items(index), DocmeeTemplateInfo)
                If template IsNot Nothing AndAlso String.Equals(template.Id, selectedTemplateId, StringComparison.Ordinal) Then
                    selectedIndex = index
                    Exit For
                End If
            Next
        End If

        If _listBox.Items.Count > 0 Then
            _listBox.SelectedIndex = Math.Min(selectedIndex, _listBox.Items.Count - 1)
        Else
            UpdatePreview(Nothing)
        End If

        RefreshPager()
    End Sub

    Private Sub RefreshPager()
        _pageLabel.Text = $"第 {CurrentPage} 页"
        _prevButton.Enabled = CurrentPage > 1 AndAlso _pageLoader IsNot Nothing
        _nextButton.Enabled = _hasNextPage AndAlso _pageLoader IsNot Nothing
        Dim hasSelection = _listBox.SelectedItem IsNot Nothing
        _okButton.Enabled = hasSelection
    End Sub

    Private Sub ListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim template = TryCast(_listBox.SelectedItem, DocmeeTemplateInfo)
        SelectedTemplate = template
        UpdatePreview(template)
        RefreshPager()
    End Sub

    Private Sub ListBox_DrawItem(sender As Object, e As DrawItemEventArgs)
        If e.Index < 0 OrElse e.Index >= _listBox.Items.Count Then Return
        Dim template = TryCast(_listBox.Items(e.Index), DocmeeTemplateInfo)
        If template Is Nothing Then Return

        Dim selected = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        Using backBrush As New SolidBrush(If(selected, OfficeAIStyleHelper.BrandPrimaryLight, OfficeAIStyleHelper.BgSurface))
            e.Graphics.FillRectangle(backBrush, e.Bounds)
        End Using

        ' 选中态左侧指示条
        If selected Then
            Using accentBrush As New SolidBrush(OfficeAIStyleHelper.BrandPrimary)
                e.Graphics.FillRectangle(accentBrush, e.Bounds.Left, e.Bounds.Top, 3, e.Bounds.Height)
            End Using
        End If

        Dim title = If(String.IsNullOrWhiteSpace(template.Name), template.Id, template.Name)
        Dim meta = BuildMetaText(template)
        Using titleFont As New Font(Font.FontFamily, 9.0F, FontStyle.Bold),
              metaFont As New Font(Font.FontFamily, 8.0F),
              titleBrush As New SolidBrush(OfficeAIStyleHelper.TextPrimary),
              metaBrush As New SolidBrush(OfficeAIStyleHelper.TextSecondary)
            Dim titleRect = New Rectangle(e.Bounds.Left + 12, e.Bounds.Top + 10, e.Bounds.Width - 20, 24)
            Dim metaRect = New Rectangle(e.Bounds.Left + 12, e.Bounds.Top + 38, e.Bounds.Width - 20, 24)
            TextRenderer.DrawText(e.Graphics, title, titleFont, titleRect, OfficeAIStyleHelper.TextPrimary, TextFormatFlags.EndEllipsis Or TextFormatFlags.VerticalCenter)
            TextRenderer.DrawText(e.Graphics, meta, metaFont, metaRect, OfficeAIStyleHelper.TextSecondary, TextFormatFlags.EndEllipsis Or TextFormatFlags.VerticalCenter)
        End Using
    End Sub

    Private Async Sub UpdatePreview(template As DocmeeTemplateInfo)
        ClearPreviewImage()
        If template Is Nothing Then
            _titleLabel.Text = ""
            Return
        End If

        _titleLabel.Text = If(String.IsNullOrWhiteSpace(template.Name), template.Id, template.Name) & Environment.NewLine & BuildMetaText(template)

        Dim cached As Image = Nothing
        If _imageCache.TryGetValue(template.Id, cached) AndAlso cached IsNot Nothing Then
            _previewBox.Image = New Bitmap(cached)
            Return
        End If

        _previewBox.Image = CreatePlaceholderImage(template, "封面加载中...")
        Try
            Dim coverUrl = If(_buildCoverUrl Is Nothing, template.CoverUrl, _buildCoverUrl(template.CoverUrl))
            If String.IsNullOrWhiteSpace(coverUrl) Then Throw New InvalidOperationException("模板没有封面地址。")

            Dim image = Await LoadImageAsync(coverUrl)
            If image Is Nothing OrElse SelectedTemplate Is Nothing OrElse Not String.Equals(SelectedTemplate.Id, template.Id, StringComparison.Ordinal) Then
                If image IsNot Nothing Then image.Dispose()
                Return
            End If

            If _imageCache.ContainsKey(template.Id) Then _imageCache(template.Id).Dispose()
            _imageCache(template.Id) = image
            ClearPreviewImage()
            _previewBox.Image = New Bitmap(image)
        Catch ex As Exception
            ClearPreviewImage()
            _previewBox.Image = CreatePlaceholderImage(template, "封面加载失败：" & ex.Message)
        End Try
    End Sub

    Private Shared Async Function LoadImageAsync(url As String) As Task(Of Image)
        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(12)
            Dim bytes = Await client.GetByteArrayAsync(url).ConfigureAwait(False)
            Using stream As New MemoryStream(bytes)
                Using loadedImage As Image = Image.FromStream(stream)
                    Return New Bitmap(loadedImage)
                End Using
            End Using
        End Using
    End Function

    Private Sub PrevButton_Click(sender As Object, e As EventArgs)
        LoadPage(CurrentPage - 1)
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs)
        LoadPage(CurrentPage + 1)
    End Sub

    Private Sub LoadPage(page As Integer)
        If _pageLoader Is Nothing Then Return
        Dim safePage = Math.Max(1, page)
        Cursor = Cursors.WaitCursor
        Try
            Dim loaded = _pageLoader(safePage)
            _templates.Clear()
            If loaded IsNot Nothing Then
                For Each template In loaded
                    If template IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(template.Id) Then _templates.Add(template)
                Next
            End If

            CurrentPage = safePage
            _hasNextPage = _templates.Count >= _pageSize
            PopulateList("")
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub GenerateAndImportButton_Click(sender As Object, e As EventArgs)
        SelectedTemplate = TryCast(_listBox.SelectedItem, DocmeeTemplateInfo)
        If SelectedTemplate Is Nothing Then
            DialogResult = DialogResult.None
            MessageBox.Show("请先选择模板。", "预览模板", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        ShouldGenerateAndImport = True
        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub ClearPreviewImage()
        If _previewBox.Image IsNot Nothing Then
            Dim oldImage = _previewBox.Image
            _previewBox.Image = Nothing
            oldImage.Dispose()
        End If
    End Sub

    Private Function CreatePlaceholderImage(template As DocmeeTemplateInfo, statusText As String) As Image
        Dim bitmap As New Bitmap(620, 349)
        Using graphics As Graphics = Graphics.FromImage(bitmap)
            graphics.SmoothingMode = SmoothingMode.AntiAlias
            graphics.Clear(OfficeAIStyleHelper.BgSurface)
            Using borderPen As New Pen(OfficeAIStyleHelper.BorderLight),
                  accentBrush As New SolidBrush(OfficeAIStyleHelper.BrandPrimary),
                  titleFont As New Font(Font.FontFamily, 18.0F, FontStyle.Bold),
                  statusFont As New Font(Font.FontFamily, 11.0F),
                  titleBrush As New SolidBrush(OfficeAIStyleHelper.TextPrimary),
                  statusBrush As New SolidBrush(OfficeAIStyleHelper.TextSecondary)
                graphics.DrawRectangle(borderPen, 0, 0, bitmap.Width - 1, bitmap.Height - 1)
                graphics.FillRectangle(accentBrush, 0, 0, 8, bitmap.Height)
                Dim title = If(template Is Nothing OrElse String.IsNullOrWhiteSpace(template.Name), "模板预览", template.Name)
                TextRenderer.DrawText(graphics, title, titleFont, New Rectangle(34, 112, 552, 58), OfficeAIStyleHelper.TextPrimary, TextFormatFlags.EndEllipsis Or TextFormatFlags.VerticalCenter)
                TextRenderer.DrawText(graphics, statusText, statusFont, New Rectangle(34, 184, 552, 54), OfficeAIStyleHelper.TextSecondary, TextFormatFlags.WordBreak Or TextFormatFlags.EndEllipsis)
            End Using
        End Using
        Return bitmap
    End Function

    Private Shared Function BuildMetaText(template As DocmeeTemplateInfo) As String
        If template Is Nothing Then Return ""
        Dim parts As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(template.Category) Then parts.Add(template.Category.Trim())
        If Not String.IsNullOrWhiteSpace(template.Style) Then parts.Add(template.Style.Trim())
        If parts.Count = 0 Then Return If(template.Id, "")
        Return String.Join(" / ", parts)
    End Function

    Private Sub TemplateSelectionForm_FormClosed(sender As Object, e As FormClosedEventArgs)
        ClearPreviewImage()
        For Each pair In _imageCache
            If pair.Value IsNot Nothing Then pair.Value.Dispose()
        Next
        _imageCache.Clear()
    End Sub
End Class
