const fs = require('fs');

function read(path) {
  return fs.readFileSync(path, 'utf8');
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

const pptRibbonDesigner = read('PowerPointAi/Ribbon1.Designer.vb');
const pptRibbon = read('PowerPointAi/Ribbon1.vb');

function captureBetween(source, start, end, message) {
  const startIndex = source.indexOf(start);
  assert(startIndex !== -1, `${message}: missing start marker`);
  const endIndex = source.indexOf(end, startIndex + start.length);
  assert(endIndex !== -1, `${message}: missing end marker`);
  return source.slice(startIndex, endIndex);
}

assert(pptRibbonDesigner.includes('Me.ContinuationButton.Label = "文本优化"'), 'PPT ribbon must expose 文本优化 button');
assert(pptRibbonDesigner.includes('Me.ReformatButton.Label = "美化单页"'), 'PPT ribbon must expose 美化单页 button');
assert(pptRibbonDesigner.includes('Me.ReformatButton.Visible = True'), 'PPT beautify button must be visible');
assert(pptRibbonDesigner.includes('选中文字后可润色、扩写、精简、填充或补全文案'), 'PPT text optimize tooltip must describe expand, shorten, fill, and complete behavior');
assert(pptRibbonDesigner.includes('对当前页应用排版美化'), 'PPT beautify tooltip must describe current-slide beautification behavior');
assert(!pptRibbonDesigner.includes('演示版'), 'PowerPoint ribbon copy must not expose demo-only wording');

assert(pptRibbon.includes('Protected Overrides Async Sub ContinuationButton_Click'), '文本优化 handler must be async');
assert(pptRibbon.includes('ReplaceCurrentSlideWithGeneratedTextAsync'), '替换单页 must call the generated-slide replacement workflow');
const replaceGenerationBlock = captureBetween(
  pptRibbon,
  'Private Async Function GenerateReplacementSlideTextAsync(requirement As String) As Task(Of String)',
  'Private Function InsertReplacementSlideAfter',
  'replacement slide generation'
);
assert(replaceGenerationBlock.includes('请先在模型配置里填写 API 地址、API Key 和模型名称'), '替换单页 must require model configuration before generating a replacement page');
assert(!replaceGenerationBlock.includes('Return requirement.Trim()'), '替换单页 must not fake AI generation by using the user requirement as generated page content');
assert(replaceGenerationBlock.includes('Throw New InvalidOperationException("模型没有返回可用的替换单页内容。")'), '替换单页 must fail clearly if the model returns no usable page content');
assert(pptRibbon.includes('ShowTextOptimizeDialog()'), '文本优化 must show a mode dialog');
assert(pptRibbon.includes('OptimizeSelectedTextAsync'), '文本优化 must call selected text optimization');
assert(pptRibbon.includes('BuildTextOptimizationPrompt'), '文本优化 must build an AI prompt locally');
assert(pptRibbon.includes('LLMUtil.CreateLlmRequestBody'), '文本优化 must use the configured model API');
assert(pptRibbon.includes('GetSelectedPptTextTargets(modeName)'), '文本优化 must read selected PPT text with mode awareness');
assert(pptRibbon.includes('target.TextRange.Text = optimizedText'), '文本优化 must replace the selected text');
assert(pptRibbon.includes('"填充"'), '文本优化 must expose a 填充 mode matching the requested feature');
assert(pptRibbon.includes('Case "填充"'), '文本优化 must implement a dedicated fill prompt branch');
assert(pptRibbon.includes('填充当前 PPT 文本框'), '文本优化 fill mode must target filling the current PPT text box');
assert(pptRibbon.includes('Dim allowBlankTextFrame = String.Equals(modeName, "填充"'), '填充 mode must allow selected blank text boxes');
assert(pptRibbon.includes('CollectShapeTextTargets(sel.ShapeRange(i), targets, allowBlankTextFrame, slideContextText)'), 'selected shape collection must receive the fill-mode blank-text flag and slide context');
assert(pptRibbon.includes('If allowBlankTextFrame OrElse Not String.IsNullOrWhiteSpace(text)'), 'blank text frames must be accepted only for fill mode');
const selectedTextSelectionBlock = captureBetween(
  pptRibbon,
  'Case PowerPoint.PpSelectionType.ppSelectionText',
  'Case PowerPoint.PpSelectionType.ppSelectionShapes',
  'selected text target collection'
);
assert(selectedTextSelectionBlock.includes('If allowBlankTextFrame OrElse Not String.IsNullOrWhiteSpace(text)'), 'fill mode must also accept a blank direct text selection or cursor target');
const tableTargetBlock = captureBetween(
  pptRibbon,
  'If shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then',
  'If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then',
  'table text target collection'
);
assert(tableTargetBlock.includes('If allowBlankTextFrame OrElse Not String.IsNullOrWhiteSpace(cellText)'), 'fill mode must accept blank table cells when a table is selected');
assert(pptRibbon.includes('当前文本框为空'), 'fill prompt must explicitly handle empty selected text boxes');
assert(pptRibbon.includes('Public Property SlideContextText As String'), 'fill mode targets must carry current slide context');
assert(pptRibbon.includes('Dim slideContextText = GetCurrentSlideContextText()'), 'text optimization must collect current slide context');
assert(pptRibbon.includes('RequestOptimizedTextAsync(modeName, target.OriginalText, target.SlideContextText)'), 'text optimization requests must pass slide context');
assert(pptRibbon.includes('BuildTextOptimizationPrompt(modeName, originalText, slideContextText)'), 'text optimization prompt must include slide context');
assert(pptRibbon.includes('当前页其他内容：'), 'fill prompt must provide neighboring slide text as context');

assert(pptRibbon.includes('ApplySimpleBeautifyToCurrentSlide()'), '美化单页 must call current-slide beautify');
assert(pptRibbon.includes('wenduoduoAI_BeautifyAccent'), '美化单页 must mark demo accent shapes');
assert(pptRibbon.includes('MsoAutoShapeType.msoShapeRectangle'), '美化单页 must add the accent shape with MsoAutoShapeType');
assert(!pptRibbon.includes('MsoShapeType.msoShapeRectangle'), '美化单页 must not use MsoShapeType for AddShape');
assert(pptRibbon.includes('BeautifyShapeText'), '美化单页 must format text shapes');
assert(pptRibbon.includes('AutoFitPptTextShape'), '美化单页 must fit text after formatting');

console.log('ppt demo feature checks passed');
