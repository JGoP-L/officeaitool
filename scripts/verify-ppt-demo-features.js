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

assert(pptRibbonDesigner.includes('Me.ContinuationButton.Label = "文本优化"'), 'PPT ribbon must expose 文本优化 button');
assert(pptRibbonDesigner.includes('Me.ReformatButton.Label = "美化单页"'), 'PPT ribbon must expose 美化单页 button');
assert(pptRibbonDesigner.includes('Me.ReformatButton.Visible = True'), 'PPT beautify button must be visible');
assert(pptRibbonDesigner.includes('选中文字后可润色、扩写、精简或补全文案'), 'PPT text optimize tooltip must describe demo behavior');
assert(pptRibbonDesigner.includes('对当前页应用演示版排版美化'), 'PPT beautify tooltip must describe demo behavior');

assert(pptRibbon.includes('Protected Overrides Async Sub ContinuationButton_Click'), '文本优化 handler must be async');
assert(pptRibbon.includes('ShowTextOptimizeDialog()'), '文本优化 must show a mode dialog');
assert(pptRibbon.includes('OptimizeSelectedTextAsync'), '文本优化 must call selected text optimization');
assert(pptRibbon.includes('BuildTextOptimizationPrompt'), '文本优化 must build an AI prompt locally');
assert(pptRibbon.includes('LLMUtil.CreateLlmRequestBody'), '文本优化 must use the configured model API');
assert(pptRibbon.includes('GetSelectedPptTextTargets'), '文本优化 must read selected PPT text');
assert(pptRibbon.includes('target.TextRange.Text = optimizedText'), '文本优化 must replace the selected text');

assert(pptRibbon.includes('ApplySimpleBeautifyToCurrentSlide()'), '美化单页 must call current-slide beautify');
assert(pptRibbon.includes('wenduoduoAI_BeautifyAccent'), '美化单页 must mark demo accent shapes');
assert(pptRibbon.includes('msoShapeRectangle'), '美化单页 must add a simple visual accent');
assert(pptRibbon.includes('BeautifyShapeText'), '美化单页 must format text shapes');
assert(pptRibbon.includes('AutoFitPptTextShape'), '美化单页 must fit text after formatting');

console.log('ppt demo feature checks passed');
