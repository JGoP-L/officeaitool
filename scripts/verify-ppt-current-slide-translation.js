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
const translateForm = read('ShareRibbon/Translate/TranslateActionForm.vb');
const pptTranslateService = read('PowerPointAi/PowerPointDocumentTranslateService.vb');

assert(pptRibbonDesigner.includes('Me.TranslateButton.Label = "文本翻译"'), 'PowerPoint ribbon must label translate button as 文本翻译');
assert(pptRibbonDesigner.includes('自动适配文本框'), 'PowerPoint translate button must describe auto-fit beautification');

assert(translateForm.includes('Private rbCurrentSlide As RadioButton'), 'translate dialog must define current-slide scope radio');
assert(translateForm.includes('Public Property TranslateCurrentSlide As Boolean = False'), 'translate dialog must expose TranslateCurrentSlide');
assert(translateForm.includes('.Text = "当前页"'), 'translate dialog must show 当前页 scope');
assert(translateForm.includes('.Checked = (_appType = "PowerPoint" AndAlso Not _hasSelection)'), 'PowerPoint current-slide scope must be checked only when there is no selection');
assert(translateForm.includes('.Checked = (_appType = "PowerPoint" AndAlso _hasSelection)'), 'PowerPoint selected content scope must be checked by default when there is a selection');
assert(translateForm.includes('rbReplace.Checked = True'), 'PowerPoint translation must default to replace mode');
assert(translateForm.includes('TranslateCurrentSlide = (_appType = "PowerPoint"'), 'translate dialog must set TranslateCurrentSlide from current-slide radio');
for (const lang of ['"en"', '"zh"', '"ja"', '"ko"', '"fr"', '"de"', '"es"', '"ru"', '"pt"', '"it"', '"vi"', '"th"', '"id"', '"ar"']) {
  assert(translateForm.includes(lang), `translate dialog must include target language ${lang}`);
}

assert(pptRibbon.includes('actionForm.TranslateCurrentSlide'), 'PowerPoint ribbon must branch on TranslateCurrentSlide');
assert(pptRibbon.includes('translateService.TranslateCurrentSlideAsync()'), 'PowerPoint ribbon must translate current slide directly');

assert(pptTranslateService.includes('Public Async Function TranslateCurrentSlideAsync()'), 'PowerPoint translate service must expose TranslateCurrentSlideAsync');
assert(pptTranslateService.includes('Public Function GetCurrentSlideParagraphs()'), 'PowerPoint translate service must collect current slide text');
assert(pptTranslateService.includes('Private Function GetCurrentSlide() As Slide'), 'PowerPoint translate service must resolve current slide');
assert(pptTranslateService.includes('FitTranslatedShapeText(item.Shape)'), 'replace mode must auto-fit translated shapes');
assert(pptTranslateService.includes('Private Sub FitTranslatedShapeText(shape As Shape)'), 'PowerPoint translate service must implement auto-fit helper');
assert(pptTranslateService.includes('Private Function TextOverflows'), 'PowerPoint translate service must detect text overflow');
assert(pptTranslateService.includes('textRange.BoundHeight'), 'auto-fit must compare text bound height');
assert(pptTranslateService.includes('textRange.BoundWidth'), 'auto-fit must compare text bound width');
assert(pptTranslateService.includes('ppAutoSizeShapeToFitText'), 'auto-fit must fall back to PowerPoint AutoSize');

console.log('ppt current-slide translation checks passed');
