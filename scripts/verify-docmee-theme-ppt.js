const fs = require('fs');

function read(path) {
  return fs.readFileSync(path, 'utf8');
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

const ribbonDesigner = read('PowerPointAi/Ribbon1.Designer.vb');
const ribbon = read('PowerPointAi/Ribbon1.vb');
const addIn = read('PowerPointAi/ThisAddIn.vb');
const project = read('PowerPointAi/PowerPointAi.vbproj');
const client = fs.existsSync('PowerPointAi/DocmeePptClient.vb') ? read('PowerPointAi/DocmeePptClient.vb') : '';
const pane = fs.existsSync('PowerPointAi/ThemePptTaskPane.vb') ? read('PowerPointAi/ThemePptTaskPane.vb') : '';

assert(ribbonDesigner.includes('Me.TemplateFormatButton.Visible = True'), '主题生成PPT button must be visible');
assert(ribbonDesigner.includes('Me.TemplateFormatButton.Label = "主题生成PPT"'), 'TemplateFormatButton must be relabeled 主题生成PPT');
assert(ribbon.includes('Globals.ThisAddIn.ShowThemePptTaskPane()'), 'Ribbon must open the theme PPT task pane');

assert(addIn.includes('themePptTaskPane'), 'ThisAddIn must keep a theme PPT task pane');
assert(addIn.includes('ShowThemePptTaskPane'), 'ThisAddIn must expose ShowThemePptTaskPane');
assert(addIn.includes('New ThemePptTaskPane(Me.Application)'), 'ThisAddIn must create ThemePptTaskPane');

assert(project.includes('Compile Include="DocmeePptClient.vb"'), 'PowerPoint project must compile DocmeePptClient');
assert(project.includes('Compile Include="ThemePptTaskPane.vb"'), 'PowerPoint project must compile ThemePptTaskPane');

assert(client.includes('https://test.docmee.cn'), 'Docmee client must use the test API base URL');
assert(client.includes('Private Const DemoToken As String = "ak_demo"'), 'Docmee client must use ak_demo token for demo');
assert(client.includes('/api/ppt/v2/createTask'), 'Docmee client must call createTask');
assert(client.includes('MultipartFormDataContent'), 'createTask must use multipart/form-data');
assert(client.includes('/api/ppt/v2/generateContent'), 'Docmee client must call generateContent');
assert(client.includes('"outlineType", "JSON"'), 'generateContent must request JSON outline');
assert(client.includes('"questionMode", False'), 'generateContent must disable questionMode');
assert(client.includes('"isNeedAsk", False'), 'generateContent must disable isNeedAsk');
assert(client.includes('/api/ppt/templates'), 'Docmee client must include the template list endpoint for later template selection');

assert(pane.includes('Class ThemePptTaskPane'), 'ThemePptTaskPane class must exist');
assert(pane.includes('CreateTaskAsync'), 'ThemePptTaskPane must create a Docmee task');
assert(pane.includes('GenerateContentAsync'), 'ThemePptTaskPane must generate Docmee outline content');
assert(pane.includes('InsertOutlineIntoPresentation'), 'ThemePptTaskPane must insert generated outline into PPT');
assert(pane.includes('children'), 'ThemePptTaskPane must consume outline children');
assert(pane.includes('pages'), 'ThemePptTaskPane must consume Docmee pages responses');
assert(pane.includes('overall_theme'), 'ThemePptTaskPane must use Docmee overall_theme as a title candidate');
assert(pane.includes('ppLayoutTitleOnly'), 'ThemePptTaskPane must create title/content slides');

console.log('docmee theme ppt checks passed');
