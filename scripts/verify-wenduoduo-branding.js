const fs = require('fs');

function read(path) {
  return fs.readFileSync(path, 'utf8');
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

const baseRibbonDesigner = read('ShareRibbon/Ribbon/BaseOfficeRibbon.Designer.vb');
const baseRibbon = read('ShareRibbon/Ribbon/BaseOfficeRibbon.vb');
const shareProject = read('ShareRibbon/ShareRibbon.vbproj');
const legacyA = ['Deep', 'seek'].join('');
const legacyAAlt = ['Deep', 'Seek'].join('');
const legacyALower = legacyA.toLowerCase();
const legacyB = ['Dou', 'bao'].join('');
const legacyBLower = legacyB.toLowerCase();
const legacyPattern = new RegExp(`${legacyA}|${legacyAAlt}|${legacyB}|${legacyALower}|${legacyBLower}`);

const appFiles = [
  'ExcelAi/ThisAddIn.vb',
  'WordAi/ThisAddIn.vb',
  'PowerPointAi/ThisAddIn.vb',
];

const ribbonFiles = [
  'ExcelAi/Ribbon1.Designer.vb',
  'WordAi/Ribbon1.Designer.vb',
  'PowerPointAi/Ribbon1.Designer.vb',
  'ExcelAi/Ribbon1.vb',
  'WordAi/Ribbon1.vb',
  'PowerPointAi/Ribbon1.vb',
];

assert(baseRibbonDesigner.includes('Me.TabAI.Label = "wenduoduoAI"'), 'shared ribbon tab label must be wenduoduoAI');
assert(baseRibbonDesigner.includes('Me.GroupConfig.Label = "模型配置"'), 'config group label must be 模型配置');
assert(baseRibbonDesigner.includes('Me.ConfigApiButton.Label = "配置模型"'), 'config button label must be 配置模型');
assert(!baseRibbonDesigner.includes(`Me.TabAI.Groups.Add(Me.Group${legacyA})`), 'legacy provider ribbon group must not be added');
assert(!baseRibbonDesigner.includes(`${legacyA}Button`), 'shared ribbon designer must not reference legacy provider button');
assert(!baseRibbonDesigner.includes(`${legacyB}Button`), 'shared ribbon designer must not reference legacy provider button');

assert(baseRibbon.includes('New SimpleOpenAIConfigForm'), 'config button must open SimpleOpenAIConfigForm');
assert(!baseRibbon.includes(`Protected MustOverride Sub ${legacyA}Button_Click`), 'base ribbon must not require legacy provider override');
assert(!baseRibbon.includes(`Protected MustOverride Sub ${legacyB}Button_Click`), 'base ribbon must not require legacy provider override');
assert(!baseRibbonDesigner.includes('AboutButton'), 'shared ribbon designer must not include About button');
assert(!baseRibbonDesigner.includes('StudyButton'), 'shared ribbon designer must not include teaching document button');
assert(!baseRibbonDesigner.includes('GroupHelp'), 'shared ribbon designer must not include help group');
assert(!baseRibbonDesigner.includes('教学文档'), 'shared ribbon designer must not include teaching document label');
assert(!baseRibbonDesigner.includes('关于'), 'shared ribbon designer must not include about label');
assert(!baseRibbon.includes('AboutButton_Click'), 'base ribbon must not handle About button');
assert(!baseRibbon.includes('StudyButton_Click'), 'base ribbon must not handle teaching document button');
assert(!shareProject.includes('Config\\AboutForm.vb'), 'AboutForm must not be compiled');
assert(!fs.existsSync('ShareRibbon/Config/AboutForm.vb'), 'AboutForm.vb must be removed');
assert(!fs.existsSync('ShareRibbon/Resources/about.png'), 'about icon must be removed');
assert(!fs.existsSync('ShareRibbon/Resources/help.png'), 'teaching document icon must be removed');
assert(!read('ShareRibbon/Resources/ShareResources.vb').match(/About|Help/), 'shared resources helper must not expose About or Help icons');
assert(!read('ShareRibbon/My Project/Resources.resx').match(/name="about"|name="help"|about\.png|help\.png/), 'resources manifest must not include About or teaching document assets');
assert(!read('ShareRibbon/My Project/Resources.Designer.vb').match(/Property about|Property help|GetObject\("about"|GetObject\("help"/), 'resources designer must not include About or teaching document assets');

assert(fs.existsSync('ShareRibbon/Config/SimpleOpenAIConfigForm.vb'), 'SimpleOpenAIConfigForm.vb must exist');
assert(shareProject.includes('Config\\SimpleOpenAIConfigForm.vb'), 'SimpleOpenAIConfigForm.vb must be included in ShareRibbon.vbproj');
assert(read('ShareRibbon/Config/ConfigManager.vb').includes('.pltform = "wenduoduoAI"'), 'default config platform must be wenduoduoAI');
assert(!read('ShareRibbon/Config/PresetProviders.vb').match(legacyPattern), 'preset providers must not expose legacy providers');
assert(!read('ShareRibbon/Resources/ShareResources.vb').match(legacyPattern), 'shared resources helper must not expose legacy provider icons');
assert(!read('ShareRibbon/My Project/Resources.resx').match(legacyPattern), 'resources manifest must not include legacy provider assets');
assert(!read('ShareRibbon/My Project/Resources.Designer.vb').match(legacyPattern), 'resources designer must not include legacy provider assets');

for (const path of appFiles) {
  const source = read(path);
  assert(source.includes('wenduoduoAI智能助手'), `${path} must use wenduoduoAI task pane title`);
  assert(!source.includes(`${legacyA} AI智能助手`), `${path} must not create legacy provider task pane`);
  assert(!source.includes(`${legacyB} AI智能助手`), `${path} must not create legacy provider task pane`);
}

for (const path of ribbonFiles) {
  const source = read(path);
  assert(!source.includes(`${legacyA}Button`), `${path} must not reference legacy provider button`);
  assert(!source.includes(`${legacyB}Button`), `${path} must not reference legacy provider button`);
  assert(!source.includes('AboutButton'), `${path} must not reference AboutButton`);
  assert(!source.includes('StudyButton'), `${path} must not reference StudyButton`);
}

for (const path of ['ExcelAi/Ribbon1.Designer.vb', 'WordAi/Ribbon1.Designer.vb', 'PowerPointAi/Ribbon1.Designer.vb']) {
  const source = read(path);
  assert(source.includes('Me.TabAI.Label = "wenduoduoAI"'), `${path} must set tab label to wenduoduoAI`);
}

for (const path of [
  `ExcelAi/${legacyA}Control.vb`,
  `ExcelAi/${legacyB}Chat.vb`,
  `WordAi/${legacyA}Control.vb`,
  `WordAi/${legacyB}Chat.vb`,
  `PowerPointAi/${legacyA}Control.vb`,
  `PowerPointAi/${legacyB}Chat.vb`,
  `ShareRibbon/Controls/Base${legacyA}Chat.vb`,
  `ShareRibbon/Controls/Base${legacyA}Chat.Designer.vb`,
  `ShareRibbon/Controls/Base${legacyB}Chat.vb`,
  `ShareRibbon/Controls/Base${legacyB}Chat.Designer.vb`,
  `ShareRibbon/Resources/${legacyALower}.png`,
  `ShareRibbon/Resources/${legacyBLower}_avatar.png`,
]) {
  assert(!fs.existsSync(path), `${path} must be removed`);
}

for (const path of [
  'ExcelAi/ExcelAi.vbproj',
  'WordAi/WordAi.vbproj',
  'PowerPointAi/PowerPointAi.vbproj',
  'ShareRibbon/ShareRibbon.vbproj',
]) {
  const source = read(path);
  assert(!source.match(legacyPattern), `${path} must not reference legacy provider files`);
}

console.log('wenduoduo branding checks passed');
