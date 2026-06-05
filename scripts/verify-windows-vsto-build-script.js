const fs = require('fs');

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

const path = 'scripts/verify-powerpoint-plugin-windows.ps1';
assert(fs.existsSync(path), 'Windows VSTO verification script must exist');

const script = fs.readFileSync(path, 'utf8');

assert(script.includes('Microsoft Visual Studio\\Installer\\vswhere.exe'), 'script must locate bundled vswhere.exe');
assert(script.includes('-requires Microsoft.Component.MSBuild'), 'script must require MSBuild through vswhere');
assert(script.includes('MSBuild\\**\\Bin\\MSBuild.exe'), 'script must find MSBuild.exe with the official vswhere pattern');
assert(script.includes('PowerPointAi\\PowerPointAi.vbproj'), 'script must build the PowerPoint VSTO project directly');
assert(!script.includes('Invoke-CheckedCommand $nodeCommand.Source'), 'script must not pass $nodeCommand.Source directly in argument mode');
assert(script.includes('$nodePath = $nodeCommand.Source'), 'script must assign node command source before invoking it');
assert(script.includes('Invoke-CheckedCommand -FilePath $nodePath -Arguments @($ScriptPath)'), 'script must invoke node verifier with named parameters');
assert(script.includes('Invoke-CheckedCommand -FilePath $msbuild -Arguments @('), 'script must invoke MSBuild with named parameters');
assert(script.includes('/restore'), 'script must restore packages before building');
assert(script.includes('/t:Rebuild'), 'script must rebuild the project');
assert(script.includes('/p:Configuration='), 'script must allow Debug/Release configuration');
assert(script.includes('/p:VisualStudioVersion='), 'script must pass VisualStudioVersion for OfficeTools targets');
assert(script.includes('node scripts/verify-docmee-theme-ppt.js'), 'script must run the Docmee/PPT structural verifier');
assert(script.includes('node scripts/verify-ppt-demo-features.js'), 'script must run the PowerPoint feature verifier');
assert(script.includes('node scripts/verify-ppt-current-slide-translation.js'), 'script must run the translation verifier');
assert(script.includes('node scripts/verify-vb-block-balance.js'), 'script must run the VB block balance verifier');
assert(script.includes('node scripts/verify-docmee-api-smoke.js'), 'script must run the Docmee external API smoke test');
assert(script.includes('DOCMEE_SMOKE_GENERATE'), 'script must optionally run the generation API smoke path');
assert(script.includes('$hadSmokeGenerate = Test-Path Env:DOCMEE_SMOKE_GENERATE'), 'script must remember whether DOCMEE_SMOKE_GENERATE existed before full smoke');
assert(script.includes('Remove-Item Env:DOCMEE_SMOKE_GENERATE'), 'script must remove DOCMEE_SMOKE_GENERATE after full smoke if it did not exist before');
assert(script.includes('[switch]$PowerPointComSmoke'), 'script must expose an optional PowerPoint COM smoke test');
assert(script.includes('New-Object -ComObject PowerPoint.Application'), 'PowerPoint COM smoke must create the PowerPoint Application object');
assert(script.includes('$powerPoint.Visible = -1'), 'PowerPoint COM smoke must make PowerPoint visible before creating a presentation');
assert(script.includes('$powerPoint.Presentations.Add()'), 'PowerPoint COM smoke must create a presentation');
assert(script.includes('$presentation.Close()'), 'PowerPoint COM smoke must close the temporary presentation');
assert(script.includes('$powerPoint.Quit()'), 'PowerPoint COM smoke must quit PowerPoint after the smoke test');
assert(script.includes('PowerPointComSmoke'), 'manual checklist must document the optional PowerPoint COM smoke switch');
assert(script.includes('AI生成PPT'), 'script must include a manual Office runtime checklist for AI生成PPT');
assert(script.includes('Docmee配置') && script.includes('token'), 'manual checklist must cover Docmee API address and token configuration');
assert(script.includes('模型配置') && script.includes('API Key'), 'manual checklist must cover model configuration for translation and text optimization');
assert(script.includes('标题生成PPT') && script.includes('文档生成PPT'), 'manual checklist must cover title and document PPT generation');
assert(script.includes('大纲编辑') && script.includes('模板生成'), 'manual checklist must cover outline editing and template generation');
assert(script.includes('一键更换主题') && script.includes('替换单页') && script.includes('美化单页'), 'manual checklist must cover theme, replace, and beautify actions');
assert(script.includes('文本翻译') && script.includes('扩写') && script.includes('精简') && script.includes('填充'), 'manual checklist must cover translation and text optimization actions');

console.log('windows vsto build script checks passed');
