param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [string]$VisualStudioVersion = "17.0",

    [switch]$FullDocmeeSmoke,

    [switch]$PowerPointComSmoke
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $RepoRoot

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message"
}

function Invoke-CheckedCommand {
    param(
        [string]$FilePath,
        [string[]]$Arguments
    )

    Write-Host ("$FilePath " + ($Arguments -join " "))
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath"
    }
}

function Find-MSBuild {
    $vswherePath = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswherePath)) {
        $vswhereCommand = Get-Command "vswhere.exe" -ErrorAction SilentlyContinue
        if ($null -eq $vswhereCommand) {
            throw "未找到 vswhere.exe。请确认已安装 Visual Studio 2022/VSTO 工作负载或 Visual Studio Build Tools。"
        }
        $vswherePath = $vswhereCommand.Source
    }

    $msbuildPaths = & $vswherePath -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe"
    $msbuildPath = $msbuildPaths | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($msbuildPath) -or -not (Test-Path $msbuildPath)) {
        throw "vswhere.exe did not return a usable MSBuild.exe path."
    }

    return $msbuildPath
}

function Invoke-NodeVerifier {
    param([string]$ScriptPath)

    $nodeCommand = Get-Command "node" -ErrorAction Stop
    $nodePath = $nodeCommand.Source
    Invoke-CheckedCommand -FilePath $nodePath -Arguments @($ScriptPath)
}

function Release-ComObjectSafely {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Invoke-PowerPointComSmoke {
    $powerPoint = $null
    $presentation = $null

    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $powerPoint.Visible = -1
        $presentation = $powerPoint.Presentations.Add()

        if ($null -eq $presentation) {
            throw "PowerPoint did not create a temporary presentation."
        }

        Write-Host "PowerPoint COM smoke test created a temporary presentation successfully."
    }
    finally {
        if ($null -ne $presentation) {
            $presentation.Close()
            Release-ComObjectSafely $presentation
        }

        if ($null -ne $powerPoint) {
            $powerPoint.Quit()
            Release-ComObjectSafely $powerPoint
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Write-Step "Locate MSBuild"
$msbuild = Find-MSBuild
Write-Host "MSBuild: $msbuild"

Write-Step "Build PowerPoint VSTO project"
$projectPath = Join-Path $RepoRoot "PowerPointAi\PowerPointAi.vbproj"
Invoke-CheckedCommand -FilePath $msbuild -Arguments @(
    $projectPath,
    "/restore",
    "/t:Rebuild",
    "/m:1",
    "/nr:false",
    "/p:Configuration=$Configuration",
    "/p:Platform=AnyCPU",
    "/p:VisualStudioVersion=$VisualStudioVersion"
)

Write-Step "Run local feature verifiers"
Write-Host "node scripts/verify-docmee-theme-ppt.js"
Invoke-NodeVerifier "scripts/verify-docmee-theme-ppt.js"

Write-Host "node scripts/verify-ppt-demo-features.js"
Invoke-NodeVerifier "scripts/verify-ppt-demo-features.js"

Write-Host "node scripts/verify-ppt-current-slide-translation.js"
Invoke-NodeVerifier "scripts/verify-ppt-current-slide-translation.js"

Write-Host "node scripts/verify-vb-block-balance.js"
Invoke-NodeVerifier "scripts/verify-vb-block-balance.js"

Write-Step "Run Docmee API smoke test"
Write-Host "node scripts/verify-docmee-api-smoke.js"
Invoke-NodeVerifier "scripts/verify-docmee-api-smoke.js"

if ($FullDocmeeSmoke) {
    Write-Step "Run full Docmee generation smoke test"
    $hadSmokeGenerate = Test-Path Env:DOCMEE_SMOKE_GENERATE
    $oldSmokeGenerate = $env:DOCMEE_SMOKE_GENERATE
    try {
        $env:DOCMEE_SMOKE_GENERATE = "1"
        Write-Host "DOCMEE_SMOKE_GENERATE=1 node scripts/verify-docmee-api-smoke.js"
        Invoke-NodeVerifier "scripts/verify-docmee-api-smoke.js"
    }
    finally {
        if ($hadSmokeGenerate) {
            $env:DOCMEE_SMOKE_GENERATE = $oldSmokeGenerate
        }
        else {
            Remove-Item Env:DOCMEE_SMOKE_GENERATE -ErrorAction SilentlyContinue
        }
    }
}

if ($PowerPointComSmoke) {
    Write-Step "Run optional PowerPoint COM smoke test"
    Invoke-PowerPointComSmoke
}

Write-Step "Manual Office runtime checklist"
$manualChecklist = @"
请在 Windows PowerPoint 中继续手工验证：
1. 打开插件 Ribbon，确认 AI生成PPT 入口可打开任务窗格。
2. 打开 Docmee配置，确认接口地址和 token 正确；未填时应使用 app.config、环境变量或测试默认值。
3. 打开模型配置，确认 API 地址、API Key、模型名称可用于文本翻译和文本优化。
4. 标题生成PPT：输入标题，生成 Markdown 大纲，确认大纲编辑框可修改且预览同步。
5. 文档生成PPT：选择支持的 docx/pdf/pptx/xlsx/md 等文件，生成 Markdown 大纲。
6. 大纲编辑：修改 Markdown 后，确认必须点“完成编辑”才能预览/选择模板生成。
7. 模板生成：选择模板后，用编辑后的 Markdown 生成并导入 PPT。
8. 一键更换主题：对刚导入的 Docmee PPT 调用更换主题，并确认当前演示文稿页被替换。
9. 替换单页：选中当前页，输入要求，确认新页替换旧页。
10. 美化单页：选中当前页，确认标题、正文、背景和文本框适配被应用。
11. 文本翻译：验证当前页/选中内容翻译，并确认译文替换后文本框不溢出。
12. 文本优化：分别验证润色、扩写、精简、填充、补全文案，尤其空文本框填充能利用当前页上下文。

可选自动烟测：在有桌面 Office 的 Windows 机器上执行
`powershell -ExecutionPolicy Bypass -File scripts\verify-powerpoint-plugin-windows.ps1 -FullDocmeeSmoke -PowerPointComSmoke`
它会启动 PowerPoint、创建一个临时空演示，然后关闭演示并退出 PowerPoint。
"@
Write-Host $manualChecklist

Write-Host ""
Write-Host "Windows VSTO verification script completed."
