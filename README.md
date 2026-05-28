# Office AI 智能体

<div align="center">

[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![Office](https://img.shields.io/badge/office-Excel%20Word%20PowerPoint-green.svg)](https://www.microsoft.com/office)

**Language / 语言选择**

[English](README_EN.md) | [中文](README.md)

</div>

## 预览

![ExcelView](./AiHelper.assets/excelai_display.png)

![WordView](./AiHelper.assets/wordai_display.png)

![PPTView](./AiHelper.assets/pptai_display.png)

## 概述

Office AI 智能体是基于 Visual Studio 2022、Visual Basic.NET、.NET Framework 4.7.2 和 VSTO 开发的 Windows 办公插件套件，为 Excel、Word、PowerPoint 提供 AI 辅助能力。

## 功能特性

- AI 数据分析、公式辅助、图表生成
- Word 文档处理、续写、审阅、排版、翻译
- PowerPoint 内容生成、审阅、排版、翻译
- WebView2 + HTML/CSS/JS 的聊天界面
- MCP Client 集成，支持配置 MCP Server
- 支持多家大模型服务商配置

## 支持产品

| 产品 | 状态 | 功能 |
|------|------|------|
| Microsoft Excel | 支持 | 数据分析、图表生成、公式辅助、ALLM/CLLM 函数 |
| Microsoft Word | 支持 | 文档处理、内容生成、审阅、续写、排版、翻译 |
| Microsoft PowerPoint | 支持 | 演示文稿创建、幻灯片处理、审阅、续写、排版、翻译 |
| WPS Office Windows 版 | 兼容 | 通过 Windows 插件注册方式加载 |

## 安装与构建

### 系统要求

- Windows 10 或 Windows 11
- Microsoft Office 2016+ 或 WPS Office Windows 版
- Visual Studio 2022
- Office/SharePoint development workload
- .NET Framework 4.7.2 Developer Pack
- Microsoft Visual Studio Installer Projects 扩展

### 从源码构建

```bash
git clone https://github.com/JGoP-L/officeAI.git
```

在 Windows 上使用 Visual Studio 2022 打开 `AiHelper.sln`，还原 NuGet 包后构建解决方案。安装包项目位于 `OfficeAgent/`。

## 项目结构

```text
officeAI/
├── ExcelAi/          # Excel 插件
├── WordAi/           # Word 插件
├── PowerPointAi/     # PowerPoint 插件
├── ShareRibbon/      # 共享组件、UI、配置、AI 通信、MCP
└── OfficeAgent/      # Windows 安装包项目
```

## 使用说明

1. 安装插件后启动 Excel、Word 或 PowerPoint。
2. 在功能区打开 AI 助手。
3. 在设置里配置大模型 API。
4. 选择文档内容、单元格或幻灯片后发起提问或执行任务。

## 开源协议

本项目采用 Apache 2.0 许可证，详情见 [LICENSE](LICENSE)。
