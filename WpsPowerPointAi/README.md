# 文多多 WPS 演示加载项

这是 PowerPoint VSTO 插件的 WPS 演示版本骨架，独立放在 `WpsPowerPointAi`，避免影响现有 Office PPT 插件。

## 当前包含的功能

- Ribbon 菜单拆分：生成PPT、文档生成PPT、润色、扩写、缩写、翻译、美化单页、生成单页、一键换主题。
- 任务窗格 UI：按文多多现有紫色风格实现。
- Docmee JS 客户端：对接创建任务、生成 Markdown 大纲、模板列表、生成 PPT、下载 PPT、文本创作、AI 生成单页。
- WPS 演示适配层：封装选中文字、替换文字、一键换主题等宿主能力。

## 本地调试

1. 用静态服务器发布本目录，例如：

   ```powershell
   cd WpsPowerPointAi
   python -m http.server 3888
   ```

2. 在浏览器打开：

   ```text
   http://127.0.0.1:3888/ui/taskpane.html
   ```

3. 在 WPS 加载项开发模式中加载本目录，入口文件为：

   - `manifest.xml`
   - `jsplugins.xml`
   - `ribbon.xml`
   - `main.js`
   - `index.html`

发布给用户时请使用 `jsplugins.release.xml` 或保持 `jsplugins.xml` 中 `enable="enable"`，不要使用 `enable_dev` / `wpsjs debug`，否则 WPS 会在功能区注入“打开JS调试器”。

## 说明

- WPS 加载项使用 `ribbon.xml` 定义功能区按钮，按钮通过 `main.js` 的 `onAction` 调用打开任务窗格。
- 由于 WPS WebView 直接访问 Docmee 接口时可能遇到跨域限制，正式打包前如果接口被浏览器 CORS 拦截，需要加一层本地/服务端代理。
- PPTX 自动下载并插入 WPS 演示的能力依赖 WPS 宿主提供本地文件访问 API，目前先保留下载链接和宿主适配入口。

参考：WPS 官方加载项开发说明 <https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/wps-addin-development-instructions>
