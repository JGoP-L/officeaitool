(function (global) {
  "use strict";

  var DOCMEE_CONFIG = {
    baseUrl: "https://test.docmee.cn",
    token: "ak_demo",
    companyId: "203361",
    lang: "zh",
    pptxId: ""
  };

  var state = {
    mode: "generate-ppt",
    selectedFile: null,
    taskId: "",
    markdown: "",
    templateId: "",
    templatePage: 1,
    templateThemeColor: "",
    generatedPptInfo: null,
    generatedPptUrl: "",
    outlineSourceMode: "generate-ppt",
    isGeneratingOutline: false,
    isGeneratingDocumentOutline: false,
    textType: "rs",
    lastTextResult: "",
    isGeneratingSinglePage: false,
    isApplyingSinglePage: false,
    autoBeautifyStarted: false,
    singlePageChoices: [],
    singlePageSelectedIndex: 0,
    singlePageResultInfo: null
  };

  var MODE_META = {
    "generate-ppt": {
      title: "AI 生成 PPT",
      subtitle: "输入主题后生成大纲，再选择模板生成演示文稿",
      panel: "generate-ppt"
    },
    "document-generate": {
      title: "文档生成 PPT",
      subtitle: "选择文档后自动解析内容并生成演示文稿",
      panel: "document-generate"
    },
    "text-polish": {
      title: "润色",
      subtitle: "优化当前选中文本，让表达更自然专业",
      panel: "text-create",
      textType: "rs",
      textTitle: "润色",
      textHint: "优化表达，让文字更自然、专业。",
      actionText: "开始润色",
      placeholder: "选中文字后打开，或直接输入要润色的文本"
    },
    "text-expand": {
      title: "扩写",
      subtitle: "补充细节，把简短表达扩展得更完整",
      panel: "text-create",
      textType: "kx",
      textTitle: "扩写",
      textHint: "补充细节，让内容更完整、更有信息量。",
      actionText: "开始扩写",
      placeholder: "选中文字后打开，或直接输入要扩写的文本"
    },
    "text-shorten": {
      title: "缩写",
      subtitle: "压缩冗余内容，保留核心信息和重点",
      panel: "text-create",
      textType: "sx",
      textTitle: "缩写",
      textHint: "精简文字，保留核心信息和重点。",
      actionText: "开始缩写",
      placeholder: "选中文字后打开，或直接输入要缩写的文本"
    },
    "text-translate": {
      title: "翻译",
      subtitle: "选择目标语言，翻译并优化当前文本",
      panel: "text-create",
      textType: "fy",
      textTitle: "翻译",
      textHint: "选择目标语言，保留原意并优化表达。",
      actionText: "开始翻译",
      placeholder: "选中文字后打开，或直接输入要翻译的文本"
    },
    "generate-page": {
      title: "AI 生成单页",
      subtitle: "输入页面标题或内容，生成新的单页候选",
      panel: "single-page",
      pageTitle: "AI 生成单页",
      pageHint: "输入页面标题或内容，我会生成一页新的 PPT 页面。",
      pagePlaceholder: "请输入页面标题或具体内容",
      pageActionText: "AI智能生成新页面",
      beautify: false
    },
    "beautify-page": {
      title: "美化单页",
      subtitle: "基于当前页内容生成更美观的单页版本",
      panel: "single-page",
      pageTitle: "美化单页",
      pageHint: "会自动读取当前页全部内容，输入框可补充额外美化要求。",
      pagePlaceholder: "可选：例如改成更商务、更简洁的版式",
      pageActionText: "开始美化单页",
      beautify: true
    },
    "theme-color": {
      title: "一键换主题",
      subtitle: "选择颜色后，把当前演示文稿切换为对应色系",
      panel: "theme-color"
    },
    "template-picker": {
      title: "挑选模板",
      subtitle: "选择模板后生成 PPT",
      panel: "template-picker"
    },
  };

  var $ = function (selector) {
    return document.querySelector(selector);
  };

  function loadConfig() {
    return Object.assign({}, DOCMEE_CONFIG, global.WENDUODUO_DOCMEE_CONFIG || {});
  }

  function client() {
    return new global.DocmeeClient(loadConfig());
  }

  function docmeeLanguageCode(languageName) {
    var map = {
      "中文": "zh",
      "英语": "en",
      "日语": "ja",
      "韩语": "ko",
      "法语": "fr",
      "德语": "de",
      "西班牙语": "es"
    };
    return map[languageName] || "en";
  }

  function formatSinglePageResult(result) {
    var lines = ["AI 单页已生成。"];
    if (result.templateId) lines.push("使用模板 ID: " + result.templateId);
    if (result.pptxId) lines.push("生成 PPT ID: " + result.pptxId);
    if (result.fileUrl) lines.push("PPTX 下载地址: " + result.fileUrl);
    return lines.join("\n\n");
  }

  function setStatus(selector, message) {
    var el = $(selector);
    if (el) el.textContent = message || "";
  }

  function showResultEditor(selector) {
    var el = $(selector);
    if (el) el.classList.remove("is-hidden");
  }

  function setVisible(selector, visible) {
    var el = $(selector);
    if (el) el.classList.toggle("is-hidden", !visible);
  }

  function getSinglePagePanel() {
    return document.querySelector('[data-panel="single-page"]');
  }

  function setSinglePageStage(stage) {
    var panel = getSinglePagePanel();
    if (!panel) return;
    panel.classList.toggle("single-page-working", stage === "loading");
    panel.classList.toggle("single-page-has-choices", stage === "choices");
  }

  function setText(selector, text) {
    var el = $(selector);
    if (el) el.textContent = text;
  }

  function setPlaceholder(selector, text) {
    var el = $(selector);
    if (el) el.placeholder = text;
  }

  function setLabelPrefix(selector, text) {
    var el = $(selector);
    if (!el) return;
    if (el.firstChild && el.firstChild.nodeType === Node.TEXT_NODE) {
      el.firstChild.nodeValue = text;
    } else {
      el.insertBefore(document.createTextNode(text), el.firstChild);
    }
  }

  function setSelectOptions(selector, items) {
    var el = $(selector);
    if (!el) return;
    el.innerHTML = "";
    items.forEach(function (item) {
      var option = document.createElement("option");
      option.value = item.value;
      option.textContent = item.text;
      el.appendChild(option);
    });
  }

  function applyStaticChineseText() {
    document.title = "文多多 WPS 演示";
    setText('[data-panel="generate-ppt"] h2', "输入创作灵感");
    setPlaceholder("#topicInput", "输入您的创作灵感，例如：AI 办公趋势");
    setText("#createOutlineButton", "立即创作");
    setText("#editOutlineButton", "修改大纲");
    setText("#pickTemplateButton", "挑选模板");
    setPlaceholder("#markdownOutput", "生成后的 Markdown 大纲会展示在这里");

    setText('[for="documentPath"]', "文档");
    setPlaceholder("#documentPath", "请选择 Word、PDF、Markdown、PPT 等文档");
    setText("#selectDocumentButton", "选择文档");
    setText("#createDocumentOutlineButton", "立即创作");
    setText("#editDocumentOutlineButton", "修改大纲");
    setText("#pickDocumentTemplateButton", "挑选模板");
    setPlaceholder("#documentMarkdownOutput", "文档解析后的 Markdown 大纲会展示在这里");

    setText('[data-panel="template-picker"] h2', "选择模板");
    setText("#backToOutlineButton", "上一步");
    setLabelPrefix('[for="styleFilter"], .filter-row label:nth-child(1)', "风格");
    setLabelPrefix('[for="categoryFilter"], .filter-row label:nth-child(2)', "类别");
    setSelectOptions("#styleFilter", [
      { value: "", text: "全部" },
      { value: "扁平简约", text: "扁平简约" },
      { value: "商务科技", text: "商务科技" },
      { value: "创意趣味", text: "创意趣味" }
    ]);
    setSelectOptions("#categoryFilter", [
      { value: "", text: "全部" },
      { value: "办公报告", text: "办公报告" },
      { value: "个人简历", text: "个人简历" },
      { value: "教育培训", text: "教育培训" },
      { value: "现代商务", text: "现代商务" }
    ]);
    var themeFilter = $("#themeColorFilter");
    if (themeFilter) themeFilter.title = "主题色";
    setText("#refreshTemplatesButton", "刷新");
    setText("#templatePrevPageButton", "上一页");
    setText("#templatePageLabel", "第 1 页");
    setText("#templateNextPageButton", "下一页");
    setText("#generatePptButton", "下一步");

    setLabelPrefix("#languageRow", "目标语言");
    setSelectOptions("#targetLanguage", [
      { value: "英语", text: "英语" },
      { value: "中文", text: "中文" },
      { value: "日语", text: "日语" },
      { value: "韩语", text: "韩语" },
      { value: "法语", text: "法语" },
      { value: "德语", text: "德语" },
      { value: "西班牙语", text: "西班牙语" }
    ]);
    setText("#runTextCreateButton", "开始润色");
    setText("#replaceSelectedTextButton", "替换回选中内容");
    setPlaceholder("#textSourceInput", "选中文字后打开，或直接输入要处理的文本");
    setPlaceholder("#textResultOutput", "AI 创作结果");

    setText(".frame-label", "新页面");
    setPlaceholder("#singlePageInput", "请输入页面标题或内容");
    setText("#singlePageActionButton", "AI智能生成新页面");
    setPlaceholder("#singlePageResult", "生成过程和结果");

    setText('[data-panel="theme-color"] h2', "一键换主题");
    setText('[data-panel="theme-color"] .hint', "选择一个颜色，当前 WPS 演示会切换为对应色系。");
    setText("#applyThemeButton", "应用主题");
  }

  function updateSinglePageCount() {
    var input = $("#singlePageInput");
    var count = $("#singlePageCount");
    if (input && count) count.textContent = input.value.length + " / 1000";
  }

  function showMode(mode) {
    var meta = MODE_META[mode] || MODE_META["generate-ppt"];
    state.mode = mode;
    if (mode !== "beautify-page") state.autoBeautifyStarted = false;
    $("#appTitle").textContent = meta.title;
    $("#appSubtitle").textContent = meta.subtitle;
    configureTextMode(meta);
    configureSinglePageMode(meta);
    document.querySelectorAll(".panel").forEach(function (panel) {
      panel.classList.toggle("active", panel.dataset.panel === meta.panel);
    });
  }

  function configureTextMode(meta) {
    if (meta.panel !== "text-create") return;
    state.textType = meta.textType;
    $("#textModeTitle").textContent = meta.textTitle;
    $("#textModeHint").textContent = meta.textHint;
    $("#textSourceInput").placeholder = meta.placeholder;
    $("#runTextCreateButton").textContent = meta.actionText;
    $("#languageRow").classList.toggle("visible", meta.textType === "fy");
  }

  function configureSinglePageMode(meta) {
    if (meta.panel !== "single-page") return;
    $("#singlePageTitle").textContent = meta.pageTitle;
    $("#singlePageHint").textContent = meta.pageHint;
    $("#singlePageInput").placeholder = meta.pagePlaceholder;
    $("#singlePageActionButton").textContent = meta.pageActionText;
    $("#singlePageLoading").setAttribute("data-label", meta.beautify ? "美化单页" : "新页面");
    $("#singlePageChooser").setAttribute("data-label", "新页面");
    updateSinglePageCount();
    resetSinglePageGeneratedView();
    if (meta.beautify) queueAutoBeautifySinglePage();
  }

  function showError(error, target) {
    var message = error && error.message ? error.message : String(error);
    if (target) {
      setStatus(target, message);
    } else {
      alert(message);
    }
  }

  function setButtonBusy(button, busyText, promiseFactory) {
    var oldText = button.textContent;
    button.disabled = true;
    button.textContent = busyText;
    return Promise.resolve()
      .then(promiseFactory)
      .finally(function () {
        button.textContent = oldText;
        button.disabled = false;
        refreshEnabledState();
      });
  }

  function refreshEnabledState() {
    var topic = $("#topicInput").value.trim();
    var hasMarkdown = !!state.markdown;
    var hasTopicOutline = hasMarkdown && state.outlineSourceMode === "generate-ppt";
    var hasDocumentOutline = hasMarkdown && state.outlineSourceMode === "document-generate";
    var showTopicResult = state.isGeneratingOutline || hasTopicOutline;
    var showDocumentResult = state.isGeneratingDocumentOutline || hasDocumentOutline;
    var generatePanel = document.querySelector('[data-panel="generate-ppt"]');
    var documentPanel = document.querySelector('[data-panel="document-generate"]');
    if (generatePanel) generatePanel.classList.toggle("outline-ready", showTopicResult);
    if (documentPanel) documentPanel.classList.toggle("outline-ready", showDocumentResult);
    $("#createOutlineButton").disabled = !topic || state.isGeneratingOutline;
    $("#createDocumentOutlineButton").disabled = !state.selectedFile || state.isGeneratingDocumentOutline;
    $("#editOutlineButton").disabled = !hasTopicOutline || state.isGeneratingOutline;
    $("#pickTemplateButton").disabled = !hasTopicOutline || state.isGeneratingOutline;
    $("#editDocumentOutlineButton").disabled = !hasDocumentOutline || state.isGeneratingDocumentOutline;
    $("#pickDocumentTemplateButton").disabled = !hasDocumentOutline || state.isGeneratingDocumentOutline;
    setVisible("#editOutlineButton", hasTopicOutline);
    setVisible("#pickTemplateButton", hasTopicOutline);
    setVisible("#editDocumentOutlineButton", hasDocumentOutline);
    setVisible("#pickDocumentTemplateButton", hasDocumentOutline);
    $("#generatePptButton").disabled = !state.templateId || !state.markdown || !state.taskId;
    $("#replaceSelectedTextButton").disabled = !state.lastTextResult;
    var singleMeta = MODE_META[state.mode] || {};
    if (singleMeta.panel === "single-page") {
      $("#singlePageActionButton").disabled = state.isGeneratingSinglePage || state.isApplyingSinglePage || (!singleMeta.beautify && !$("#singlePageInput").value.trim());
      $("#applySinglePageButton").disabled = state.isGeneratingSinglePage || state.isApplyingSinglePage || !state.singlePageChoices.length;
    }
  }

  function appendGeneratedMarkdown(markdown, selector) {
    state.markdown = markdown || "";
    $(selector).value = state.markdown;
    refreshEnabledState();
  }

  async function createOutlineFromTopic() {
    var topic = $("#topicInput").value.trim();
    if (!topic) {
      setStatus("#outlineStatus", "请输入您的创作灵感。");
      return;
    }
    state.outlineSourceMode = "generate-ppt";
    state.isGeneratingOutline = true;
    state.markdown = "";
    state.templateId = "";
    setStatus("#outlineStatus", "");
    showResultEditor("#markdownOutput");
    $("#markdownOutput").value = "正在生成大纲...\n";
    $("#markdownOutput").scrollTop = $("#markdownOutput").scrollHeight;
    refreshEnabledState();

    try {
      state.taskId = await client().createTask(topic, "1");
      var markdown = await client().generateMarkdown(state.taskId, function (partial) {
        $("#markdownOutput").value = partial || "正在生成大纲...\n";
        $("#markdownOutput").scrollTop = $("#markdownOutput").scrollHeight;
      });
      state.isGeneratingOutline = false;
      appendGeneratedMarkdown(markdown, "#markdownOutput");
      $("#markdownOutput").scrollTop = $("#markdownOutput").scrollHeight;
      setStatus("#outlineStatus", "");
    } catch (error) {
      state.isGeneratingOutline = false;
      refreshEnabledState();
      throw error;
    }
  }

  async function createOutlineFromDocument() {
    if (!state.selectedFile) {
      setStatus("#documentStatus", "请先选择文档。");
      return;
    }
    state.outlineSourceMode = "document-generate";
    state.isGeneratingDocumentOutline = true;
    state.markdown = "";
    state.templateId = "";
    setStatus("#documentStatus", "");
    showResultEditor("#documentMarkdownOutput");
    $("#documentMarkdownOutput").value = "正在解析文档...\n";
    $("#documentMarkdownOutput").scrollTop = $("#documentMarkdownOutput").scrollHeight;
    refreshEnabledState();

    try {
      state.taskId = await client().createFileTask(state.selectedFile);
      var markdown = await client().generateMarkdown(state.taskId, function (partial) {
        $("#documentMarkdownOutput").value = partial || "正在解析文档...\n";
        $("#documentMarkdownOutput").scrollTop = $("#documentMarkdownOutput").scrollHeight;
      });
      state.isGeneratingDocumentOutline = false;
      appendGeneratedMarkdown(markdown, "#documentMarkdownOutput");
      $("#documentMarkdownOutput").scrollTop = $("#documentMarkdownOutput").scrollHeight;
      setStatus("#documentStatus", "");
    } catch (error) {
      state.isGeneratingDocumentOutline = false;
      refreshEnabledState();
      throw error;
    }
  }

  function activeMarkdownEditor() {
    if (state.outlineSourceMode === "document-generate") return $("#documentMarkdownOutput");
    return $("#markdownOutput");
  }

  async function loadTemplates() {
    var grid = $("#templateGrid");
    grid.innerHTML = "<div class=\"status-line\">正在加载模板...</div>";
    $("#templatePageLabel").textContent = "第 " + state.templatePage + " 页";
    var templates;
    try {
      templates = await client().listTemplates({
        page: state.templatePage,
        size: 10,
        style: $("#styleFilter").value,
        category: $("#categoryFilter").value,
        themeColor: state.templateThemeColor
      });
    } catch (error) {
      templates = global.WPS_AI_FALLBACK_TEMPLATES || [];
      grid.innerHTML = "<div class=\"status-line\">模板接口暂不可用，已展示本地兜底模板。</div>";
    }
    if ((!templates || !templates.length) && state.templateThemeColor) {
      state.templateThemeColor = "";
      $("#themeColorFilter").classList.remove("active");
      templates = await client().listTemplates({
        page: state.templatePage,
        size: 10,
        style: $("#styleFilter").value,
        category: $("#categoryFilter").value,
        themeColor: ""
      });
    }
    if ((!templates || !templates.length) && global.WPS_AI_FALLBACK_TEMPLATES) {
      templates = global.WPS_AI_FALLBACK_TEMPLATES;
    }
    renderTemplates(templates);
  }

  function renderTemplates(templates) {
    var grid = $("#templateGrid");
    if (!templates || !templates.length) {
      grid.innerHTML = "<div class=\"status-line\">暂无模板。</div>";
      return;
    }
    grid.innerHTML = "";
    templates.forEach(function (template) {
      var card = document.createElement("article");
      var coverUrl = template.coverUrl || templateCoverDataUrl(template);
      card.className = "template-card" + (state.templateId === template.id ? " selected" : "");
      card.innerHTML =
        "<img alt=\"模板封面\" src=\"" + escapeHtml(coverUrl) + "\">" +
        "<h3 title=\"" + escapeHtml(template.name || template.id) + "\">" + escapeHtml(template.name || template.id) + "</h3>" +
        "<p>" + escapeHtml([template.category, template.style].filter(Boolean).join(" / ")) + " | ID " + escapeHtml(template.id) + "</p>" +
        "<button class=\"" + (state.templateId === template.id ? "primary-button" : "secondary-button") + "\" type=\"button\">" + (state.templateId === template.id ? "已选择" : "选择模板") + "</button>";
      card.querySelector("img").addEventListener("error", function () {
        this.onerror = null;
        this.src = templateCoverDataUrl(template);
      });
      card.querySelector("button").addEventListener("click", function () {
        state.templateId = template.id;
        renderTemplates(templates);
        refreshEnabledState();
      });
      grid.appendChild(card);
    });
  }

  function templateCoverDataUrl(template) {
    var palette = ["#4f46e5", "#f97316", "#0891b2", "#16a34a", "#db2777"];
    var color = palette[Math.abs(hashCode(template.id || template.name || "0")) % palette.length];
    var name = escapeSvg(clipText(template.name || "文多多模板", 18));
    var category = escapeSvg([template.category, template.style].filter(Boolean).join(" / "));
    var svg =
      "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"960\" height=\"540\" viewBox=\"0 0 960 540\">" +
      "<rect width=\"960\" height=\"540\" fill=\"#f8fafc\"/>" +
      "<rect x=\"0\" y=\"0\" width=\"960\" height=\"96\" fill=\"" + color + "\"/>" +
      "<circle cx=\"820\" cy=\"80\" r=\"150\" fill=\"" + color + "\" opacity=\"0.16\"/>" +
      "<circle cx=\"900\" cy=\"460\" r=\"210\" fill=\"" + color + "\" opacity=\"0.1\"/>" +
      "<rect x=\"72\" y=\"150\" width=\"560\" height=\"20\" rx=\"10\" fill=\"" + color + "\" opacity=\"0.16\"/>" +
      "<text x=\"72\" y=\"250\" font-family=\"Microsoft YaHei, Arial\" font-size=\"54\" font-weight=\"700\" fill=\"#0f172a\">" + name + "</text>" +
      "<text x=\"72\" y=\"322\" font-family=\"Microsoft YaHei, Arial\" font-size=\"28\" fill=\"#64748b\">" + category + "</text>" +
      "<rect x=\"72\" y=\"400\" width=\"160\" height=\"36\" rx=\"18\" fill=\"" + color + "\" opacity=\"0.9\"/>" +
      "<text x=\"104\" y=\"425\" font-family=\"Microsoft YaHei, Arial\" font-size=\"18\" fill=\"#ffffff\">文多多</text>" +
      "</svg>";
    return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  }

  function hashCode(value) {
    var hash = 0;
    for (var i = 0; i < String(value).length; i += 1) {
      hash = ((hash << 5) - hash) + String(value).charCodeAt(i);
      hash |= 0;
    }
    return hash;
  }

  function clipText(text, maxLength) {
    text = String(text || "");
    return text.length > maxLength ? text.slice(0, maxLength - 1) + "…" : text;
  }

  function escapeSvg(text) {
    return String(text || "").replace(/[&<>]/g, function (char) {
      return {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;"
      }[char];
    });
  }

  function escapeHtml(text) {
    return String(text || "").replace(/[&<>"']/g, function (char) {
      return {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        "\"": "&quot;",
        "'": "&#39;"
      }[char];
    });
  }

  async function generatePptx() {
    state.markdown = activeMarkdownEditor().value.trim() || state.markdown;
    if (!state.markdown) throw new Error("缺少 Markdown 大纲。");
    if (!state.templateId) throw new Error("请先选择模板。");
    $("#generatePptButton").textContent = "生成中...";
    state.generatedPptInfo = await client().generatePptx(state.taskId, state.templateId, state.markdown);
    state.generatedPptUrl = await client().downloadPptx(state.generatedPptInfo.id, false, 6, 1200);
    $("#generatePptButton").textContent = "导入中...";
    var importResult = await global.WpsPptAdapter.importPptxFromUrl(state.generatedPptUrl, state.generatedPptInfo.id);
    if (importResult.importedCount > 0) {
      alert("已生成并导入当前演示文稿，共 " + importResult.importedCount + " 页。");
      global.WpsPptAdapter.hideTaskPane();
    } else {
      alert("PPT 已生成，已打开下载地址。当前环境未提供 WPS 自动导入能力。");
    }
  }

  async function runTextCreation() {
    var source = $("#textSourceInput").value.trim();
    if (!source && global.WpsPptAdapter.isWpsHost()) {
      try {
        source = global.WpsPptAdapter.getSelectedText();
        $("#textSourceInput").value = source;
      } catch (error) {}
    }
    if (!source) {
      alert("请先选中或输入要处理的文本。");
      return;
    }
    var targetLanguageName = $("#targetLanguage").value;
    showResultEditor("#textResultOutput");
    $("#textResultOutput").value = "正在创作...\n";
    state.lastTextResult = await client().textOptimize({
      type: state.textType,
      text: source,
      targetLanguageName: targetLanguageName,
      targetLanguageCode: docmeeLanguageCode(targetLanguageName)
    }, function (partial) {
      $("#textResultOutput").value = partial;
      $("#textResultOutput").scrollTop = $("#textResultOutput").scrollHeight;
    });
    $("#textResultOutput").value = state.lastTextResult;
    refreshEnabledState();
  }

  function resetSinglePageGeneratedView() {
    state.singlePageChoices = [];
    state.singlePageSelectedIndex = 0;
    state.singlePageResultInfo = null;
    setSinglePageStage("");
    setVisible(".single-page-frame", true);
    setVisible("#singlePageLoading", false);
    setVisible("#singlePageChooser", false);
    setVisible("#singlePageResult", false);
    var preview = $("#singlePageMainPreview");
    if (preview) preview.removeAttribute("src");
    var shell = document.querySelector(".single-page-preview-shell");
    if (shell) shell.classList.remove("has-image");
    var thumbs = $("#singlePageThumbList");
    if (thumbs) thumbs.innerHTML = "";
  }

  function setSinglePageLoading(visible, message) {
    setSinglePageStage(visible ? "loading" : (state.singlePageChoices.length ? "choices" : ""));
    if (visible) {
      setVisible(".single-page-frame", false);
      setVisible("#singlePageChooser", false);
    } else if (!state.singlePageChoices.length) {
      setVisible(".single-page-frame", true);
    }
    setVisible("#singlePageLoading", visible);
    if (message) setText("#singlePageLoadingTitle", message);
  }

  function selectSinglePageChoice(index) {
    if (!state.singlePageChoices.length) return;
    state.singlePageSelectedIndex = Math.max(0, Math.min(index, state.singlePageChoices.length - 1));
    var choice = state.singlePageChoices[state.singlePageSelectedIndex];
    var preview = $("#singlePageMainPreview");
    var shell = document.querySelector(".single-page-preview-shell");
    if (preview) {
      if (choice.previewUrl) {
        preview.src = choice.previewUrl;
        if (shell) shell.classList.add("has-image");
      } else {
        preview.removeAttribute("src");
        if (shell) shell.classList.remove("has-image");
      }
    }
    document.querySelectorAll(".single-page-thumb").forEach(function (thumb, thumbIndex) {
      thumb.classList.toggle("active", thumbIndex === state.singlePageSelectedIndex);
    });
  }

  function renderSinglePageChoices(result, choiceResult) {
    state.singlePageResultInfo = result;
    state.singlePageChoices = (choiceResult && choiceResult.choices ? choiceResult.choices : []).filter(function (choice) {
      return choice && choice.localPath && choice.slideIndex;
    });
    state.singlePageSelectedIndex = 0;

    var list = $("#singlePageThumbList");
    list.innerHTML = "";
    if (!state.singlePageChoices.length) {
      setSinglePageStage("");
      setVisible(".single-page-frame", true);
      setStatus("#singlePageStatus", "已生成 PPTX，但没有读到可选择的候选页。");
      return;
    }

    state.singlePageChoices.forEach(function (choice, index) {
      var card = document.createElement("button");
      card.type = "button";
      card.className = "single-page-thumb";
      var image = document.createElement("img");
      if (choice.previewUrl) image.src = choice.previewUrl;
      image.alt = choice.title || ("第 " + choice.slideIndex + " 页");
      var title = document.createElement("div");
      title.className = "single-page-thumb-title";
      title.textContent = "第 " + choice.slideIndex + " 页 - " + (choice.title || "候选页面");
      card.appendChild(image);
      card.appendChild(title);
      card.addEventListener("click", function () {
        selectSinglePageChoice(index);
      });
      list.appendChild(card);
    });

    setSinglePageStage("choices");
    setVisible(".single-page-frame", false);
    setVisible("#singlePageChooser", true);
    selectSinglePageChoice(0);
  }

  async function applySinglePageChoice() {
    if (!state.singlePageChoices.length) {
      setStatus("#singlePageStatus", "请先生成并选择一个候选页面。");
      return;
    }
    var choice = state.singlePageChoices[state.singlePageSelectedIndex];
    state.isApplyingSinglePage = true;
    $("#applySinglePageButton").disabled = true;
    $("#applySinglePageButton").textContent = "正在应用...";
    setStatus("#singlePageStatus", "正在替换当前页...");
    try {
      await global.WpsPptAdapter.applySinglePageChoice(choice);
      setStatus("#singlePageStatus", "已应用到当前页。");
      global.WpsPptAdapter.closeHostSurface();
    } finally {
      state.isApplyingSinglePage = false;
      $("#applySinglePageButton").disabled = false;
      $("#applySinglePageButton").textContent = "将此页面插入演示文档";
      refreshEnabledState();
    }
  }

  function closeSinglePageExperience() {
    resetSinglePageGeneratedView();
    try {
      global.WpsPptAdapter.closeHostSurface();
    } catch (error) {}
  }

  async function runSinglePage(content, beautify) {
    if (!content.trim()) {
      setStatus("#singlePageStatus", "请输入页面标题或内容。");
      return;
    }
    var label = beautify ? "正在美化当前页..." : "正在生成新页面...";
    setStatus("#singlePageStatus", label);
    showResultEditor("#singlePageResult");
    $("#singlePageResult").value = label + "\n";
    var progressLines = [];
    var result = await client().newPageWithAiV2(content, function (partial) {
      progressLines.push(partial);
      $("#singlePageResult").value = progressLines.slice(-20).join("\n");
      $("#singlePageResult").scrollTop = $("#singlePageResult").scrollHeight;
    });
    $("#singlePageResult").value = formatSinglePageResult(result);
    if (result.fileUrl) global.WpsPptAdapter.openPptxDownload(result.fileUrl);
    setStatus("#singlePageStatus", "生成完成，已打开下载地址。WPS 自动应用当前页能力将在宿主导入接口接通后启用。");
  }

  async function runSinglePageV2(content, beautify) {
    if (!content.trim()) {
      setStatus("#singlePageStatus", "请输入页面标题或内容。");
      return;
    }

    resetSinglePageGeneratedView();
    state.isGeneratingSinglePage = true;
    var label = beautify ? "正在生成美化页面..." : "正在生成候选页面...";
    setStatus("#singlePageStatus", label);
    setSinglePageLoading(true, label);

    try {
      var result = await client().newPageWithAiV2(content, function (partial) {
        if (partial) {
          setStatus("#singlePageStatus", beautify ? "正在美化当前页..." : "正在生成新页面...");
          setSinglePageLoading(true, label);
        }
      });

      if (!result.fileUrl) throw new Error("Docmee 没有返回可下载的 PPTX。");
      setStatus("#singlePageStatus", "正在准备候选页面预览...");
      setSinglePageLoading(true, "正在准备候选页面...");
      var choiceResult = await global.WpsPptAdapter.createSinglePageChoicesFromUrl(result.fileUrl, result.pptxId);
      setVisible("#singlePageLoading", false);
      renderSinglePageChoices(result, choiceResult);
      setStatus("#singlePageStatus", "请选择一个候选页面应用到当前页。");
    } finally {
      state.isGeneratingSinglePage = false;
      setSinglePageLoading(false);
      refreshEnabledState();
    }
  }

  function buildBeautifySinglePageContent() {
    var extraRequirement = String($("#singlePageInput").value || "").trim();
    var currentPageContent = "";
    try {
      currentPageContent = global.WpsPptAdapter.getCurrentSlideText() || "";
    } catch (error) {}
    currentPageContent = String(currentPageContent || "").trim();
    if (!currentPageContent) throw new Error("当前页没有可用于美化的内容。");

    var content = "请基于当前页全部内容进行美化，保留原有信息，不要丢失要点。\n\n当前页内容：\n" + currentPageContent;
    if (extraRequirement) content += "\n\n用户美化要求：\n" + extraRequirement;
    return content;
  }

  function runCurrentSinglePageMode() {
    var meta = MODE_META[state.mode] || MODE_META["generate-page"];
    var content = meta.beautify ? buildBeautifySinglePageContent() : $("#singlePageInput").value;
    return runSinglePageV2(content, !!meta.beautify);
  }

  function queueAutoBeautifySinglePage() {
    if (state.autoBeautifyStarted || state.isGeneratingSinglePage || state.singlePageChoices.length) return;
    state.autoBeautifyStarted = true;
    global.setTimeout(function () {
      var button = $("#singlePageActionButton");
      if (!button || state.mode !== "beautify-page") return;
      setButtonBusy(button, "美化中", runCurrentSinglePageMode).catch(function (error) {
        state.autoBeautifyStarted = false;
        showError(error, "#singlePageStatus");
      });
    }, 60);
  }

  function renderThemeSwatches() {
    var colors = ["#4f46e5", "#2563eb", "#0891b2", "#16a34a", "#f97316", "#ef4444", "#db2777", "#475569"];
    var wrap = $("#themeSwatches");
    wrap.innerHTML = "";
    colors.forEach(function (color) {
      var button = document.createElement("button");
      button.className = "swatch" + (color === $("#themeColorInput").value ? " active" : "");
      button.type = "button";
      button.style.background = color;
      button.title = color.toUpperCase();
      button.addEventListener("click", function () {
        $("#themeColorInput").value = color;
        $("#themeColorLabel").textContent = color.toUpperCase();
        renderThemeSwatches();
      });
      wrap.appendChild(button);
    });
  }

  function initEvents() {
    $("#topicInput").addEventListener("input", refreshEnabledState);
    $("#singlePageInput").addEventListener("input", function () {
      updateSinglePageCount();
      refreshEnabledState();
    });
    $("#createOutlineButton").addEventListener("click", function () {
      setButtonBusy($("#createOutlineButton"), "正在生成", createOutlineFromTopic).catch(function (error) {
        showError(error, "#outlineStatus");
      });
    });
    $("#editOutlineButton").addEventListener("click", function () {
      activeMarkdownEditor().focus();
    });
    $("#pickTemplateButton").addEventListener("click", function () {
      state.outlineSourceMode = "generate-ppt";
      state.markdown = $("#markdownOutput").value.trim();
      showMode("template-picker");
      loadTemplates();
    });
    $("#pickDocumentTemplateButton").addEventListener("click", function () {
      state.outlineSourceMode = "document-generate";
      state.markdown = $("#documentMarkdownOutput").value.trim();
      showMode("template-picker");
      loadTemplates();
    });

    $("#selectDocumentButton").addEventListener("click", function () {
      $("#documentFileInput").click();
    });
    $("#documentFileInput").addEventListener("change", function (event) {
      state.selectedFile = event.target.files && event.target.files[0] ? event.target.files[0] : null;
      $("#documentPath").value = state.selectedFile ? state.selectedFile.name : "";
      refreshEnabledState();
    });
    $("#createDocumentOutlineButton").addEventListener("click", function () {
      setButtonBusy($("#createDocumentOutlineButton"), "正在生成", createOutlineFromDocument).catch(function (error) {
        showError(error, "#documentStatus");
      });
    });
    $("#editDocumentOutlineButton").addEventListener("click", function () {
      $("#documentMarkdownOutput").focus();
    });

    $("#backToOutlineButton").addEventListener("click", function () {
      showMode(state.selectedFile ? "document-generate" : "generate-ppt");
    });
    $("#refreshTemplatesButton").addEventListener("click", loadTemplates);
    $("#styleFilter").addEventListener("change", function () {
      state.templatePage = 1;
      loadTemplates();
    });
    $("#categoryFilter").addEventListener("change", function () {
      state.templatePage = 1;
      loadTemplates();
    });
    $("#themeColorFilter").addEventListener("input", function () {
      state.templateThemeColor = $("#themeColorFilter").value;
      $("#themeColorFilter").classList.add("active");
      state.templatePage = 1;
      loadTemplates();
    });
    $("#templatePrevPageButton").addEventListener("click", function () {
      state.templatePage = Math.max(1, state.templatePage - 1);
      loadTemplates();
    });
    $("#templateNextPageButton").addEventListener("click", function () {
      state.templatePage += 1;
      loadTemplates();
    });
    $("#generatePptButton").addEventListener("click", function () {
      setButtonBusy($("#generatePptButton"), "生成中", generatePptx).catch(showError);
    });

    $("#runTextCreateButton").addEventListener("click", function () {
      setButtonBusy($("#runTextCreateButton"), "创作中", runTextCreation).catch(showError);
    });
    $("#replaceSelectedTextButton").addEventListener("click", function () {
      try {
        global.WpsPptAdapter.replaceSelectedText(state.lastTextResult);
      } catch (error) {
        showError(error);
      }
    });

    $("#singlePageActionButton").addEventListener("click", function () {
      var meta = MODE_META[state.mode] || MODE_META["generate-page"];
      var busyText = meta.beautify ? "美化中" : "生成中";
      setButtonBusy($("#singlePageActionButton"), busyText, runCurrentSinglePageMode).catch(function (error) {
        showError(error, "#singlePageStatus");
      });
    });

    $("#applySinglePageButton").addEventListener("click", function () {
      applySinglePageChoice().catch(function (error) {
        showError(error, "#singlePageStatus");
      });
    });
    $("#singlePageLoadingCloseButton").addEventListener("click", closeSinglePageExperience);
    $("#singlePageChooserCloseButton").addEventListener("click", closeSinglePageExperience);

    $("#themeColorInput").addEventListener("input", function () {
      $("#themeColorLabel").textContent = $("#themeColorInput").value.toUpperCase();
      renderThemeSwatches();
    });
    $("#applyThemeButton").addEventListener("click", function () {
      try {
        var changed = global.WpsPptAdapter.applyThemeColor($("#themeColorInput").value);
        alert("已应用主题色，处理对象：" + changed + " 个。");
      } catch (error) {
        showError(error);
      }
    });

  }

  function initFromHash() {
    var mode = decodeURIComponent((global.location.hash || "").replace(/^#\/?/, ""));
    var map = {
      "generate-ppt": "generate-ppt",
      "document-generate": "document-generate",
      "text-polish": "text-polish",
      "text-expand": "text-expand",
      "text-shorten": "text-shorten",
      "text-translate": "text-translate",
      "beautify-page": "beautify-page",
      "generate-page": "generate-page",
      "theme-color": "theme-color",
      "template-picker": "template-picker"
    };
    var mapped = map[mode] || "generate-ppt";
    showMode(mapped);
    if (mapped === "template-picker") loadTemplates();
    refreshEnabledState();
  }

  document.addEventListener("DOMContentLoaded", function () {
    applyStaticChineseText();
    initEvents();
    initFromHash();
    renderThemeSwatches();
    refreshEnabledState();
  });

  global.addEventListener("hashchange", initFromHash);
})(window);
