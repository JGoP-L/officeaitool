var WENDUODUO_TASKPANE_ID = "wenduoduo_wps_ppt_taskpane_id";
var WENDUODUO_ASSET_VERSION = "20260615-wps-preview-btoa-fallback";
var WENDUODUO_DIRECT_TEXT_BUSY = false;
var WENDUODUO_SCRIPT_LOAD_PROMISE = null;

var WENDUODUO_DOCMEE_CONFIG = {
  baseUrl: "https://test.docmee.cn",
  token: "ak_demo",
  companyId: "203361",
  lang: "zh",
  pptxId: ""
};

var WENDUODUO_DIRECT_TEXT_ACTIONS = {
  btnPolish: { type: "rs", label: "润色" },
  btnExpand: { type: "kx", label: "扩写" },
  btnShorten: { type: "sx", label: "缩写" }
};

function OnAddinLoad(ribbonUI) {
  if (typeof window.Application.ribbonUI !== "object") {
    window.Application.ribbonUI = ribbonUI;
  }
  return true;
}

function OnAction(control) {
  if (control.Id === "btnTranslate") {
    HideWenduoduoTaskPane();
    ShowTranslationLanguageDialog();
    return true;
  }

  var textAction = WENDUODUO_DIRECT_TEXT_ACTIONS[control.Id];
  if (textAction) {
    HideWenduoduoTaskPane();
    RunDirectTextAction(textAction);
    return true;
  }

  var modeMap = {
    btnGeneratePpt: "generate-ppt",
    btnDocumentGeneratePpt: "document-generate",
    btnPolish: "text-polish",
    btnExpand: "text-expand",
    btnShorten: "text-shorten",
    btnTranslate: "text-translate",
    btnBeautifyPage: "beautify-page",
    btnGeneratePage: "generate-page",
    btnThemeColor: "theme-color"
  };

  var mode = modeMap[control.Id] || "generate-ppt";
  if (control.Id === "btnBeautifyPage" || control.Id === "btnGeneratePage") {
    HideWenduoduoTaskPane();
    ShowSinglePageDialog(mode);
    return true;
  }
  ShowWenduoduoTaskPane(mode);
  return true;
}

function OnGetEnabled() {
  return true;
}

function OnGetVisible() {
  return true;
}

function GetImage(control) {
  var imageMap = {
    btnGeneratePpt: "images/generate-ppt.svg",
    btnDocumentGeneratePpt: "images/document-ppt.svg",
    btnPolish: "images/polish.svg",
    btnExpand: "images/expand.svg",
    btnShorten: "images/shorten.svg",
    btnTranslate: "images/translate.svg",
    btnBeautifyPage: "images/beautify.svg",
    btnGeneratePage: "images/page.svg",
    btnThemeColor: "images/theme.svg"
  };
  return imageMap[control.Id] || "images/generate-ppt.svg";
}

function ShowWenduoduoTaskPane(mode) {
  var url = GetUrlPath() + "/ui/taskpane.html?v=" + encodeURIComponent(WENDUODUO_ASSET_VERSION) + "#/" + encodeURIComponent(mode || "generate-ppt");
  var app = window.Application;

  try {
    var paneId = app.PluginStorage.getItem(WENDUODUO_TASKPANE_ID);
    var pane = null;

    if (paneId) {
      try {
        pane = app.GetTaskPane(paneId);
      } catch (error) {
        pane = null;
      }
    }

    if (!pane) {
      pane = app.CreateTaskPane(url);
      app.PluginStorage.setItem(WENDUODUO_TASKPANE_ID, pane.ID);
    } else if (typeof pane.Navigate === "function") {
      pane.Navigate(url);
    }

    pane.Visible = true;
  } catch (error) {
    alert("打开文多多任务窗格失败：" + (error && error.message ? error.message : error));
  }
}

function HideWenduoduoTaskPane() {
  var app = window.Application;
  if (!app || !app.PluginStorage || !app.GetTaskPane) return;
  try {
    var paneId = app.PluginStorage.getItem(WENDUODUO_TASKPANE_ID);
    if (!paneId) return;
    var pane = app.GetTaskPane(paneId);
    if (pane) pane.Visible = false;
  } catch (error) {}
}

function ShowSinglePageDialog(mode) {
  var title = mode === "beautify-page" ? "美化单页" : "AI生成单页";
  var url = GetUrlPath() + "/ui/taskpane.html?v=" + encodeURIComponent(WENDUODUO_ASSET_VERSION) + "#/" + encodeURIComponent(mode);

  try {
    if (window.wps && typeof window.wps.ShowDialog === "function") {
      window.wps.ShowDialog(url, title, 1180, 760, true);
      return;
    }
    if (window.Application && typeof window.Application.ShowDialog === "function") {
      window.Application.ShowDialog(url, title, 1180, 760, true);
      return;
    }

    ShowWenduoduoTaskPane(mode);
  } catch (error) {
    alert("打开" + title + "失败：" + (error && error.message ? error.message : error));
  }
}

function ShowTranslationLanguageDialog() {
  var url = GetUrlPath() + "/ui/translate-dialog.html?v=" + encodeURIComponent(WENDUODUO_ASSET_VERSION);
  try {
    if (window.wps && typeof window.wps.ShowDialog === "function") {
      window.wps.ShowDialog(url, "选择翻译语言", 520, 340, true);
      return;
    }
    if (window.Application && typeof window.Application.ShowDialog === "function") {
      window.Application.ShowDialog(url, "选择翻译语言", 520, 340, true);
      return;
    }
    alert("当前 WPS 环境不支持语言选择弹窗。");
  } catch (error) {
    alert("打开翻译语言选择框失败：" + (error && error.message ? error.message : error));
  }
}

function LoadScriptOnce(path, globalName) {
  if (window[globalName]) return Promise.resolve();
  return new Promise(function (resolve, reject) {
    var script = document.createElement("script");
    script.charset = "utf-8";
    script.src = GetUrlPath() + "/" + path + "?v=" + encodeURIComponent(WENDUODUO_ASSET_VERSION);
    script.onload = resolve;
    script.onerror = function () {
      reject(new Error("加载脚本失败：" + path));
    };
    document.head.appendChild(script);
  });
}

function EnsureDirectTextDependencies() {
  if (WENDUODUO_SCRIPT_LOAD_PROMISE) return WENDUODUO_SCRIPT_LOAD_PROMISE;
  WENDUODUO_SCRIPT_LOAD_PROMISE = LoadScriptOnce("src/docmee-client.js", "DocmeeClient")
    .then(function () {
      return LoadScriptOnce("src/wps-ppt-adapter.js", "WpsPptAdapter");
    });
  return WENDUODUO_SCRIPT_LOAD_PROMISE;
}

function SetWpsStatus(message) {
  try {
    var app = window.wps && typeof window.wps.WppApplication === "function"
      ? window.wps.WppApplication()
      : window.Application;
    if (app) app.StatusBar = message || false;
  } catch (error) {}
}

async function RunDirectTextAction(action) {
  if (WENDUODUO_DIRECT_TEXT_BUSY) {
    alert("正在处理选中文本，请稍候。");
    return;
  }

  WENDUODUO_DIRECT_TEXT_BUSY = true;
  SetWpsStatus("文多多正在" + action.label + "选中文本...");

  try {
    await EnsureDirectTextDependencies();
    var source = window.WpsPptAdapter.getSelectedText();
    if (!source || !String(source).trim()) {
      alert("请先在 WPS 演示中选中要" + action.label + "的文字或文本框。");
      return;
    }

    var options = {
      type: action.type,
      text: source
    };

    var client = new window.DocmeeClient(WENDUODUO_DOCMEE_CONFIG);
    var result = await client.textOptimize(options);
    if (!result || !String(result).trim()) throw new Error("接口没有返回可用内容。");
    window.WpsPptAdapter.replaceSelectedText(String(result).trim());
    SetWpsStatus("文多多" + action.label + "完成");
  } catch (error) {
    alert(action.label + "出错：" + (error && error.message ? error.message : error));
  } finally {
    WENDUODUO_DIRECT_TEXT_BUSY = false;
    window.setTimeout(function () {
      SetWpsStatus("");
    }, 1200);
  }
}

function GetUrlPath() {
  var href = String(document.location.href || "");
  var hashIndex = href.indexOf("#");
  if (hashIndex >= 0) href = href.slice(0, hashIndex);
  var queryIndex = href.indexOf("?");
  if (queryIndex >= 0) href = href.slice(0, queryIndex);
  return href.substring(0, href.lastIndexOf("/"));
}
