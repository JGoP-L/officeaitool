(function (global) {
  "use strict";

  var DOCMEE_CONFIG = {
    baseUrl: "https://test.docmee.cn",
    token: "ak_demo",
    companyId: "203361",
    lang: "zh",
    pptxId: ""
  };

  var LANGUAGES = [
    { name: "英语", code: "en" },
    { name: "中文", code: "zh" },
    { name: "日语", code: "ja" },
    { name: "韩语", code: "ko" },
    { name: "法语", code: "fr" },
    { name: "德语", code: "de" },
    { name: "西班牙语", code: "es" }
  ];

  var busy = false;

  function $(selector) {
    return document.querySelector(selector);
  }

  function setStatus(message) {
    $("#status").textContent = message || "";
  }

  function closeDialog() {
    try {
      global.close();
    } catch (error) {}
  }

  function setBusy(value) {
    busy = value;
    $("#targetLanguage").disabled = value;
    $("#translateButton").disabled = value;
    $("#cancelButton").disabled = value;
  }

  function selectedLanguage() {
    var code = $("#targetLanguage").value;
    for (var i = 0; i < LANGUAGES.length; i += 1) {
      if (LANGUAGES[i].code === code) return LANGUAGES[i];
    }
    return LANGUAGES[0];
  }

  async function translateSelectedText() {
    if (busy) return;
    var language = selectedLanguage();
    setBusy(true);

    try {
      setStatus("正在读取选中文本...");
      var source = global.WpsPptAdapter.getSelectedText();
      if (!source || !String(source).trim()) {
        setStatus("请先在 WPS 演示中选中要翻译的文字或文本框。");
        setBusy(false);
        return;
      }

      setStatus("正在翻译为" + language.name + "...");
      var client = new global.DocmeeClient(Object.assign({}, DOCMEE_CONFIG, global.WENDUODUO_DOCMEE_CONFIG || {}));
      var result = await client.textOptimize({
        type: "fy",
        text: source,
        targetLanguageName: language.name,
        targetLanguageCode: language.code
      });

      if (!result || !String(result).trim()) throw new Error("接口没有返回可用内容。");
      global.WpsPptAdapter.replaceSelectedText(String(result).trim());
      setStatus("翻译完成，已替换选中内容。");
      global.setTimeout(closeDialog, 700);
    } catch (error) {
      setStatus("翻译出错：" + (error && error.message ? error.message : error));
      setBusy(false);
    }
  }

  function renderLanguageOptions() {
    var select = $("#targetLanguage");
    select.innerHTML = "";
    LANGUAGES.forEach(function (language) {
      var option = document.createElement("option");
      option.value = language.code;
      option.textContent = language.name;
      select.appendChild(option);
    });
    select.value = "en";
  }

  document.addEventListener("DOMContentLoaded", function () {
    renderLanguageOptions();
    $("#translateButton").addEventListener("click", translateSelectedText);
    $("#cancelButton").addEventListener("click", closeDialog);
  });
})(window);
