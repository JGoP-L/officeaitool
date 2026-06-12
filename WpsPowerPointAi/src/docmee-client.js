(function (global) {
  "use strict";

  var NEW_PAGE_TEMPLATE_IDS = [
    "1804885538940116992", "1804889500284084224", "1804893649646116864",
    "1804898801006403584", "1804901831135191040", "1804902366068334592",
    "1804904403862544384", "1804905770857521152", "1804906177663066112",
    "1805081814809960448", "1806268304982269952", "1806271593098502144",
    "1806279004165234688", "1806280891782389760", "1806285286083387392",
    "1806287200204349440", "1806290661188820992", "1806291734771261440",
    "1806297030256222208", "1806297660265848832", "1806299845762473984",
    "1806301544875024384", "1806303058985213952", "1806304212762746880",
    "1806506174552727552", "1807601742553276416", "1807617227806203904",
    "1807655586138152960", "1807658348435464192", "1807659875694796800"
  ];

  function normalizeBaseUrl(baseUrl) {
    return String(baseUrl || "https://test.docmee.cn").replace(/\/+$/, "");
  }

  function randomNewPageTemplateId() {
    return NEW_PAGE_TEMPLATE_IDS[Math.floor(Math.random() * NEW_PAGE_TEMPLATE_IDS.length)];
  }

  function getString(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "string") return value;
    if (typeof value === "number" || typeof value === "boolean") return String(value);
    return "";
  }

  function firstText(payload, keys) {
    if (!payload || typeof payload !== "object") return "";
    for (var i = 0; i < keys.length; i += 1) {
      var value = payload[keys[i]];
      if (typeof value === "string" && value) return value;
      if (value && typeof value === "object") {
        var nested = firstText(value, keys);
        if (nested) return nested;
      }
    }
    return "";
  }

  function pathText(payload, path) {
    if (!payload || !path) return "";
    var parts = path.split(".");
    var current = payload;
    for (var i = 0; i < parts.length; i += 1) {
      if (!current || typeof current !== "object") return "";
      current = current[parts[i]];
    }
    return getString(current);
  }

  function firstPathText(payload, paths) {
    for (var i = 0; i < paths.length; i += 1) {
      var value = pathText(payload, paths[i]);
      if (value) return value;
    }
    return "";
  }

  function ensureSuccessEnvelope(payload) {
    if (!payload || typeof payload !== "object") return;
    var code = payload.code;
    if (code === undefined || code === null || code === 0 || code === 200 || code === "0" || code === "200") return;
    var message = firstText(payload, ["message", "msg", "error", "errorMessage", "detail"]) || JSON.stringify(payload);
    throw new Error("Docmee 返回失败: " + message);
  }

  function tryParseJson(text) {
    try {
      return JSON.parse(text);
    } catch (error) {
      return null;
    }
  }

  function extractMarkdown(payload) {
    if (!payload) return "";
    var result = payload.result || (payload.data && payload.data.result);
    if (typeof result === "string") return result;
    var text = payload.text || payload.content || (payload.data && (payload.data.text || payload.data.content));
    return getString(text);
  }

  function extractTaskId(payload) {
    return getString(payload && payload.data && payload.data.id);
  }

  function extractPptInfo(payload) {
    var pptInfo = payload && payload.data && payload.data.pptInfo;
    if (!pptInfo) return null;
    return {
      id: getString(pptInfo.id),
      templateId: getString(pptInfo.templateId),
      subject: getString(pptInfo.subject),
      coverUrl: getString(pptInfo.coverUrl)
    };
  }

  function extractDownloadUrl(payload) {
    return firstPathText(payload, [
      "data.fileUrl",
      "data.url",
      "data.downloadUrl",
      "data.pptInfo.fileUrl",
      "fileUrl",
      "url",
      "downloadUrl"
    ]) || firstText(payload, ["fileUrl", "downloadUrl"]);
  }

  function mergeNewPageResult(result, payload) {
    if (!payload || typeof payload !== "object") return result;
    var pptxId = firstPathText(payload, [
      "pptxId",
      "pptId",
      "data.pptxId",
      "data.pptId",
      "data.id",
      "data.pptInfo.id",
      "id"
    ]);
    var fileUrl = extractDownloadUrl(payload);
    var templateId = firstPathText(payload, ["templateId", "data.templateId", "data.pptInfo.templateId"]);
    if (pptxId) result.pptxId = pptxId;
    if (fileUrl) result.fileUrl = fileUrl;
    if (templateId) result.templateId = templateId;
    return result;
  }

  function sleep(ms) {
    return new Promise(function (resolve) {
      global.setTimeout(resolve, ms);
    });
  }

  async function readStreamText(response, onProgress, extractor) {
    if (!response.body || !response.body.getReader) {
      var text = await response.text();
      var direct = tryParseJson(text);
      if (direct) {
        ensureSuccessEnvelope(direct);
        return extractor ? extractor(direct, text) : text;
      }
      return text;
    }

    var reader = response.body.getReader();
    var decoder = new TextDecoder("utf-8");
    var raw = "";
    var finalValue = "";
    var accumulated = "";

    while (true) {
      var part = await reader.read();
      if (part.done) break;

      raw += decoder.decode(part.value, { stream: true });
      var lines = raw.split(/\r?\n/);
      raw = lines.pop() || "";

      for (var i = 0; i < lines.length; i += 1) {
        var line = lines[i].trim();
        if (!line) continue;
        if (line.indexOf("data:") === 0) line = line.slice(5).trim();
        if (!line || line === "[DONE]") continue;

        var payload = tryParseJson(line);
        if (!payload) {
          accumulated += line;
          if (onProgress) onProgress(accumulated);
          continue;
        }

        ensureSuccessEnvelope(payload);
        var chunk = firstText(payload, ["text", "content", "result"]);
        if (chunk) {
          accumulated += chunk;
          if (onProgress) onProgress(accumulated);
        }

        var status = payload.status;
        if (status === 4 || status === "4") {
          var extracted = extractor ? extractor(payload, line) : "";
          if (extracted) finalValue = extracted;
        }
      }
    }

    if (finalValue) return finalValue;
    if (accumulated) return accumulated;
    var directJson = tryParseJson(raw);
    if (directJson) {
      ensureSuccessEnvelope(directJson);
      return extractor ? extractor(directJson, raw) : raw;
    }
    return raw;
  }

  async function readNewPageResultStream(response, onProgress) {
    var result = {
      pptxId: "",
      templateId: "",
      fileUrl: "",
      raw: ""
    };

    if (!response.body || !response.body.getReader) {
      var text = await response.text();
      result.raw = text;
      var direct = tryParseJson(text);
      if (direct) {
        ensureSuccessEnvelope(direct);
        mergeNewPageResult(result, direct);
      }
      return result;
    }

    var reader = response.body.getReader();
    var decoder = new TextDecoder("utf-8");
    var buffer = "";

    try {
      while (true) {
        var part = await reader.read();
        if (part.done) break;

        buffer += decoder.decode(part.value, { stream: true });
        var lines = buffer.split(/\r?\n/);
        buffer = lines.pop() || "";

        for (var i = 0; i < lines.length; i += 1) {
          var line = lines[i].trim();
          if (!line) continue;
          if (line.indexOf("data:") === 0) line = line.slice(5).trim();
          if (!line || line === "[DONE]") continue;

          result.raw += line + "\n";
          if (onProgress) onProgress(line);

          var payload = tryParseJson(line);
          if (!payload) continue;
          ensureSuccessEnvelope(payload);
          mergeNewPageResult(result, payload);
        }
      }
    } catch (error) {
      if (!result.pptxId && !result.fileUrl) throw error;
    }

    if (buffer.trim()) {
      result.raw += buffer.trim() + "\n";
      var trailing = tryParseJson(buffer.trim());
      if (trailing) {
        ensureSuccessEnvelope(trailing);
        mergeNewPageResult(result, trailing);
      }
    }

    return result;
  }

  function DocmeeClient(config) {
    this.config = config || {};
  }

  DocmeeClient.prototype.endpoint = function (path) {
    return normalizeBaseUrl(this.config.baseUrl) + path;
  };

  DocmeeClient.prototype.jsonHeaders = function (langOverride) {
    var headers = {
      "Accept": "*/*",
      "Content-Type": "application/json",
      "lang": langOverride || this.config.lang || "zh"
    };
    if (this.config.token) headers.token = this.config.token;
    if (this.config.companyId) headers.companyId = this.config.companyId;
    return headers;
  };

  DocmeeClient.prototype.formHeaders = function (langOverride) {
    var headers = {
      "Accept": "*/*",
      "lang": langOverride || this.config.lang || "zh"
    };
    if (this.config.token) headers.token = this.config.token;
    if (this.config.companyId) headers.companyId = this.config.companyId;
    return headers;
  };

  DocmeeClient.prototype.requireToken = function () {
    if (!this.config.token) {
      throw new Error("缺少 Docmee token，请检查代码内置配置。");
    }
  };

  DocmeeClient.prototype.createTask = async function (content, type) {
    this.requireToken();
    var form = new FormData();
    form.append("type", type || "1");
    form.append("content", String(content || "").trim());

    var response = await fetch(this.endpoint("/api/ppt/v2/createTask"), {
      method: "POST",
      headers: this.formHeaders(),
      body: form
    });
    var payload = await this.parseJsonResponse(response);
    var taskId = extractTaskId(payload);
    if (!taskId) throw new Error("Docmee 创建任务成功，但没有返回任务 ID。");
    return taskId;
  };

  DocmeeClient.prototype.createFileTask = async function (file) {
    this.requireToken();
    if (!file) throw new Error("请选择文档。");
    var type = /\.(xmind|mm)$/i.test(file.name || "") ? "3" : "2";
    var form = new FormData();
    form.append("type", type);
    form.append("file", file, file.name || "document");

    var response = await fetch(this.endpoint("/api/ppt/v2/createTask"), {
      method: "POST",
      headers: this.formHeaders(),
      body: form
    });
    var payload = await this.parseJsonResponse(response);
    var taskId = extractTaskId(payload);
    if (!taskId) throw new Error("Docmee 创建文档任务成功，但没有返回任务 ID。");
    return taskId;
  };

  DocmeeClient.prototype.generateMarkdown = async function (taskId, onProgress) {
    this.requireToken();
    var body = {
      id: String(taskId || "").trim(),
      stream: true,
      outlineType: "MD",
      questionMode: false,
      isNeedAsk: false,
      length: "long",
      scene: "产品介绍",
      audience: "客户",
      lang: this.config.lang || "zh",
      prompt: "语气专业，适合演示",
      aiSearch: false,
      isGenImg: false
    };

    var response = await fetch(this.endpoint("/api/ppt/v2/generateContent"), {
      method: "POST",
      headers: this.jsonHeaders(),
      body: JSON.stringify(body)
    });
    this.ensureHttpSuccess(response);
    return readStreamText(response, onProgress, function (payload) {
      return extractMarkdown(payload);
    });
  };

  DocmeeClient.prototype.listTemplates = async function (filters) {
    this.requireToken();
    var page = Math.max(1, Number(filters && filters.page) || 1);
    var size = Math.max(1, Number(filters && filters.size) || 10);
    var body = {
      page: page,
      size: size,
      filters: {
        type: 1,
        category: filters && filters.category ? filters.category : null,
        style: filters && filters.style ? filters.style : null,
        themeColor: filters && filters.themeColor ? filters.themeColor : null
      }
    };

    var response = await fetch(this.endpoint("/api/ppt/templates?lang=zh-CN"), {
      method: "POST",
      headers: this.jsonHeaders(),
      body: JSON.stringify(body)
    });
    var payload = await this.parseJsonResponse(response);
    var list = Array.isArray(payload.data) ? payload.data : [];
    if (!list.length && payload.data && Array.isArray(payload.data.list)) list = payload.data.list;
    if (!list.length && payload.data && Array.isArray(payload.data.records)) list = payload.data.records;
    return list.map(function (item) {
      return {
        id: getString(item.id),
        name: getString(item.name || item.subject),
        category: getString(item.category),
        style: getString(item.style),
        themeColor: getString(item.themeColor),
        coverUrl: getString(item.coverUrl)
      };
    }).filter(function (item) {
      return item.id;
    });
  };

  DocmeeClient.prototype.generatePptx = async function (taskId, templateId, markdown) {
    this.requireToken();
    var body = {
      id: String(taskId || "").trim(),
      templateId: String(templateId || "").trim(),
      markdown: String(markdown || "").trim()
    };
    var response = await fetch(this.endpoint("/api/ppt/v2/generatePptx"), {
      method: "POST",
      headers: this.jsonHeaders(),
      body: JSON.stringify(body)
    });
    var payload = await this.parseJsonResponse(response);
    var info = extractPptInfo(payload);
    if (!info || !info.id) throw new Error("Docmee 已生成 PPT，但未返回 PPT ID。");
    return info;
  };

  DocmeeClient.prototype.downloadPptx = async function (pptId, refresh, maxAttempts, retryDelayMs) {
    this.requireToken();
    var id = String(pptId || "").trim();
    if (!id) throw new Error("缺少 PPT ID。");
    var attempts = Math.max(1, Number(maxAttempts) || 1);
    var delay = Math.max(200, Number(retryDelayMs) || 1000);
    for (var attempt = 1; attempt <= attempts; attempt += 1) {
      var response = await fetch(this.endpoint("/api/ppt/downloadPptx"), {
        method: "POST",
        headers: this.jsonHeaders(),
        body: JSON.stringify({ id: id, refresh: !!refresh })
      });
      var payload = await this.parseJsonResponse(response);
      var fileUrl = extractDownloadUrl(payload);
      if (fileUrl) return fileUrl;
      if (attempt < attempts) await sleep(delay);
    }
    throw new Error("Docmee 返回成功，但没有 PPT 下载地址。");
  };

  DocmeeClient.prototype.textOptimize = async function (options, onProgress) {
    this.requireToken();
    var type = String(options && options.type ? options.type : "rs").trim();
    var targetLanguageCode = String(options && options.targetLanguageCode ? options.targetLanguageCode : "").trim();
    var targetLanguageName = String(options && options.targetLanguageName ? options.targetLanguageName : "").trim();
    var requestLang = type === "fy" && targetLanguageCode ? targetLanguageCode : (this.config.lang || "zh");
    var body = {
      id: String(options && options.id ? options.id : this.config.pptxId || "2029735886777888768").trim(),
      type: type,
      text: String(options && options.text ? options.text : ""),
      stream: true,
      lang: requestLang
    };
    if (!body.lang) body.lang = "zh";
    if (!body.text.trim()) throw new Error("缺少要创作的文本。");
    if (type === "fy") {
      if (targetLanguageCode) {
        body.targetLang = targetLanguageCode;
        body.language = targetLanguageCode;
      }
      if (targetLanguageName) body.targetLanguage = targetLanguageName;
    }

    var response = await fetch(this.endpoint("/api/ppt/textOptimize"), {
      method: "POST",
      headers: this.jsonHeaders(requestLang),
      body: JSON.stringify(body)
    });
    this.ensureHttpSuccess(response);
    return readStreamText(response, onProgress, function (payload) {
      return firstText(payload, ["result", "text", "content"]);
    });
  };

  DocmeeClient.prototype.newPageWithAiV2 = async function (content, onProgress) {
    this.requireToken();
    var templateId = randomNewPageTemplateId();
    var pptxId = String(this.config.pptxId || "").trim() || randomNewPageTemplateId();
    var body = {
      content: String(content || "").trim(),
      pptxId: pptxId,
      templateId: templateId,
      stream: true,
      lang: this.config.lang || "zh"
    };
    if (!body.content) throw new Error("请输入页面标题或内容。");

    var response = await fetch(this.endpoint("/api/ppt/v2/newPageWithAiV2"), {
      method: "POST",
      headers: this.jsonHeaders(),
      body: JSON.stringify(body)
    });
    this.ensureHttpSuccess(response);
    var result = await readNewPageResultStream(response, onProgress);
    if (!result.pptxId) result.pptxId = pptxId;
    if (!result.templateId) result.templateId = templateId;
    if (!result.fileUrl) {
      result.fileUrl = await this.downloadPptx(result.pptxId, true, 6, 1200);
    }
    return result;
  };

  DocmeeClient.prototype.parseJsonResponse = async function (response) {
    this.ensureHttpSuccess(response);
    var text = await response.text();
    var payload = tryParseJson(text);
    if (!payload) throw new Error("接口返回不是有效 JSON。");
    ensureSuccessEnvelope(payload);
    return payload;
  };

  DocmeeClient.prototype.ensureHttpSuccess = function (response) {
    if (!response.ok) {
      throw new Error("接口请求失败：" + response.status + " " + response.statusText);
    }
  };

  global.DocmeeClient = DocmeeClient;
})(window);
