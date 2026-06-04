const https = require('https');
const http = require('http');

const baseUrl = (process.env.OFFICE_AI_DOCMEE_API_BASE_URL || 'https://test.docmee.cn').replace(/\/+$/, '');
const token = process.env.OFFICE_AI_DOCMEE_TOKEN || 'ak_demo';
const runGeneration = process.env.DOCMEE_SMOKE_GENERATE === '1';

function requestJson(method, path, body, headers = {}) {
  return new Promise((resolve, reject) => {
    const url = new URL(path, baseUrl);
    const transport = url.protocol === 'http:' ? http : https;
    const payload = body == null ? null : Buffer.from(JSON.stringify(body));
    const req = transport.request(url, {
      method,
      headers: {
        token,
        ...(payload ? { 'Content-Type': 'application/json', 'Content-Length': payload.length } : {}),
        ...headers,
      },
      timeout: 60000,
    }, (res) => {
      let data = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode < 200 || res.statusCode >= 300) {
          reject(new Error(`${method} ${url.pathname} failed with ${res.statusCode}: ${data}`));
          return;
        }

        try {
          const json = JSON.parse(data);
          if (json.code !== 0) {
            reject(new Error(`${method} ${url.pathname} returned code ${json.code}: ${json.message || data}`));
            return;
          }
          resolve(json);
        } catch (error) {
          reject(new Error(`${method} ${url.pathname} returned non-JSON: ${error.message}`));
        }
      });
    });

    req.on('timeout', () => {
      req.destroy(new Error(`${method} ${path} timed out`));
    });
    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

function requestMultipart(path, fields, files = []) {
  return new Promise((resolve, reject) => {
    const boundary = `office-ai-${Date.now().toString(16)}`;
    const chunks = [];
    for (const [name, value] of Object.entries(fields)) {
      chunks.push(Buffer.from(`--${boundary}\r\n`));
      chunks.push(Buffer.from(`Content-Disposition: form-data; name="${name}"\r\n\r\n`));
      chunks.push(Buffer.from(String(value)));
      chunks.push(Buffer.from('\r\n'));
    }
    for (const file of files) {
      chunks.push(Buffer.from(`--${boundary}\r\n`));
      chunks.push(Buffer.from(`Content-Disposition: form-data; name="${file.name}"; filename="${file.filename}"\r\n`));
      chunks.push(Buffer.from(`Content-Type: ${file.contentType || 'application/octet-stream'}\r\n\r\n`));
      chunks.push(Buffer.isBuffer(file.content) ? file.content : Buffer.from(String(file.content)));
      chunks.push(Buffer.from('\r\n'));
    }
    chunks.push(Buffer.from(`--${boundary}--\r\n`));
    const payload = Buffer.concat(chunks);
    const url = new URL(path, baseUrl);
    const transport = url.protocol === 'http:' ? http : https;

    const req = transport.request(url, {
      method: 'POST',
      headers: {
        token,
        'Content-Type': `multipart/form-data; boundary=${boundary}`,
        'Content-Length': payload.length,
      },
      timeout: 30000,
    }, (res) => {
      let data = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode < 200 || res.statusCode >= 300) {
          reject(new Error(`POST ${url.pathname} failed with ${res.statusCode}: ${data}`));
          return;
        }

        try {
          const json = JSON.parse(data);
          if (json.code !== 0) {
            reject(new Error(`POST ${url.pathname} returned code ${json.code}: ${json.message || data}`));
            return;
          }
          resolve(json);
        } catch (error) {
          reject(new Error(`POST ${url.pathname} returned non-JSON: ${error.message}`));
        }
      });
    });

    req.on('timeout', () => {
      req.destroy(new Error(`POST ${path} timed out`));
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

function requestText(method, path, body, headers = {}) {
  return new Promise((resolve, reject) => {
    const url = new URL(path, baseUrl);
    const transport = url.protocol === 'http:' ? http : https;
    const payload = body == null ? null : Buffer.from(JSON.stringify(body));
    const req = transport.request(url, {
      method,
      headers: {
        token,
        ...(payload ? { 'Content-Type': 'application/json', 'Content-Length': payload.length } : {}),
        ...headers,
      },
      timeout: 120000,
    }, (res) => {
      let data = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode < 200 || res.statusCode >= 300) {
          reject(new Error(`${method} ${url.pathname} failed with ${res.statusCode}: ${data}`));
          return;
        }
        resolve(data);
      });
    });

    req.on('timeout', () => {
      req.destroy(new Error(`${method} ${path} timed out`));
    });
    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

function requestBytes(urlText) {
  return new Promise((resolve, reject) => {
    const url = new URL(urlText);
    const transport = url.protocol === 'http:' ? http : https;
    const req = transport.request(url, {
      method: 'GET',
      timeout: 60000,
    }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => { chunks.push(chunk); });
      res.on('end', () => {
        if (res.statusCode < 200 || res.statusCode >= 300) {
          reject(new Error(`GET ${url.pathname} failed with ${res.statusCode}`));
          return;
        }
        resolve(Buffer.concat(chunks));
      });
    });

    req.on('timeout', () => {
      req.destroy(new Error(`GET ${urlText} timed out`));
    });
    req.on('error', reject);
    req.end();
  });
}

function collectStrings(value, output) {
  if (typeof value === 'string') {
    output.push(value);
    return;
  }
  if (Array.isArray(value)) {
    value.forEach((item) => collectStrings(item, output));
    return;
  }
  if (value && typeof value === 'object') {
    Object.values(value).forEach((item) => collectStrings(item, output));
  }
}

function extractMarkdownFromGenerateContent(responseText) {
  const collected = [];

  for (const line of responseText.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    const payloadText = trimmed.startsWith('data:') ? trimmed.slice(5).trim() : trimmed;
    if (!payloadText || payloadText === '[DONE]') continue;

    try {
      const payload = JSON.parse(payloadText);
      collectStrings(payload, collected);
    } catch (_error) {
      collected.push(payloadText);
    }
  }

  const markdown = collected.join('\n').replace(/\\n/g, '\n');
  if (!markdown.includes('#') || !markdown.includes('##')) {
    throw new Error('Docmee generateContent for uploaded document did not return a Markdown outline');
  }

  return markdown;
}

async function verifyDocumentUploadOutline() {
  const documentContent = '# 文档生成PPT烟测\n\n## 背景\n\n这是一份用于验证 Office AI 插件文档生成 PPT 链路的最小文档。\n\n## 结论\n\n用户需要先得到可编辑 Markdown 大纲，再选择模板生成 PPT。';
  const documentTask = await requestMultipart(
    '/api/ppt/v2/createTask',
    { type: '2' },
    [{
      name: 'file',
      filename: 'office-ai-docmee-document-smoke.md',
      contentType: 'text/markdown; charset=utf-8',
      content: Buffer.from(documentContent, 'utf8'),
    }],
  );
  const taskId = documentTask.data && documentTask.data.id;
  if (!taskId) throw new Error('Docmee uploaded document createTask did not return data.id');

  const responseText = await requestText('POST', '/api/ppt/v2/generateContent', {
    id: taskId,
    stream: true,
    outlineType: 'MD',
    questionMode: false,
    isNeedAsk: false,
    length: 'short',
    scene: '产品介绍',
    audience: '客户',
    lang: 'zh',
    prompt: '提炼文档重点',
    aiSearch: false,
    isGenImg: false,
  });
  const markdown = extractMarkdownFromGenerateContent(responseText);

  return {
    taskId,
    hasMarkdownHeading: markdown.includes('#'),
    preview: markdown.replace(/\s+/g, ' ').slice(0, 120),
  };
}

async function verifyTitleOutline() {
  const titleTask = await requestMultipart('/api/ppt/v2/createTask', {
    type: '1',
    content: 'Office AI 插件标题生成 PPT 烟测',
  });
  const taskId = titleTask.data && titleTask.data.id;
  if (!taskId) throw new Error('Docmee title createTask did not return data.id');

  const responseText = await requestText('POST', '/api/ppt/v2/generateContent', {
    id: taskId,
    stream: true,
    outlineType: 'MD',
    questionMode: false,
    isNeedAsk: false,
    length: 'short',
    scene: '产品介绍',
    audience: '客户',
    lang: 'zh',
    prompt: '生成演示大纲',
    aiSearch: false,
    isGenImg: false,
  });
  const markdown = extractMarkdownFromGenerateContent(responseText);

  return {
    taskId,
    hasMarkdownHeading: markdown.includes('#'),
    preview: markdown.replace(/\s+/g, ' ').slice(0, 120),
  };
}

async function main() {
  const templateList = await requestJson('POST', '/api/ppt/templates?lang=zh-CN', {
    page: 1,
    size: 2,
    filters: { type: 1, category: null, style: null, themeColor: null },
  });
  const templates = Array.isArray(templateList.data) ? templateList.data : [];
  if (templates.length === 0) throw new Error('Docmee template list returned no templates');

  const result = {
    baseUrl,
    templates: templates.map((item) => ({ id: item.id, name: item.name })).filter((item) => item.id),
  };

  if (runGeneration) {
    result.titleGeneration = await verifyTitleOutline();
    result.documentUpload = await verifyDocumentUploadOutline();

    const markdown = '# 接口烟测\n\n## 测试章节\n\n### 测试页面\n\n- 这是用于 Office AI 插件验证的最小 Markdown 内容。';
    const task = await requestMultipart('/api/ppt/v2/createTask', { type: '7', content: markdown });
    const taskId = task.data && task.data.id;
    if (!taskId) throw new Error('Docmee createTask did not return data.id');

    const primaryTemplateId = templates[0].id;
    const ppt = await requestJson('POST', '/api/ppt/v2/generatePptx', {
      id: taskId,
      templateId: primaryTemplateId,
      markdown,
    });
    const pptInfo = ppt.data && ppt.data.pptInfo;
    if (!pptInfo || !pptInfo.id) throw new Error('Docmee generatePptx did not return data.pptInfo.id');

    const download = await requestJson('POST', '/api/ppt/downloadPptx', {
      id: pptInfo.id,
      refresh: false,
    });
    const fileUrl = download.data && download.data.fileUrl;
    if (!fileUrl) throw new Error('Docmee downloadPptx did not return data.fileUrl');
    const pptxBytes = await requestBytes(fileUrl);
    if (pptxBytes.length < 4 || pptxBytes[0] !== 0x50 || pptxBytes[1] !== 0x4b) {
      throw new Error('Docmee downloadPptx fileUrl did not return a PPTX/ZIP payload');
    }

    result.generated = {
      taskId,
      pptId: pptInfo.id,
      templateId: pptInfo.templateId,
      hasFileUrl: true,
      downloadedBytes: pptxBytes.length,
    };

    const secondaryTemplate = templates.find((item) => item.id && item.id !== primaryTemplateId);
    if (secondaryTemplate) {
      const updated = await requestJson('POST', '/api/ppt/updatePptTemplate', {
        pptId: pptInfo.id,
        templateId: secondaryTemplate.id,
        sync: false,
      });
      result.updated = {
        pptId: updated.data && updated.data.pptId,
        templateId: updated.data && updated.data.templateId,
      };
    }
  }

  console.log(JSON.stringify(result, null, 2));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
