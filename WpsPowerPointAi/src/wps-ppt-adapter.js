(function (global) {
  "use strict";

  function getApplication() {
    if (global.wps && typeof global.wps.WppApplication === "function") {
      return global.wps.WppApplication();
    }
    if (global.Application && global.Application.ActivePresentation) return global.Application;
    if (global.ActivePresentation) return { ActivePresentation: global.ActivePresentation };
    if (global.wps && global.wps.Application) return global.wps.Application;
    throw new Error("未检测到 WPS 演示宿主，请在 WPS 加载项中使用。");
  }

  function getPresentation(app) {
    if (global.ActivePresentation) return global.ActivePresentation;
    if (app.ActivePresentation) return app.ActivePresentation;
    if (app.Presentations && app.Presentations.ActivePresentation) return app.Presentations.ActivePresentation;
    throw new Error("未找到当前演示文稿。");
  }

  function getAddinRoot() {
    return global.Application || (global.wps && global.wps.Application) || null;
  }

  function getEnv(app) {
    var root = getAddinRoot();
    return (root && root.Env) || (global.wps && global.wps.Env) || (app && app.Env) || null;
  }

  function getFileSystem(app) {
    var root = getAddinRoot();
    return (root && root.FileSystem) || (global.wps && global.wps.FileSystem) || (app && app.FileSystem) || null;
  }

  function collectionCount(collection) {
    if (!collection) return 0;
    return Number(collection.Count || collection.count || collection.length || 0);
  }

  function item(collection, index) {
    if (!collection) return null;
    if (typeof collection.Item === "function") return collection.Item(index);
    if (typeof collection.item === "function") return collection.item(index);
    return collection[index - 1] || null;
  }

  function hexToOfficeRgb(hex) {
    var clean = String(hex || "#4f46e5").replace("#", "");
    var r = parseInt(clean.slice(0, 2), 16);
    var g = parseInt(clean.slice(2, 4), 16);
    var b = parseInt(clean.slice(4, 6), 16);
    return r + g * 256 + b * 65536;
  }

  function ensureTrailingSlash(path) {
    path = String(path || "");
    if (!path) return "";
    return /[\\/]$/.test(path) ? path : path + "\\";
  }

  function safeFileName(value) {
    return String(value || Date.now()).replace(/[^\w.-]+/g, "_");
  }

  function toFileUrl(filePath) {
    var normalized = String(filePath || "").replace(/\\/g, "/");
    if (!normalized) return "";
    if (/^[a-zA-Z]:\//.test(normalized)) return "file:///" + encodeURI(normalized);
    return "file://" + encodeURI(normalized);
  }

  function parseDownloadPath(value) {
    if (!value) return "";
    if (typeof value === "string") {
      var text = value.trim();
      if (!text) return "";
      if ((text[0] === "{" && text[text.length - 1] === "}") || (text[0] === "[" && text[text.length - 1] === "]")) {
        try {
          return parseDownloadPath(JSON.parse(text));
        } catch (error) {}
      }
      text = text.replace(/^["']+|["']+$/g, "").replace(/"/g, "").trim();
      return text;
    }
    if (typeof value === "object") {
      return parseDownloadPath(value.path || value.filePath || value.localPath || value.data || value.result || "");
    }
    return "";
  }

  async function fileExists(fileSystem, filePath) {
    if (!fileSystem || !filePath) return null;
    try {
      var result = null;
      if (typeof fileSystem.existsSync === "function") result = fileSystem.existsSync(filePath);
      else if (typeof fileSystem.exists === "function") result = fileSystem.exists(filePath);
      else if (typeof fileSystem.FileExists === "function") result = fileSystem.FileExists(filePath);
      else if (typeof fileSystem.Exists === "function") result = fileSystem.Exists(filePath);
      else return null;
      if (result && typeof result.then === "function") result = await result;
      return !!result;
    } catch (error) {
      return false;
    }
    return null;
  }

  function sleep(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  async function waitForFileReady(fileSystem, filePath, timeoutMs) {
    var deadline = Date.now() + Math.max(300, Number(timeoutMs) || 2000);

    while (Date.now() <= deadline) {
      var exists = await fileExists(fileSystem, filePath);
      if (exists === true || exists === null) return filePath;
      await sleep(120);
    }

    throw new Error("PPTX 已下载，但 WPS 临时文件还没有准备好，请稍后重试。");
  }

  function arrayBufferToBinaryString(buffer) {
    var bytes = new Uint8Array(buffer);
    var chunkSize = 0x8000;
    var chunks = [];
    for (var i = 0; i < bytes.length; i += chunkSize) {
      chunks.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize)));
    }
    return chunks.join("");
  }

  async function writeBinaryFile(fileSystem, filePath, binary) {
    if (!fileSystem) throw new Error("当前 WPS 环境没有 FileSystem，无法保存 PPTX。");
    if (typeof fileSystem.writeAsBinaryString === "function") {
      var writeResult = fileSystem.writeAsBinaryString(filePath, binary);
      if (writeResult && typeof writeResult.then === "function") await writeResult;
      return;
    }
    if (typeof fileSystem.WriteFile === "function") {
      var fallbackWriteResult = fileSystem.WriteFile(filePath, binary);
      if (fallbackWriteResult && typeof fallbackWriteResult.then === "function") await fallbackWriteResult;
      return;
    }
    throw new Error("当前 WPS FileSystem 不支持写入二进制文件。");
  }

  async function downloadFileToTemp(url, pptId) {
    if (!url) throw new Error("缺少 PPTX 下载地址。");
    var app = getApplication();
    var fileSystem = getFileSystem(app);
    var env = getEnv(app);
    var tempPath = "";
    if (env && typeof env.GetTempPath === "function") tempPath = env.GetTempPath();
    if (!tempPath && fileSystem && typeof fileSystem.tmpdir === "function") tempPath = fileSystem.tmpdir();
    if (!tempPath) throw new Error("当前 WPS 环境没有临时目录 API，无法保存 PPTX。");

    var localPath = ensureTrailingSlash(tempPath) + "wenduoduoAI_" + safeFileName(pptId) + "_" + Date.now() + ".pptx";
    var response = await fetch(url);
    if (!response.ok) throw new Error("下载 PPTX 失败：" + response.status + " " + response.statusText);
    var buffer = await response.arrayBuffer();
    await writeBinaryFile(fileSystem, localPath, arrayBufferToBinaryString(buffer));
    await waitForFileReady(fileSystem, localPath, 3000);
    return localPath;
  }

  async function deleteTempFile(fileSystem, filePath) {
    if (!fileSystem || !filePath) return;
    try {
      var result = null;
      if (typeof fileSystem.unlink === "function") result = fileSystem.unlink(filePath);
      else if (typeof fileSystem.deleteFile === "function") result = fileSystem.deleteFile(filePath);
      else if (typeof fileSystem.DeleteFile === "function") result = fileSystem.DeleteFile(filePath);
      else if (typeof fileSystem.remove === "function") result = fileSystem.remove(filePath);
      if (result && typeof result.then === "function") await result;
    } catch (error) {}
  }

  function tempFilePath(app, prefix, extension) {
    var fileSystem = getFileSystem(app);
    var env = getEnv(app);
    var tempPath = "";
    if (env && typeof env.GetTempPath === "function") tempPath = env.GetTempPath();
    if (!tempPath && fileSystem && typeof fileSystem.tmpdir === "function") tempPath = fileSystem.tmpdir();
    if (!tempPath) return "";
    return ensureTrailingSlash(tempPath) + prefix + "_" + Date.now() + "_" + Math.floor(Math.random() * 1000000) + extension;
  }

  async function maybeAwait(value) {
    if (value && typeof value.then === "function") return await value;
    return value;
  }

  async function openPresentation(app, filePath, withWindow) {
    if (!app || !app.Presentations || typeof app.Presentations.Open !== "function") {
      throw new Error("当前 WPS 环境不支持打开源 PPTX。");
    }

    filePath = parseDownloadPath(filePath);
    if (!filePath) throw new Error("缺少要打开的 PPTX 本地路径。");

    try {
      return await maybeAwait(app.Presentations.Open(filePath, true, false, !!withWindow));
    } catch (error) {
      return await maybeAwait(app.Presentations.Open(filePath));
    }
  }

  async function closePresentation(presentation) {
    if (!presentation || typeof presentation.Close !== "function") return;
    try {
      await maybeAwait(presentation.Close());
    } catch (error) {}
  }

  function collectSlides(slides, startIndex, count) {
    var importedSlides = [];
    for (var i = 0; i < count; i += 1) {
      var slide = item(slides, startIndex + i);
      if (slide) importedSlides.push(slide);
    }
    return importedSlides;
  }

  async function exportSlideBackdropAsPicture(app, sourceSlide, destSlide) {
    var duplicateRange = null;
    var tempSlide = null;
    var bgPath = tempFilePath(app, "wenduoduoAI_bg", ".png");
    var fileSystem = getFileSystem(app);
    if (!bgPath || !sourceSlide || !destSlide) return false;

    try {
      duplicateRange = await maybeAwait(sourceSlide.Duplicate());
      tempSlide = item(duplicateRange, 1) || duplicateRange;
      while (tempSlide && tempSlide.Shapes && collectionCount(tempSlide.Shapes) > 0) {
        await maybeAwait(item(tempSlide.Shapes, 1).Delete());
      }
      if (!tempSlide || typeof tempSlide.Export !== "function") return false;
      await maybeAwait(tempSlide.Export(bgPath, "PNG"));
      if (!destSlide.Background || !destSlide.Background.Fill || typeof destSlide.Background.Fill.UserPicture !== "function") return false;
      destSlide.FollowMasterBackground = 0;
      await maybeAwait(destSlide.Background.Fill.UserPicture(bgPath));
      try {
        destSlide.Background.Fill.Visible = -1;
      } catch (error) {}
      return true;
    } catch (error) {
      return false;
    } finally {
      if (tempSlide && typeof tempSlide.Delete === "function") {
        try {
          await maybeAwait(tempSlide.Delete());
        } catch (error) {}
      }
      await deleteTempFile(fileSystem, bgPath);
    }
  }

  async function copySlideBackground(app, sourceSlide, destSlide) {
    if (!sourceSlide || !destSlide || !sourceSlide.Background || !destSlide.Background) return false;

    try {
      destSlide.FollowMasterBackground = 0;
    } catch (error) {}

    try {
      var srcFill = sourceSlide.Background.Fill;
      var dstFill = destSlide.Background.Fill;
      var sourceFollowsMaster = Number(sourceSlide.FollowMasterBackground || 0) !== 0;
      var fillType = Number(srcFill && srcFill.Type ? srcFill.Type : 0);

      if (!sourceFollowsMaster && (fillType === 1 || fillType === 3) && dstFill) {
        if (srcFill.ForeColor && dstFill.ForeColor) dstFill.ForeColor.RGB = srcFill.ForeColor.RGB;
        if (srcFill.BackColor && dstFill.BackColor) dstFill.BackColor.RGB = srcFill.BackColor.RGB;
        dstFill.Visible = -1;
        return true;
      }
    } catch (error) {}

    return await exportSlideBackdropAsPicture(app, sourceSlide, destSlide);
  }

  async function restoreImportedSlideBackgrounds(app, sourcePath, importedSlides) {
    if (!importedSlides || importedSlides.length === 0) return 0;
    var sourcePresentation = null;
    var restoredCount = 0;

    try {
      sourcePresentation = await openPresentation(app, sourcePath, false);
      var sourceSlides = sourcePresentation && sourcePresentation.Slides;
      var maxCount = Math.min(importedSlides.length, collectionCount(sourceSlides));
      for (var i = 0; i < maxCount; i += 1) {
        if (await copySlideBackground(app, item(sourceSlides, i + 1), importedSlides[i])) restoredCount += 1;
      }
    } catch (error) {
      restoredCount = 0;
    } finally {
      await closePresentation(sourcePresentation);
    }

    return restoredCount;
  }

  async function importPptxFileIntoPresentation(presentation, localPath, insertAfterIndex) {
    var slides = presentation.Slides;
    var insertedCount = 0;

    try {
      var insertResult = slides.InsertFromFile(localPath, insertAfterIndex);
      insertedCount = Number(await maybeAwait(insertResult) || 0);
    } catch (error) {
      var fallbackInsertResult = slides.InsertFromFile(localPath, insertAfterIndex, 1, 9999);
      insertedCount = Number(await maybeAwait(fallbackInsertResult) || 0);
    }

    if (!insertedCount || insertedCount < 0 || Number.isNaN(insertedCount)) {
      insertedCount = Math.max(0, collectionCount(slides) - insertAfterIndex);
    }

    return collectSlides(slides, insertAfterIndex + 1, insertedCount);
  }

  async function copyPptxSlidesIntoPresentation(app, presentation, localPath, insertAfterIndex) {
    var sourcePresentation = null;
    var importedSlides = [];
    var targetSlides = presentation.Slides;

    try {
      sourcePresentation = await openPresentation(app, localPath, false);
      var sourceSlides = sourcePresentation.Slides;
      var sourceCount = collectionCount(sourceSlides);

      for (var i = 1; i <= sourceCount; i += 1) {
        var beforeCount = collectionCount(targetSlides);
        var pasteIndex = Math.min(insertAfterIndex + importedSlides.length + 1, beforeCount + 1);
        await maybeAwait(item(sourceSlides, i).Copy());
        var pastedRange = await maybeAwait(targetSlides.Paste(pasteIndex));
        var pastedSlide = item(pastedRange, 1) || item(targetSlides, pasteIndex);
        if (pastedSlide) importedSlides.push(pastedSlide);
      }
    } finally {
      await closePresentation(sourcePresentation);
    }

    return importedSlides;
  }

  async function exportSlidePreview(app, slide, pptId, slideIndex) {
    var previewPath = tempFilePath(app, "wenduoduoAI_preview_" + safeFileName(pptId) + "_" + slideIndex, ".png");
    if (!previewPath || !slide || typeof slide.Export !== "function") return "";
    try {
      await maybeAwait(slide.Export(previewPath, "PNG", 960, 540));
      return previewPath;
    } catch (error) {
      return "";
    }
  }

  async function replaceCurrentSlideFromPptx(app, presentation, localPath, slideIndex) {
    var slides = presentation.Slides;
    var originalIndex = activeSlideIndex(app, presentation);
    var originalSlide = item(slides, originalIndex);
    if (!originalSlide) throw new Error("请先选中要替换的幻灯片。");

    var sourcePresentation = null;
    var insertedSlide = null;

    try {
      var insertedCount = 0;
      try {
        insertedCount = Number(await maybeAwait(slides.InsertFromFile(localPath, originalIndex, slideIndex, slideIndex)) || 0);
      } catch (error) {
        insertedCount = 0;
      }

      if (insertedCount > 0) {
        insertedSlide = item(slides, originalIndex + 1);
      } else {
        sourcePresentation = await openPresentation(app, localPath, false);
        var sourceSlide = item(sourcePresentation.Slides, slideIndex);
        await maybeAwait(sourceSlide.Copy());
        var pastedRange = await maybeAwait(slides.Paste(originalIndex + 1));
        insertedSlide = item(pastedRange, 1) || item(slides, originalIndex + 1);
      }

      if (!insertedSlide) throw new Error("候选页面导入失败。");

      if (!sourcePresentation) sourcePresentation = await openPresentation(app, localPath, false);
      await copySlideBackground(app, item(sourcePresentation.Slides, slideIndex), insertedSlide);

      await maybeAwait(originalSlide.Delete());
      var targetIndex = Math.max(1, Math.min(originalIndex, collectionCount(slides)));
      try {
        item(slides, targetIndex).Select();
        if (app.ActiveWindow && app.ActiveWindow.View && typeof app.ActiveWindow.View.GotoSlide === "function") {
          app.ActiveWindow.View.GotoSlide(targetIndex);
        }
      } catch (error) {}
      return true;
    } finally {
      await closePresentation(sourcePresentation);
    }
  }

  function setColor(target, officeRgb) {
    try {
      if (target && target.ForeColor) target.ForeColor.RGB = officeRgb;
    } catch (error) {
      return false;
    }
    return true;
  }

  function selectedShapeRange(app) {
    try {
      var selection = app.ActiveWindow && app.ActiveWindow.Selection;
      if (!selection) return null;
      if (selection.ShapeRange) return selection.ShapeRange;
    } catch (error) {
      return null;
    }
    return null;
  }

  function selectedTextRange(app) {
    try {
      var selection = app.ActiveWindow && app.ActiveWindow.Selection;
      if (!selection) return null;
      if (selection.TextRange) return selection.TextRange;
    } catch (error) {
      return null;
    }
    return null;
  }

  function readShapeText(shape) {
    try {
      if (shape.TextFrame && shape.TextFrame.TextRange) {
        return String(shape.TextFrame.TextRange.Text || "");
      }
    } catch (error) {
      return "";
    }
    return "";
  }

  function writeShapeText(shape, text) {
    try {
      if (shape.TextFrame && shape.TextFrame.TextRange) {
        shape.TextFrame.TextRange.Text = text;
        return true;
      }
    } catch (error) {
      return false;
    }
    return false;
  }

  function getShapeText(shape) {
    return readShapeText(shape).trim();
  }

  function pushLabeledText(lines, seen, label, text) {
    String(text || "").split(/\r?\n/).forEach(function (line) {
      var value = line.trim();
      if (!value || seen[label + ":" + value]) return;
      seen[label + ":" + value] = true;
      lines.push(label ? label + ": " + value : value);
    });
  }

  function pushPlainText(lines, seen, text) {
    String(text || "").split(/\r?\n/).forEach(function (line) {
      var value = line.trim();
      if (!value || seen[value]) return;
      seen[value] = true;
      lines.push(value);
    });
  }

  function collectShapePlainText(shape, lines, seen) {
    if (!shape) return;
    pushPlainText(lines, seen, getShapeText(shape));

    try {
      pushLabeledText(lines, seen, "替代文本", shape.AlternativeText || shape.Title);
    } catch (error) {}

    try {
      var chart = shape.Chart;
      if (chart) {
        if (chart.HasTitle && chart.ChartTitle) pushLabeledText(lines, seen, "图表标题", chart.ChartTitle.Text);
        try {
          var axes = chart.Axes();
          var axisCount = collectionCount(axes);
          for (var axisIndex = 1; axisIndex <= axisCount; axisIndex += 1) {
            var axis = item(axes, axisIndex);
            if (axis && axis.HasTitle && axis.AxisTitle) pushLabeledText(lines, seen, "图表轴标题", axis.AxisTitle.Text);
          }
        } catch (error) {}
      }
    } catch (error) {}

    try {
      var groupItems = shape.GroupItems;
      var groupCount = collectionCount(groupItems);
      for (var i = 1; i <= groupCount; i += 1) {
        collectShapePlainText(item(groupItems, i), lines, seen);
      }
    } catch (error) {}

    try {
      var table = shape.Table;
      var rowCount = collectionCount(table && table.Rows);
      var columnCount = collectionCount(table && table.Columns);
      for (var r = 1; r <= rowCount; r += 1) {
        for (var c = 1; c <= columnCount; c += 1) {
          var cell = table.Cell(r, c);
          collectShapePlainText(cell && cell.Shape, lines, seen);
        }
      }
    } catch (error) {}
  }

  function collectNotesPlainText(slide, lines, seen) {
    try {
      var notesPage = slide.NotesPage;
      var notesShapes = notesPage && notesPage.Shapes;
      var count = collectionCount(notesShapes);
      for (var i = 1; i <= count; i += 1) {
        var text = getShapeText(item(notesShapes, i));
        if (text) pushLabeledText(lines, seen, "备注", text);
      }
    } catch (error) {}
  }

  function activeSlideIndex(app, presentation) {
    try {
      var selection = app.ActiveWindow && app.ActiveWindow.Selection;
      if (selection && selection.SlideRange && collectionCount(selection.SlideRange) > 0) {
        var selectedSlide = item(selection.SlideRange, 1);
        if (selectedSlide && selectedSlide.SlideIndex) return Number(selectedSlide.SlideIndex);
      }
    } catch (error) {}

    try {
      var viewSlide = app.ActiveWindow && app.ActiveWindow.View && app.ActiveWindow.View.Slide;
      if (viewSlide && viewSlide.SlideIndex) return Number(viewSlide.SlideIndex);
    } catch (error) {}

    return Math.max(1, Math.min(collectionCount(presentation && presentation.Slides), 1));
  }

  function getSlideTitle(slide) {
    if (!slide) return "";
    try {
      if (slide.Shapes && slide.Shapes.HasTitle && slide.Shapes.Title) {
        var title = getShapeText(slide.Shapes.Title);
        if (title) return title;
      }
    } catch (error) {}

    try {
      var shapes = slide.Shapes;
      var count = collectionCount(shapes);
      for (var i = 1; i <= count; i += 1) {
        var text = getShapeText(item(shapes, i));
        if (text) return text;
      }
    } catch (error) {}

    return "候选页面";
  }

  function getSlidePlainText(slide) {
    var lines = [];
    var seen = {};
    var title = getSlideTitle(slide);
    if (title && title !== "候选页面") pushLabeledText(lines, seen, "页面标题", title);
    try {
      var shapes = slide && slide.Shapes;
      var count = collectionCount(shapes);
      for (var i = 1; i <= count; i += 1) {
        collectShapePlainText(item(shapes, i), lines, seen);
      }
    } catch (error) {}
    collectNotesPlainText(slide, lines, seen);
    return lines.join("\n");
  }

  var adapter = {
    isWpsHost: function () {
      return Boolean(global.wps || global.Application || global.ActivePresentation);
    },

    getSelectedText: function () {
      var app = getApplication();
      var textRange = selectedTextRange(app);
      if (textRange && textRange.Text) return String(textRange.Text || "");
      var range = selectedShapeRange(app);
      var count = collectionCount(range);
      var texts = [];
      for (var i = 1; i <= count; i += 1) {
        var text = readShapeText(item(range, i));
        if (text) texts.push(text);
      }
      return texts.join("\n");
    },

    replaceSelectedText: function (text) {
      var app = getApplication();
      var textRange = selectedTextRange(app);
      if (textRange && String(textRange.Text || "")) {
        textRange.Text = text;
        return;
      }
      var range = selectedShapeRange(app);
      var count = collectionCount(range);
      if (!count) throw new Error("请先在 WPS 演示中选中一个文本框。");
      var changed = false;
      for (var i = 1; i <= count; i += 1) {
        changed = writeShapeText(item(range, i), text) || changed;
        if (changed) break;
      }
      if (!changed) throw new Error("选中对象没有可替换的文本。");
    },

    getCurrentSlideText: function () {
      var app = getApplication();
      var presentation = getPresentation(app);
      var slide = item(presentation.Slides, activeSlideIndex(app, presentation));
      return getSlidePlainText(slide);
    },

    applyThemeColor: function (hex) {
      var app = getApplication();
      var presentation = getPresentation(app);
      var slides = presentation.Slides;
      var slideCount = collectionCount(slides);
      var officeRgb = hexToOfficeRgb(hex);
      var changed = 0;

      for (var s = 1; s <= slideCount; s += 1) {
        var slide = item(slides, s);
        var shapes = slide && slide.Shapes;
        var shapeCount = collectionCount(shapes);
        for (var i = 1; i <= shapeCount; i += 1) {
          var shape = item(shapes, i);
          try {
            if (shape.Fill && shape.Fill.ForeColor) {
              shape.Fill.ForeColor.RGB = officeRgb;
              changed += 1;
            }
          } catch (error) {}
          try {
            if (shape.Line && shape.Line.ForeColor) {
              shape.Line.ForeColor.RGB = officeRgb;
              changed += 1;
            }
          } catch (error) {}
          try {
            if (shape.TextFrame && shape.TextFrame.TextRange && shape.TextFrame.TextRange.Font) {
              setColor(shape.TextFrame.TextRange.Font.Color, officeRgb);
              changed += 1;
            }
          } catch (error) {}
        }
      }

      return changed;
    },

    openPptxDownload: function (url) {
      if (!url) throw new Error("缺少 PPTX 下载地址。");
      global.open(url, "_blank");
    },

    createSinglePageChoicesFromUrl: async function (url, pptId) {
      var app = getApplication();
      var localPath = await downloadFileToTemp(url, pptId);
      var sourcePresentation = null;
      var choices = [];

      try {
        sourcePresentation = await openPresentation(app, localPath, false);
        var slides = sourcePresentation.Slides;
        var slideCount = collectionCount(slides);
        for (var i = 1; i <= slideCount; i += 1) {
          var slide = item(slides, i);
          var previewPath = await exportSlidePreview(app, slide, pptId, i);
          choices.push({
            slideIndex: i,
            title: getSlideTitle(slide),
            previewPath: previewPath,
            previewUrl: toFileUrl(previewPath),
            localPath: localPath
          });
        }
      } finally {
        await closePresentation(sourcePresentation);
      }

      return {
        localPath: localPath,
        choices: choices
      };
    },

    applySinglePageChoice: async function (choice) {
      if (!choice || !choice.localPath || !choice.slideIndex) throw new Error("请先选择一个候选页面。");
      var app = getApplication();
      var presentation = getPresentation(app);
      await replaceCurrentSlideFromPptx(app, presentation, choice.localPath, Number(choice.slideIndex));
      return true;
    },

    importPptxFromUrl: async function (url, pptId) {
      if (!this.isWpsHost()) {
        this.openPptxDownload(url);
        return {
          importedCount: 0,
          localPath: "",
          openedDownload: true
        };
      }

      var localPath = await downloadFileToTemp(url, pptId);
      var app = getApplication();
      var presentation = getPresentation(app);
      var slides = presentation.Slides;
      var insertAfterIndex = collectionCount(slides);
      var importedSlides = [];
      var restoredBackgrounds = 0;

      try {
        importedSlides = await importPptxFileIntoPresentation(presentation, localPath, insertAfterIndex);
      } catch (error) {
        importedSlides = await copyPptxSlidesIntoPresentation(app, presentation, localPath, insertAfterIndex);
      }

      if (!importedSlides.length) {
        importedSlides = collectSlides(slides, insertAfterIndex + 1, Math.max(0, collectionCount(slides) - insertAfterIndex));
      }

      if (!importedSlides.length) throw new Error("PPTX 已下载，但没有成功导入任何幻灯片。");

      restoredBackgrounds = await restoreImportedSlideBackgrounds(app, localPath, importedSlides);

      try {
        var firstImportedIndex = Math.max(1, insertAfterIndex + 1);
        if (slides.Item) slides.Item(firstImportedIndex).Select();
        if (app.ActiveWindow && app.ActiveWindow.View && typeof app.ActiveWindow.View.GotoSlide === "function") {
          app.ActiveWindow.View.GotoSlide(firstImportedIndex);
        }
      } catch (error) {}

      return {
        importedCount: importedSlides.length,
        restoredBackgrounds: restoredBackgrounds,
        localPath: localPath,
        openedDownload: false
      };
    },

    hideTaskPane: function () {
      var root = getAddinRoot();
      if (!root || !root.PluginStorage || !root.GetTaskPane) return;
      var paneId = root.PluginStorage.getItem("wenduoduo_wps_ppt_taskpane_id");
      if (!paneId) return;
      try {
        var pane = root.GetTaskPane(paneId);
        if (pane) pane.Visible = false;
      } catch (error) {}
    },

    closeHostSurface: function () {
      try {
        if (typeof global.close === "function") global.close();
      } catch (error) {}
      try {
        this.hideTaskPane();
      } catch (error) {}
    }
  };

  global.WpsPptAdapter = adapter;
})(window);
