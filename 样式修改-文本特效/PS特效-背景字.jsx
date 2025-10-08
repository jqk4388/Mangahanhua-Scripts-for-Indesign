/*
  PS特效-背景字.jsx
  说明: 将选中文本框的文字信息传入 Photoshop，在图像上重新排版文字并保存为分层PSD（或覆盖原文件），然后在 InDesign 中更新链接。
  要求: 兼容 ExtendScript (ES3 风格)，大量注释，所有功能封装为函数，使用 try-catch 捕获异常。
*/

// 主入口
function main() {
    try {
        // 1. 检查环境与文档
        if (app.documents.length === 0) {
            alert("未打开任何文档。请先打开或创建一个 InDesign 文档。");
            return;
        }
        var doc = app.activeDocument;

        // 检查文档是否已保存
        if (!isDocumentSaved(doc)) {
            alert("请先保存文档，然后再运行脚本。");
            return;
        }

        // 2. 检查选择
        var sel = app.selection;
        if (!sel || sel.length === 0) {
            alert("请先选中一个文本框或一个编组，再运行脚本。脚本将隐藏选中的文本框并处理。");
            return;
        }

        // 处理选中对象
        var first = sel[0];
        var frameInfo = null;
        var originalFrameId = null;

        if (first instanceof TextFrame) {
            // 文本框
            frameInfo = extractTextFrameInfo(first);
            originalFrameId = first.id;
            // 隐藏文本框
            hideItem(first);
        } else if (first instanceof Group) {
            // 处理编组: 尝试导出编组为一个矢量文件（PDF）以便 Photoshop place 为智能对象
            frameInfo = { isGroup: true };
            originalFrameId = first.id;
            // 隐藏原编组
            hideItem(first);
            // 导出编组为临时 PDF
            try {
                var exported = exportGroupToPDF(first);
                if (exported) {
                    frameInfo.exportedPath = exported;
                } else {
                    alert('导出编组为 PDF 失败，脚本将终止。');
                    return;
                }
            } catch (eExp) {
                alert('导出编组失败: ' + eExp);
                return;
            }
        } else {
            alert("所选对象不是文本框也不是编组，脚本仅支持文本框或编组。");
            return;
        }

        // 3. 检查当前页面上的链接图片
        var currentPage = doc.layoutWindows[0].activePage; // 取得活动页面（最常见的用法）
        var links = getLinksOnPage(doc, currentPage);
        if (links.length === 0) {
            alert("当前页面没有可用的链接图片。");
            return;
        }

        var chosenLink = null;
        if (links.length === 1) {
            chosenLink = links[0];
        } else {
            chosenLink = promptUserToChooseLink(links);
            if (!chosenLink) {
                // 用户取消
                return;
            }
        }

        // 4. 准备参数文件（写入系统临时目录）
        var tempFile = File(Folder.temp + "/id_to_ps_params.jsxdata");
        var params = {
            docPath: doc.filePath.fullName,
            linkPath: chosenLink.filePath, // 绝对路径字符串
            frameInfo: frameInfo,
            originalFrameId: originalFrameId,
            platform: $.os // 平台信息
        };

        // 将 params 序列化为可被 Photoshop 解析的字符串（使用 toSource 安全且 ES3 可用）
        try {
            writeStringToFile(tempFile, paramsToSource(params));
        } catch (e) {
            alert('写入临时参数文件失败: ' + e);
            return;
        }

        // 5. 构造并发送 BridgeTalk 脚本到 Photoshop
        var psWorker = buildPhotoshopWorkerScript();
        var psScriptString = psWorker.toString() + "\n" + psWorker.name + "();";

        // 通过 BridgeTalk 发送脚本，脚本会读取上面的临时参数文件
        var btResult = sendScriptToPhotoshop(psScriptString);
        if (!btResult) {
            alert("与 Photoshop 通信失败。请检查 Photoshop 是否已安装并正在运行。");
            // 尝试启动 Photoshop（平台相关）
            tryLaunchPhotoshop();
            return;
        }

        // 6. 根据 BridgeTalk 结果，更新 InDesign 中的链接（如果 Photoshop 返回了替换的路径）
        // 约定: PS 工作脚本会把结果写入另一个临时文件：id_to_ps_result.jsxdata
        var resultFile = File(Folder.temp + "/id_to_ps_result.jsxdata");
        if (resultFile.exists) {
            try {
                resultFile.open('r');
                var resultText = resultFile.read();
                resultFile.close();
                // 解析结果
                var resultObj = eval('(' + resultText + ')');
                if (resultObj && resultObj.newPath) {
                    // 更新链接: 在文档中查找原链接并替换
                    updateLinkInDocument(doc, chosenLink, resultObj.newPath);
                    alert('处理完成，已更新链接: ' + resultObj.newPath);
                } else {
                    alert('处理完成，但未返回新的链接路径。');
                }
            } catch (e) {
                // 忽略解析错误，仍然提示已完成
                alert('处理完成，但解析结果失败: ' + e);
            }
        } else {
            alert('已向 Photoshop 发送请求，但未检测到结果文件。请检查 Photoshop 是否正确执行了脚本。');
        }

    } catch (eMain) {
        alert('脚本执行发生错误: ' + eMain);
    }
}

// ==================== 帮助函数 ====================

// 检查文档是否已保存（有 filePath ）
function isDocumentSaved(doc) {
    try {
        if (!doc) return false;
        // InDesign Document 的 filePath 在未保存时不存在或为空
        if (!doc.saved || !doc.filePath) return false;
        // filePath 可能为 an alias File 对象
        return true;
    } catch (e) {
        return false;
    }
}

// 隐藏页面项（设置 visible=false 或 geometricBounds 之外的方式），优先使用 visible
function hideItem(item) {
    try {
        if (item.hasOwnProperty('visible')) {
            item.visible = false;
        } else {
            // 尝试移动到隐藏图层
            var doc = app.activeDocument;
            var hiddenLayer = null;
            // 查找名为"_hidden_by_script"的图层
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === '_hidden_by_script') {
                    hiddenLayer = doc.layers[i];
                    break;
                }
            }
            if (!hiddenLayer) hiddenLayer = doc.layers.add({name: '_hidden_by_script'});
            try { item.itemLayer = hiddenLayer; } catch (e) {}
        }
    } catch (e) {
        // 忽略隐藏失败
    }
}

// 获取文本框信息: 文本内容、每字符样式（尽量）、文本框在页面上的相对位置与旋转角
function extractTextFrameInfo(textFrame) {
    var info = {};
    try {
        // 基本检查
        if (!textFrame) return info;
        // 文本内容
        var contents = textFrame.contents;
        if (contents === undefined || contents === "") {
            alert('文本框为空或未包含文本。请填入文字后再运行。');
            throw 'EmptyText';
        }

        // 是否有溢流文本
        if (textFrame.overflows) {
            alert('文本框有溢流文本，请先处理溢流文本后再运行脚本。');
            throw 'OversetText';
        }

        info.contents = contents;

        // 尝试获取每个字符的样式（尽量），使用 characters 集合
        var chars = textFrame.parentStory.characters;
        var charsInfo = [];
        // 我们需要只获取与 textFrame 相关的 characters 的子串，先找到文本框中文本的索引范围
        // 比较保守做法: 取 textFrame.texts[0].characters
        var tfChars = textFrame.texts[0].characters;
        for (var i = 0; i < tfChars.length; i++) {
            var ch = tfChars[i];
            var chInfo = {};
            try {
                chInfo.character = ch.contents;
                // 记录字体名称和 PostScript 名称（若可用）
                try { chInfo.font = ch.appliedFont.toString(); } catch (e) { chInfo.font = null; }
                try { chInfo.postScriptName = ch.appliedFont.postscriptName; } catch (e) { chInfo.postScriptName = null; }
                chInfo.fontSize = ch.pointSize;
                chInfo.tracking = ch.tracking;
                // 填充颜色
                try { chInfo.fillColor = colorToObject(ch.fillColor); } catch (e) { chInfo.fillColor = null; }
                try { chInfo.strokeColor = colorToObject(ch.strokeColor); } catch (e) { chInfo.strokeColor = null; }
                try { chInfo.strokeWeight = ch.strokeWeight; } catch (e) { chInfo.strokeWeight = null; }
            } catch (eChar) {
                // 忽略单个字符获取失败
            }
            charsInfo.push(chInfo);
        }
        info.chars = charsInfo;

        // 文本框在页面上的相对位置（百分比）: 使用 geometricBounds 相对于页面 bounds
        var gb = textFrame.geometricBounds; // [y1, x1, y2, x2]
        var pageGB = textFrame.parentPage.bounds; // [y1, x1, y2, x2]
        var fb = {};
        try {
            var frameTop = gb[0];
            var frameLeft = gb[1];
            var frameBottom = gb[2];
            var frameRight = gb[3];
            var pageTop = pageGB[0];
            var pageLeft = pageGB[1];
            var pageBottom = pageGB[2];
            var pageRight = pageGB[3];
            var pageWidth = pageRight - pageLeft;
            var pageHeight = pageBottom - pageTop;
            if (pageWidth === 0) pageWidth = 1;
            if (pageHeight === 0) pageHeight = 1;

            fb.x = (frameLeft - pageLeft) / pageWidth; // 百分比[0-1]
            fb.y = (frameTop - pageTop) / pageHeight;
            fb.w = (frameRight - frameLeft) / pageWidth;
            fb.h = (frameBottom - frameTop) / pageHeight;
        } catch (e) { fb = {x:0,y:0,w:1,h:1}; }
        info.frameRelative = fb;

        // 旋转角度
        try {
            info.rotation = textFrame.rotationAngle;
        } catch (e) { info.rotation = 0; }

    } catch (e) {
        // 抛出以便上层处理
        throw e;
    }
    return info;
}

// 将 InDesign 颜色对象转换为可序列化的数值对象（支持 RGB/CMYK/Gray/Spot）
function colorToObject(col) {
    var out = {name: null, model: null, values: null};
    try {
        if (!col) return null;
        out.name = col.name;
        // 某些颜色可能是 Spot颜色
        try {
            if (col.space) out.model = col.space; // e.g. 'RGB', 'CMYK'
        } catch (e) {}
        // 尝试读取颜色通道值
        try {
            if (col.model === ColorModel.PROCESS) {
                // Process colors: RGB or CMYK
                try {
                    if (col.space === ColorSpace.RGB) {
                        out.model = 'RGB';
                        out.values = {r: col.colorValue[0], g: col.colorValue[1], b: col.colorValue[2]};
                    } else {
                        out.model = 'CMYK';
                        out.values = {c: col.colorValue[0], m: col.colorValue[1], y: col.colorValue[2], k: col.colorValue[3]};
                    }
                } catch (e) { out.values = null; }
            } else if (col.model === ColorModel.SPOT) {
                out.model = 'SPOT';
                try { out.values = {spotName: col.spot.name}; } catch (e) { out.values = null; }
            } else {
                out.values = null;
            }
        } catch (e) { out.values = null; }
    } catch (e) { return null; }
    return out;
}

// 将选中编组导出为临时 PDF，返回导出文件路径（字符串）或 null
function exportGroupToPDF(groupObj) {
    try {
        var tmpFolder = Folder.temp;
        var baseName = 'id_group_export_' + new Date().getTime();
        var pdfFile = File(tmpFolder + '/' + baseName + '.pdf');

        // 创建一个临时文档，将编组复制进去导出，然后删除临时文档
        var doc = app.activeDocument;
        // 复制组到一个新的临时文档: 新建文档并粘贴
        var tmpDoc = app.documents.add();
        // 复制组到剪贴板并粘贴
        try {
            app.copy();
        } catch (eCopy) {
            // 如果 app.copy 不可用，尝试使用 selection.copy
            try { groupObj.copy(); } catch (e2) {}
        }
        // 激活临时文档并粘贴
        app.activeDocument = tmpDoc;
        try { tmpDoc.pages[0].paste(); } catch (ePaste) {}

        // 导出页面为 PDF
        var pdfExpPreset = app.pdfExportPresets.itemByName('High Quality Print');
        var usePreset = null;
        try {
            if (pdfExpPreset && pdfExpPreset.isValid) usePreset = pdfExpPreset;
        } catch (e) { usePreset = null; }
        if (!usePreset) {
            // 尝试使用默认
            try { usePreset = app.pdfExportPresets[0]; } catch (e) { usePreset = null; }
        }
        // 构建导出范围为当前页面
        try {
            tmpDoc.exportFile(ExportFormat.pdfType, pdfFile, false, usePreset);
        } catch (eExport) {
            // 若上面失败，尝试 simpler: print export
            try { tmpDoc.exportFile(ExportFormat.pdfType, pdfFile); } catch (e2) { throw e2; }
        }

        // 关闭临时文档不保存
        try { tmpDoc.close(SaveOptions.NO); } catch (e) {}
        // 还原原文档为活动文档
        app.activeDocument = doc;
        return pdfFile.fullName;
    } catch (e) {
        try { if (tmpDoc) tmpDoc.close(SaveOptions.NO); } catch (er) {}
        throw e;
    }
}

// 获取页面上的链接图片（返回 Link 对象数组，包含 filePath 字符串）
function getLinksOnPage(doc, page) {
    var out = [];
    try {
        var allLinks = doc.links;
        for (var i = 0; i < allLinks.length; i++) {
            try {
                var lk = allLinks[i];
                // Link 对象有 parent 或 linkResource
                var parentPage = null;
                try {
                    var parent = lk.parent;
                    if (parent && parent.parentPage) parentPage = parent.parentPage;
                } catch (e) {}
                // 一些链接的 parentPage 可能为 null，如果 parentPage 等于请求 page 则加入
                if (parentPage && parentPage === page) {
                    out.push({name: lk.name, filePath: lk.filePath, linkObj: lk});
                }
            } catch (e) {}
        }
    } catch (e) {}
    return out;
}

// 弹窗让用户选择要处理的链接（返回选中的 link 对象）
function promptUserToChooseLink(links) {
    try {
        // 使用 ScriptUI 构建简单列表
        var w = new Window('dialog', '选择链接图片');
        var list = w.add('listbox', undefined, [], {multiselect:false});
        list.preferredSize = [400, 200];
        for (var i = 0; i < links.length; i++) {
            list.add('item', links[i].name + '  —  ' + links[i].filePath);
        }
        var btnGroup = w.add('group');
        btnGroup.orientation = 'row';
        var okBtn = btnGroup.add('button', undefined, '确定');
        var cancelBtn = btnGroup.add('button', undefined, '取消');
        if (w.show() === 1) {
            var selIndex = list.selection ? list.selection.index : -1;
            if (selIndex >= 0) return links[selIndex];
        }
    } catch (e) {}
    return null;
}

// 将对象序列化为 toSource 风格字符串（兼容 ExtendScript）
function paramsToSource(obj) {
    try {
        // ExtendScript 提供 toSource；但为了安全，简单实现：
        if (obj && obj.toSource) return obj.toSource();
        // Fallback: 手工构建（只支持基本类型）
        return manualSerialize(obj);
    } catch (e) {
        return manualSerialize(obj);
    }
}

function manualSerialize(v) {
    if (v === null) return 'null';
    var t = typeof v;
    if (t === 'number' || t === 'boolean') return String(v);
    if (t === 'string') return '"' + v.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
    if (v instanceof Array) {
        var parts = [];
        for (var i = 0; i < v.length; i++) parts.push(manualSerialize(v[i]));
        return '[' + parts.join(',') + ']';
    }
    // object
    var parts2 = [];
    for (var k in v) {
        if (v.hasOwnProperty(k)) {
            parts2.push('"' + k + '":' + manualSerialize(v[k]));
        }
    }
    return '{' + parts2.join(',') + '}';
}

// 将字符串写入文件（覆盖）
function writeStringToFile(fileObj, str) {
    try {
        fileObj.encoding = 'UTF-8';
        fileObj.open('w');
        fileObj.write(str);
        fileObj.close();
    } catch (e) {
        throw e;
    }
}

// 构建发送到 Photoshop 的 worker 函数（被 toString 并发送）
function buildPhotoshopWorkerScript() {
    function psWorker() {
        // Photoshop 侧: 读取临时参数文件，打开图片，尝试将图片转为灰度（若为位图），将文字按参数重建到图层中，保存覆盖或另存为 PSD
        try {
            var tempFile = File(Folder.temp + '/id_to_ps_params.jsxdata');
            if (!tempFile.exists) {
                throw 'Param file not found: ' + tempFile.fsName;
            }
            tempFile.open('r');
            var txt = tempFile.read();
            tempFile.close();
            // 使用 eval 解析（在 ExtendScript 中通常可用）
            var params = eval('(' + txt + ')');

            // 打开背景图片（源链接）
            var imgFile = File(params.linkPath);
            if (!imgFile.exists) {
                throw 'Image file not found: ' + params.linkPath;
            }
            var docRef = null;
            try {
                docRef = app.open(imgFile);
            } catch (eOpen) {
                // 无法直接打开某些格式，尝试抛出错误
                throw eOpen;
            }

            // 如果存在导出的矢量文件（来自 InDesign 编组导出），将其以智能对象方式放置到当前打开的背景文档中
            try {
                if (params.frameInfo && params.frameInfo.exportedPath) {
                    var placeFile = File(params.frameInfo.exportedPath);
                    if (placeFile.exists) {
                        // 使用 ActionManager 的 Place 命令将文件作为智能对象放入当前文档
                        var idPlc = charIDToTypeID('Plc ');
                        var descPlc = new ActionDescriptor();
                        descPlc.putPath(charIDToTypeID('null'), placeFile);
                        // 抑制用户交互
                        executeAction(idPlc, descPlc, DialogModes.NO);
                        // 放置后，placed layer 将作为当前图层存在，继续执行
                    }
                }
            } catch (ePlace) {
                // 放置失败则继续，不影响后续文本绘制
            }

            // 如果模式是 Bitmap，则转换为 Grayscale
            try {
                if (docRef.mode === DocumentMode.BITMAP) {
                    docRef.changeMode(ChangeMode.GRAYSCALE);
                }
            } catch (eMode) {
                // 忽略不能转换的错误
            }

            // 现在根据 params.frameInfo 尝试创建文本图层
            var fi = params.frameInfo || {};
            var contents = '';
            if (fi.contents) contents = fi.contents;

            // 新建文本图层（Point 文本，置于中心位置的近似位置）
            var textLayer = docRef.artLayers.add();
            textLayer.kind = LayerKind.TEXT;
            var ti = textLayer.textItem;
            try { ti.contents = contents; } catch (e) { ti.contents = contents + ''; }

            // 优先尝试整体设置字体、大小、tracking、颜色（若提供第一个字符信息）
            if (fi.chars && fi.chars.length > 0) {
                var firstChar = fi.chars[0];
                try {
                    // 优先使用 PostScript 名称
                    if (firstChar.postScriptName && firstChar.postScriptName !== null) ti.font = firstChar.postScriptName;
                    else if (firstChar.font) ti.font = firstChar.font.toString();
                } catch (e) {}
                try { if (firstChar.fontSize) ti.size = firstChar.fontSize; } catch (e) {}
                try { if (firstChar.tracking) ti.tracking = firstChar.tracking; } catch (e) {}
                // 颜色映射: 根据捕获的颜色模型尽量转换为 SolidColor
                try {
                    if (firstChar.fillColor && firstChar.fillColor.values) {
                        var c = new SolidColor();
                        var fc = firstChar.fillColor;
                        if (fc.model === 'RGB' && fc.values) {
                            c.rgb.red = fc.values.r;
                            c.rgb.green = fc.values.g;
                            c.rgb.blue = fc.values.b;
                        } else if (fc.model === 'CMYK' && fc.values) {
                            // 简单 CMYK -> RGB 近似转换（0-100 假定）
                            var cC = fc.values.c / 100;
                            var cM = fc.values.m / 100;
                            var cY = fc.values.y / 100;
                            var cK = fc.values.k / 100;
                            var r = 255 * (1 - Math.min(1, cC * (1 - cK) + cK));
                            var g = 255 * (1 - Math.min(1, cM * (1 - cK) + cK));
                            var b = 255 * (1 - Math.min(1, cY * (1 - cK) + cK));
                            c.rgb.red = Math.round(r);
                            c.rgb.green = Math.round(g);
                            c.rgb.blue = Math.round(b);
                        } else {
                            // 兜底黑色
                            c.rgb.red = 0; c.rgb.green = 0; c.rgb.blue = 0;
                        }
                        ti.color = c;
                    }
                } catch (e) {}
            }

            // 尽量按字符设置样式：创建多个小文本图层进行模拟（可能很慢），如果字符数量大于 100 则跳过逐字符操作
            try {
                if (fi.chars && fi.chars.length > 0 && fi.chars.length <= 100) {
                    // 使用宽度估算与 tracking 模拟字符间距
                    // 先删除上面整体图层，改为逐字符单独图层
                    try { textLayer.remove(); } catch (e) {}
                    var startX = docRef.width.as('px') / 2; // 简单居中
                    var startY = docRef.height.as('px') / 2;
                    var xPos = startX;
                    for (var i = 0; i < fi.chars.length; i++) {
                        var ch = fi.chars[i];
                        var tl = docRef.artLayers.add();
                        tl.kind = LayerKind.TEXT;
                        var titem = tl.textItem;
                        titem.contents = ch.character || '';
                        try { if (ch.postScriptName) titem.font = ch.postScriptName; else if (ch.font) titem.font = ch.font; } catch (e) {}
                        try { if (ch.fontSize) titem.size = ch.fontSize; } catch (e) {}
                        try { titem.position = [xPos, startY]; } catch (e) {}
                        // 估算宽度，advance by font size/2 + tracking
                        var adv = (ch.fontSize || 12) * 0.6;
                        if (ch.tracking) adv += ch.tracking/20;
                        xPos += adv;
                    }
                }
            } catch (eCharLayer) {
                // 如果逐字符失败，忽略并保留整体图层
            }

            // 决定保存: 如果原文件是 PSD 或 TIF 则覆盖保存，否则另存为 PSD
            var origName = imgFile.name;
            var lower = origName.toLowerCase();
            var saveAsPSD = true;
            if (lower.indexOf('.psd') > -1 || lower.indexOf('.tif') > -1 || lower.indexOf('.tiff') > -1) {
                // 尝试覆盖保存原文件
                try {
                    var saveFile = imgFile;
                    var psdSaveOptions = new PhotoshopSaveOptions();
                    // 保存为 PSD 或覆盖 TIFF：对 TIFF 需要 TIFF 保存选项，简化为 PSD 覆盖
                    docRef.save();
                    saveAsPSD = false;
                } catch (eSave) {
                    saveAsPSD = true;
                }
            }

            var newPath = imgFile.fullName;
            if (saveAsPSD) {
                // 另存为同名 PSD
                var psdFile = File(imgFile.path + '/' + imgFile.name.replace(/\.[^\.]+$/, '') + '.psd');
                var psdSaveOptions = new PhotoshopSaveOptions();
                psdSaveOptions.layers = true;
                docRef.saveAs(psdFile, psdSaveOptions, true);
                newPath = psdFile.fullName;
            }

            // 关闭文档
            try { docRef.close(SaveOptions.DONOTSAVECHANGES); } catch (e) {}

            // 将结果写回结果文件
            var resultFile = File(Folder.temp + '/id_to_ps_result.jsxdata');
            resultFile.open('w');
            resultFile.write('{"newPath": "' + newPath.replace(/\\/g, '\\\\') + '"}');
            resultFile.close();

        } catch (e) {
            // 写入失败结果并抛出
            try {
                var rf = File(Folder.temp + '/id_to_ps_result.jsxdata');
                rf.open('w');
                rf.write('{"error": "' + String(e).replace(/\\/g, '\\\\') + '"}');
                rf.close();
            } catch (er) {}
            throw e;
        }
    }
    return psWorker;
}

// 通过 BridgeTalk 发送脚本到 Photoshop 并等待返回
function sendScriptToPhotoshop(scriptString) {
    try {
        if (BridgeTalk === undefined) {
            return false;
        }
    } catch (e) {
        return false;
    }

    try {
        var bt = new BridgeTalk();
        bt.target = 'photoshop';
        bt.body = scriptString;
        var finished = false;
        var success = false;

        bt.onError = function (inBT) {
            finished = true; success = false;
        };
        bt.onResult = function (inBT) {
            finished = true; success = true;
        };
        bt.onTimeout = function (inBT) {
            finished = true; success = false;
        };
        bt.timeout = 120000; // 120s
        bt.send();

        // 等待完成（同步等待）
        var start = new Date().getTime();
        while (!finished) {
            // 处理事件队列
            $.sleep(200);
            // 超时二次保险
            if ((new Date().getTime() - start) > 120000) break;
        }
        return success;
    } catch (e) {
        return false;
    }
}

// 更新 InDesign 中文档中的链接为新的路径
function updateLinkInDocument(doc, chosenLink, newPath) {
    try {
        if (!newPath) return;
        // 通过 Link 对象 relink
        var linkObj = chosenLink.linkObj;
        if (linkObj && linkObj.relink) {
            try {
                linkObj.relink(File(newPath));
                linkObj.update();
            } catch (e) {
                // 如果失败，尝试通过 doc.links 查找匹配名称并替换
                for (var i = 0; i < doc.links.length; i++) {
                    try {
                        var l = doc.links[i];
                        if (l.name === chosenLink.name) {
                            l.relink(File(newPath));
                            l.update();
                        }
                    } catch (e2) {}
                }
            }
        }
    } catch (e) {}
}

// 尝试启动 Photoshop（基于平台）
function tryLaunchPhotoshop() {
    try {
        var os = $.os.toLowerCase();
        if (os.indexOf('mac') !== -1) {
            // macOS: 尝试通过 app.launchApp
            try { app.launchApplication('Adobe Photoshop'); } catch (e) {}
        } else {
            // Windows: 尝试通过系统命令打开（尽量）
            // 注意：ExtendScript 中没有直接的 Shell 执行，尝试使用 File.execute
            var psExe = File('C:/Program Files/Adobe/Adobe Photoshop 2025/Adobe Photoshop.exe');
            try { if (psExe.exists) psExe.execute(); } catch (e) {}
        }
    } catch (e) {}
}

// 运行主函数
try { main(); } catch (e) { alert('致命错误: ' + e); }
