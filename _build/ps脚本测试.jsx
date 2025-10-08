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
            // 如果模式是 Bitmap，则转换为 Grayscale
            try {
                if (docRef.mode === DocumentMode.BITMAP) {
                    docRef.changeMode(ChangeMode.GRAYSCALE);
                }
            } catch (eMode) {
                // 忽略不能转换的错误
            }
            // 如果存在导出的矢量文件（来自 InDesign 编组导出），优先将其作为 PDF 打开并把所有图层复制到当前打开的背景文档中
            try {
                if (params.frameInfo && params.frameInfo.isGroup === true && params.frameInfo.exportedPath) {
                    var placeFile = File(params.frameInfo.exportedPath);
                    if (placeFile.exists) {
                        try {
                            // 打开 PDF 时尽量使用出血裁剪并选择第 1 页、CMYK 模式
                            var pdfOpts = new PDFOpenOptions();
                            try { pdfOpts.page = 1; } catch (e) {}
                            // 尝试使用出血裁剪（不同 PS 版本枚举名可能不同），失败则忽略
                            try { pdfOpts.crop = CropTo.BLEEDBOX; } catch (e) { try { pdfOpts.crop = CropTo.BLEED; } catch (e2) {} }
                            try { pdfOpts.mode = OpenDocumentMode.CMYK; } catch (e) {}
                            // 打开 PDF
                            var pdfDoc = app.open(placeFile, pdfOpts);
                            // 尽量设置为 8 位
                            try { pdfDoc.bitsPerChannel = BitsPerChannelType.EIGHT; } catch (e) {}
                            // 将 PDF 文档的所有顶级图层复制到目标图片文档（docRef）中
                            try {
                                // 使用从上到下复制，保留图层结构
                                for (var li = 0; li < pdfDoc.layers.length; li++) {
                                    try { pdfDoc.layers[li].duplicate(docRef, ElementPlacement.PLACEATEND); } catch (eDup) {}
                                }
                            } catch (eDupAll) {}
                            // 关闭 PDF，不保存任何修改
                            try { pdfDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e) {}
                        } catch (eOpenPdf) {
                            // 如果直接以 PDFOpenOptions 打开失败，回退到原来的 place 方法（smart object）
                            try {
                                var idPlc = charIDToTypeID('Plc ');
                                var descPlc = new ActionDescriptor();
                                descPlc.putPath(charIDToTypeID('null'), placeFile);
                                executeAction(idPlc, descPlc, DialogModes.NO);
                            } catch (ePlace) {}
                        }
                    }
                }
            } catch (ePlace) {
                // 放置/复制失败则继续，不影响后续文本绘制
            }

            // 现在根据 params.frameInfo 尝试创建文本图层，若为编组导出（isGroup === true）则已复制 PDF 图层，不再新建文本图层
            var fi = params.frameInfo || {};
            var contents = '';
            if (fi.contents) contents = fi.contents;

            var textLayer = null;
            var ti = null;
            if (!(fi.isGroup === true)) {
                // 新建文本图层（Point 文本）并尝试使用 page-percentage 将其放置到画布对应位置
                try {
                    textLayer = docRef.artLayers.add();
                    textLayer.kind = LayerKind.TEXT;
                    ti = textLayer.textItem;
                    try { ti.contents = contents; } catch (e) { ti.contents = contents + ''; }
                } catch (eNew) {
                    textLayer = null; ti = null;
                }
            } else {
                // 使用 PDF 导入的图层作为视觉文本层，保持 docRef.activeLayer 为最后一个复制的图层
                try { /* nothing to create */ } catch (e) {}
            }

            // 计算定位与缩放参数
            var pageBounds = params.pageBounds; // 来自 InDesign: [y1,x1,y2,x2]
            var placedBounds = params.placedBounds; // 来自 InDesign 放置框
            var frameRel = (fi.frameRelative) ? fi.frameRelative : null; // {x,y,w,h}
            var docWidthPx = docRef.width.as('mm');
            var docHeightPx = docRef.height.as('mm');

            // 计算目标像素位置的辅助函数：优先在 placedBounds 内计算，否则按整张画布
            function computePixelPosFromPagePercent(rel) {
                // rel: {x,y,w,h} 相对于页面
                try {
                    if (rel && pageBounds && placedBounds) {
                        var pageTop = pageBounds[0]; var pageLeft = pageBounds[1]; var pageBottom = pageBounds[2]; var pageRight = pageBounds[3];
                        var pageW = pageRight - pageLeft; if (pageW === 0) pageW = 1;
                        var pageH = pageBottom - pageTop; if (pageH === 0) pageH = 1;
                        // 计算 frame 在页面上的绝对位置（点）：frameLeft = pageLeft + rel.x * pageW
                        var frameLeftPts = pageLeft + rel.x * pageW;
                        var frameTopPts = pageTop + rel.y * pageH;
                        // placedBounds: [y1,x1,y2,x2]
                        var placedTop = placedBounds[0]; var placedLeft = placedBounds[1]; var placedBottom = placedBounds[2]; var placedRight = placedBounds[3];
                        var placedW = placedRight - placedLeft; if (placedW === 0) placedW = 1;
                        var placedH = placedBottom - placedTop; if (placedH === 0) placedH = 1;
                        // 计算相对于 placed 的比例
                                    try {
                                        // 读取 InDesign 传来的缩放信息，计算与文本相同的放大因子
                                        var psLinkScale = params.linkScale || {h:100, v:100};
                                        var psAvgScale = ((psLinkScale.h || 100) + (psLinkScale.v || 100)) / 2;
                                        var psScaleFactor = 1;
                                        if (psAvgScale && psAvgScale !== 0) psScaleFactor = 100 / psAvgScale;
                                        var resizePercent = psScaleFactor * 100; // resize() 使用百分比值
                                        // 获取当前活动图层并尝试缩放（智能对象层可缩放）
                                        try {
                                            var placedLayer = docRef.activeLayer;
                                            if (placedLayer && placedLayer.resize) {
                                                placedLayer.resize(resizePercent, resizePercent, AnchorPosition.MIDDLECENTER);
                                            }
                                        } catch (eResize) {
                                            // 若 ActionManager 需要更复杂的变换，可在后续迭代中实现
                                        }
                                    } catch (eScale) {}
                        var rx = (frameLeftPts - placedLeft) / placedW;
                        var ry = (frameTopPts - placedTop) / placedH;
                        // 限制在 [0,1]
                        if (rx < 0) rx = 0; if (rx > 1) rx = 1;
                        if (ry < 0) ry = 0; if (ry > 1) ry = 1;
                        // 转为像素
                        var xPx = rx * docWidthPx;
                        var yPx = ry * docHeightPx;
                        return [xPx, yPx];
                    }
                } catch (e) {}
                // 兜底：按整个页面百分比映射到画布
                try {
                    if (rel) {
                        var xPx2 = rel.x * docWidthPx;
                        var yPx2 = rel.y * docHeightPx;
                        return [xPx2, yPx2];
                    }
                } catch (e) {}
                // 最后兜底居中
                return [docWidthPx/2, docHeightPx/2];
            }

            // 优先从第一个字符获得样式用于整体设置
            var baseFontSize = null;
            var baseFontName = null;
            var baseTracking = null;
            var baseColor = null;
            if (fi.chars && fi.chars.length > 0) {
                var firstChar = fi.chars[0];
                try { if (firstChar.postScriptName) baseFontName = firstChar.postScriptName; else if (firstChar.font) baseFontName = firstChar.font; } catch (e) {}
                try { if (firstChar.fontSize) baseFontSize = firstChar.fontSize; } catch (e) {}
                try { if (firstChar.tracking) baseTracking = firstChar.tracking; } catch (e) {}
                try { if (firstChar.fillColor) baseColor = firstChar.fillColor; } catch (e) {}
            }

            // 考虑 InDesign 中链接的缩放比例：在 ID 中放置为 60% 时，我们需要把字体乘以 (100/60) 才能在原图上显示同样视觉大小
            var linkScale = params.linkScale || {h:100,v:100};
            var avgScale = ( (linkScale.h||100) + (linkScale.v||100) ) / 2;
            var scaleFactor = 1;
            if (avgScale && avgScale !== 0) scaleFactor = 100 / avgScale;

            // 应用整体样式（仅当我们创建了文本图层时）
            if (ti) {
                try {
                    if (baseFontName) ti.font = baseFontName;
                    ti.autoLeadingAmount = 120; // 自动行距
                } catch (e) {}
                try {
                    if (baseFontSize) ti.size = baseFontSize * scaleFactor;
                } catch (e) {}
                try { if (baseTracking) ti.tracking = baseTracking; } catch (e) {}
            }
            // 颜色
            try {
                if (baseColor && baseColor.values) {
                    var c = new SolidColor();
                    var fc = baseColor;
                    if (fc.model === 'RGB' && fc.values) {
                        c.rgb.red = fc.values.r;
                        c.rgb.green = fc.values.g;
                        c.rgb.blue = fc.values.b;
                    } else if (fc.model === 'CMYK' && fc.values) {
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
                    } else { c.rgb.red = 0; c.rgb.green = 0; c.rgb.blue = 0; }
                    ti.color = c;
                }
            } catch (e) {}

            // 计算并应用位置（仅当我们创建了文本图层时）
            var pos = null;
            try {
                pos = computePixelPosFromPagePercent(frameRel);
                // Photoshop 文本位置通常以像素或文档单位给出，确保为数字数组
                if (ti && pos && pos.length === 2) {
                    ti.position = pos;
                }
            } catch (e) {}

            // 尽量按字符设置样式：创建多个小文本图层进行模拟（可能很慢），如果字符数量大于 6 则跳过逐字符操作
            try {
                // 仅在我们创建了文本图层（非 isGroup）时才尝试逐字符创建
                if (ti && fi.chars && fi.chars.length > 0 && fi.chars.length <= 6) {
                    // 使用宽度估算与 tracking 模拟字符间距
                    // 先删除上面整体图层，改为逐字符单独图层
                    try { if (textLayer) textLayer.remove(); } catch (e) {}
                    var startX = (pos && pos.length === 2) ? pos[0] : (docRef.width.as('px')/2);
                    var startY = (pos && pos.length === 2) ? pos[1] : (docRef.height.as('px')/2);
                    // 默认沿 y 方向排列（竖排），如果 orientation === 'horizontal' 则沿 x 方向排列
                    var orient = (fi.orientation && fi.orientation === 'vertical') ? 'vertical' : 'horizontal';
                    var cursorX = startX;
                    var cursorY = startY;
                    for (var i = 0; i < fi.chars.length; i++) {
                        var ch = fi.chars[i];
                        var tl = docRef.artLayers.add();
                        tl.kind = LayerKind.TEXT;
                        var titem = tl.textItem;
                        titem.contents = ch.character || '';
                        try { if (ch.postScriptName) titem.font = ch.postScriptName; else if (ch.font) titem.font = ch.font; } catch (e) {}
                        try { if (ch.fontSize) titem.size = ch.fontSize * scaleFactor; } catch (e) {}
                        try { titem.position = [cursorX, cursorY]; } catch (e) {}
                        // 估算 advance：基于字体大小和 tracking
                        var adv = (ch.fontSize || 12) * 0.6;
                        if (ch.tracking) adv += ch.tracking/20;
                        if (orient === 'vertical') {
                            // 竖排：每个字沿 Y 增加
                            cursorY += adv;
                        } else {
                            // 横排：每个字沿 X 增加（向右）
                            cursorX += adv;
                        }
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
            resultFile.encoding = 'UTF-8';
            resultFile.open('w');
            resultFile.write('{"newPath": "' + newPath.replace(/\\/g, '\\\\') + '"}');
            resultFile.close();

        } catch (e) {
            // 写入失败结果并抛出
            try {
                var rf = File(Folder.temp + '/id_to_ps_result.jsxdata');
                rf.encoding = 'UTF-8';
                rf.open('w');
                rf.write('{"error": "' + String(e).replace(/\\/g, '\\\\') + '"}');
                rf.close();
            } catch (er) {}
            throw e;
        }
    }

psWorker();