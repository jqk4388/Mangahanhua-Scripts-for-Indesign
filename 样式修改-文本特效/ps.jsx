
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
            // 使用 eval 解析
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
            resultFile.encoding = 'UTF-8'
            resultFile.open('w');
            resultFile.write('{"newPath": "' + newPath.replace(/\\/g, '\\\\') + '"}');
            resultFile.close();

        } catch (e) {
            // 写入失败结果并抛出
            try {
                var rf = File(Folder.temp + '/id_to_ps_result.jsxdata');
                rf.encoding = 'UTF-8'
                rf.open('w');
                rf.write('{"error": "' + String(e).replace(/\\/g, '\\\\') + '"}');
                rf.close();
            } catch (er) {}
            throw e;
        }
    }

psWorker();