// #target "InDesign"

if (app.documents.length === 0) {
    alert("请先打开一个文档。");
    exit();
}

var doc = app.activeDocument;
    // 设置标尺原点和单位
    doc.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;
    doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
    doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
    doc.zeroPoint = [0, 0];
    //文本大小的单位改成点
    doc.viewPreferences.textSizeMeasurementUnits = MeasurementUnits.POINTS;
    doc.viewPreferences.typographicMeasurementUnits = MeasurementUnits.POINTS;

// 主函数
function main() {
    var outputFile = getOutputFilePath();
    var file = new File(outputFile);

    // 打开文件并设置为 UTF-8 格式
    file.encoding = "UTF-8";
    file.open("w");
    writeHeader(file);

    for (var p = 0; p < doc.pages.length; p++) {
        var page = doc.pages[p];
        processPage(page, file);
    }

    file.close();
    alert("文本已导出到桌面: " + outputFile);
}

// 获取导出文件路径
function getOutputFilePath() {
    var docName = doc.name.replace(/\.[^\.]+$/, ""); // 去掉扩展名
    return Folder.desktop + "/" + docName + "_校对用lptxt.txt";
}

// 写入文件头部信息
function writeHeader(file) {
    file.write("1,0\n-\n框内\n框外\n-\n备注备注备注\n");
}

// 将 kenten（着重号）插入应用到原始文本：在被加着重号的连续序列前后插入方括号
function applyKentenInsertions(textObj, text) {
    try {
        var chars = textObj.characters;
        var insertions = []; // { pos: index, text: string, priority: number }
        var searchStart = 0;
        for (var ci = 0; ci < chars.length; ci++) {
            var ch = chars[ci];
            var kentenKind = (ch && typeof ch.kentenKind !== 'undefined') ? String(ch.kentenKind) : "NONE";
            if (ch && kentenKind !== "NONE") {
                var k_end = ci;
                while (k_end + 1 < chars.length && chars[k_end + 1] && String(typeof chars[k_end + 1].kentenKind !== 'undefined' ? chars[k_end + 1].kentenKind : "NONE") !== "NONE") {
                    k_end++;
                }
                var k_seq = "";
                for (var kk = ci; kk <= k_end; kk++) {
                    k_seq += String(chars[kk].contents);
                }
                var k_found = -1;
                if (k_seq !== "") {
                    k_found = text.indexOf(k_seq, searchStart);
                }
                if (k_found === -1) {
                    var fallbackPos2 = searchStart;
                    for (var kk = ci; kk <= k_end; kk++) {
                        var single2 = String(chars[kk].contents);
                        var f3 = text.indexOf(single2, fallbackPos2);
                        if (f3 === -1) break;
                        fallbackPos2 = f3 + single2.length;
                    }
                    k_found = fallbackPos2 - k_seq.length;
                }
                var k_start_pos = (k_found !== -1) ? k_found : searchStart;
                var k_end_pos = k_start_pos + k_seq.length;
                insertions.push({ pos: k_start_pos, text: "【", priority: 0 });
                insertions.push({ pos: k_end_pos, text: "】", priority: 1 });
                searchStart = k_end_pos;
                ci = k_end;
            } else {
                var simple = ch ? String(ch.contents) : "";
                if (simple !== "") {
                    var f2 = text.indexOf(simple, searchStart);
                    if (f2 !== -1) searchStart = f2 + simple.length;
                }
            }
        }

        if (insertions.length > 0) {
            insertions.sort(function(a, b) { if (a.pos !== b.pos) return a.pos - b.pos; return a.priority - b.priority; });
            var finalText = text;
            var added = 0;
            for (var ii = 0; ii < insertions.length; ii++) {
                var it = insertions[ii];
                var ins = it.text;
                var pos = it.pos + added;
                if (pos < 0) pos = 0;
                if (pos > finalText.length) pos = finalText.length;
                finalText = finalText.slice(0, pos) + ins + finalText.slice(pos);
                added += ins.length;
            }
            return finalText;
        }
    } catch (e) {
        return text;
    }
    return text;
}

// 将 ruby（拼音）插入应用到文本：在连续带相同 rubyString 的序列末尾插入括号并放入拼音
function applyRubyInsertions(textObj, text) {
    try {
        var chars = textObj.characters;
        var insertions = [];
        var searchStart = 0;
        for (var ci = 0; ci < chars.length; ci++) {
            var ch = chars[ci];
            if (ch && ch.rubyFlag) {
                var ruby = ch.rubyString || "";
                var end = ci;
                while (end + 1 < chars.length && chars[end + 1] && chars[end + 1].rubyFlag && String(chars[end + 1].rubyString) === String(ruby)) {
                    end++;
                }
                var seqStr = "";
                for (var k = ci; k <= end; k++) seqStr += String(chars[k].contents);
                var found = -1;
                if (seqStr !== "") found = text.indexOf(seqStr, searchStart);
                if (found === -1) {
                    var fallbackPos = searchStart;
                    for (var k = ci; k <= end; k++) {
                        var single = String(chars[k].contents);
                        var f = text.indexOf(single, fallbackPos);
                        if (f === -1) break;
                        fallbackPos = f + single.length;
                    }
                    found = fallbackPos - seqStr.length;
                }
                var insertPos = (found !== -1) ? (found + seqStr.length) : searchStart;
                if (ruby !== "") insertions.push({ pos: insertPos, text: "（" + ruby + "）", priority: 0 });
                searchStart = insertPos;
                ci = end;
            } else {
                var simple = ch ? String(ch.contents) : "";
                if (simple !== "") {
                    var f2 = text.indexOf(simple, searchStart);
                    if (f2 !== -1) searchStart = f2 + simple.length;
                }
            }
        }

        if (insertions.length > 0) {
            insertions.sort(function(a, b) { return a.pos - b.pos; });
            var finalText = text;
            var added = 0;
            for (var ii = 0; ii < insertions.length; ii++) {
                var it = insertions[ii];
                var ins = it.text;
                var pos = it.pos + added;
                if (pos < 0) pos = 0;
                if (pos > finalText.length) pos = finalText.length;
                finalText = finalText.slice(0, pos) + ins + finalText.slice(pos);
                added += ins.length;
            }
            return finalText;
        }
    } catch (e) {
        return text;
    }
    return text;
}

// 处理单页
function processPage(page, file) {
    var pageNumber = page.name; // 页码
    var textFrames = getAllTextFrames(page);

    // 按从上到下、从右到左排序
    textFrames.sort(function (a, b) {
        var aBounds = a.geometricBounds;
        var bBounds = b.geometricBounds;
        if (aBounds[0] !== bBounds[0]) {
            return aBounds[0] - bBounds[0]; // 按顶部坐标排序
        }
        return bBounds[1] - aBounds[1]; // 按左侧坐标排序（从右到左）
    });

    file.writeln(">>>>>>>>[" + doc.name.replace(/\.[^\.]+$/, "") + "_" + pageNumber + ".jpg]<<<<<<<<" + "\n");

    var layerGroups = {}; // 图层分组
    var textIndex = 1; // 文本序号

    for (var i = 0; i < textFrames.length; i++) {
        var textFrame = textFrames[i];

        // 跳过隐藏的文本框
        if (!textFrame.visible) continue;
        //跳过在隐藏的图层的文本框
        if (textFrame.itemLayer.visible === false) continue;

        var textObj = textFrame.texts[0];
        var text = textObj.contents;
        if (!text || String(text).replace(/^\s+|\s+$/g, '') === "") continue; // 跳过空文本框

        // 先应用着重号插入，再应用拼音插入（两个独立函数）
        var processedText = text;
        processedText = applyKentenInsertions(textObj, processedText);
        processedText = applyRubyInsertions(textObj, processedText);

        var geometricBounds = textFrame.geometricBounds; // [上, 左, 下, 右]
        var x = geometricBounds[1] / doc.documentPreferences.pageWidth;
        var y = geometricBounds[0] / doc.documentPreferences.pageHeight;

        // 获取图层分组序号
        var layerName = textFrame.itemLayer.name;
        if (!layerGroups[layerName]) {
            var count = 0;
            for (var key in layerGroups) {
                if (layerGroups.hasOwnProperty(key)) {
                    count++;
                }
            }
            layerGroups[layerName] = count + 1;
        }
        var groupIndex = layerGroups[layerName];

        // 获取字体和字号信息
        var fontName = textFrame.texts[0].appliedFont.name; // 获取字体名称
        var fontSize = textFrame.texts[0].pointSize; // 获取字号

        // 写入文本信息
        file.writeln("----------------[" + textIndex + "]----------------[" + x.toFixed(5) + "," + y.toFixed(5) + "," + groupIndex + "]");
        file.writeln("{字体：" + fontName + "}{字号：" + fontSize + "}" + processedText);

        textIndex++;
    }
}

// 获取页面中所有文本框，包括编组中的文本框
function getAllTextFrames(page) {
    var textFrames = [];

    function collectTextFrames(item) {
        if (item.constructor.name === "TextFrame") {
            textFrames.push(item);
        } else if (item.constructor.name === "Group") {
            for (var i = 0; i < item.allPageItems.length; i++) {
                collectTextFrames(item.allPageItems[i]);
            }
        }
    }

    for (var i = 0; i < page.allPageItems.length; i++) {
        collectTextFrames(page.allPageItems[i]);
    }

    // 过滤掉文字转曲后的路径
    var validFrames = [];
    for (var i = 0; i < textFrames.length; i++) {
        if (textFrames[i].contents !== null && String(textFrames[i].contents).replace(/^\s+|\s+$/g, '') !== "") {
            validFrames.push(textFrames[i]);
        }
    }
    return validFrames;
}

// 执行主函数
main();