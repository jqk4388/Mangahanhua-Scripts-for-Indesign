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

        var text = textFrame['parentStory']['contents'];
        if (!text || String(text).replace(/^\s+|\s+$/g, '') === "") continue; // 跳过空文本框

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
        file.writeln("{字体：" + fontName + "}{字号：" + fontSize + "}"+text);

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