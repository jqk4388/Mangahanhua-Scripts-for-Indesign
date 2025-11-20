// 主入口函数
function main() {
    try {
        var activeDoc = getActiveDocument();
        var sourceFile = selectInddFile();
        if (!sourceFile) {
            alert("未选择任何文件。");
            return;
        }
        var sourceDoc = openSourceDocument(sourceFile);
        if (!sourceDoc) {
            alert("无法打开源文档。");
            return;
        }
        importTextAndObjectStyles(activeDoc, sourceFile);
        copyCompositeFonts(sourceDoc, activeDoc);
        alert("字符样式、段落样式、对象样式及复合字体已成功导入！");
        sourceDoc.close(SaveOptions.NO);
    } catch (e) {
        alert("发生错误: " + e.message);
    }
}

// 获取当前激活文档
function getActiveDocument() {
    try {
        return app.activeDocument;
    } catch (e) {
        throw new Error("没有打开的文档。");
    }
}

// 选择模板文件
function selectInddFile() {
    try {
        //需要选择indd或者indt文件
        return File.openDialog("选择一个模板文件", "*.indd;*.indt");
    } catch (e) {
        alert("文件选择出错: " + e.message);
        return null;
    }
}

// 打开源文档（不显示）
function openSourceDocument(file) {
    try {
        return app.open(file, false);
    } catch (e) {
        alert("无法打开文件: " + e.message);
        return null;
    }
}

// 导入字符、段落、对象样式
function importTextAndObjectStyles(targetDoc, file) {
    try {
        targetDoc.importStyles(ImportFormat.textStylesFormat, file, GlobalClashResolutionStrategy.loadAllWithOverwrite);
        targetDoc.importStyles(ImportFormat.OBJECT_STYLES_FORMAT, file, GlobalClashResolutionStrategy.loadAllWithOverwrite);
    } catch (e) {
        alert("导入样式出错: " + e.message);
    }
}

// 复制复合字体
function copyCompositeFonts(sourceDoc, targetDoc) {
    try {
        var sourceFonts = sourceDoc.compositeFonts;
        var targetFonts = targetDoc.compositeFonts;
        $.writeln("源文档复合字体数量: " + sourceFonts.length);
        for (var i = 1; i < sourceFonts.length; i++) {
            var srcFont = sourceFonts[i];
            $.writeln("复合字体[" + i + "] 名称: " + srcFont.name);

            // 新建复合字体，继承名称
            var newFont = null;
            try {
                newFont = targetFonts.add({ name: srcFont.name });
            } catch (e) {
                $.writeln("无法新建复合字体: " + srcFont.name + "，原因: " + e.message);
                continue;
            }

            // 遍历并复制CompositeFontEntries
            var srcFontEntries = srcFont.compositeFontEntries;
            var newFontEntries = newFont.compositeFontEntries;
            $.writeln("  源复合字体包含 " + srcFontEntries.length + " 个CompositeFontEntry");
            for (var j = 0; j < srcFontEntries.length; j++) {
                var srcEntry = srcFontEntries[j];
                var newEntry = newFontEntries[j];
                if (!newEntry) continue;
                $.writeln("    复制CompositeFontEntry: " + srcEntry.name);

                // j=0时为汉字，跳过部分参数
                if (j === 0) {
                    try { newEntry.name = srcEntry.name; } catch (e) {}
                    try { newEntry.appliedFont = srcEntry.appliedFont; } catch (e) {}
                    try { newEntry.fontStyle = srcEntry.fontStyle; } catch (e) {}
                    try { newEntry.customCharacters = srcEntry.customCharacters; } catch (e) {}
                    try { newEntry.label = srcEntry.label; } catch (e) {}
                    // 跳过：baselineShift, horizontalScale, verticalScale, relativeSize, scaleOption
                } else {
                    try { newEntry.name = srcEntry.name; } catch (e) {}
                    try { newEntry.appliedFont = srcEntry.appliedFont; } catch (e) {}
                    try { newEntry.fontStyle = srcEntry.fontStyle; } catch (e) {}
                    try { newEntry.horizontalScale = srcEntry.horizontalScale; } catch (e) {}
                    try { newEntry.verticalScale = srcEntry.verticalScale; } catch (e) {}
                    try { newEntry.baselineShift = srcEntry.baselineShift; } catch (e) {}
                    try { newEntry.customCharacters = srcEntry.customCharacters; } catch (e) {}
                    try { newEntry.label = srcEntry.label; } catch (e) {}
                    try { newEntry.relativeSize = srcEntry.relativeSize; } catch (e) {}
                    try { newEntry.scaleOption = srcEntry.scaleOption; } catch (e) {}
                }
            }
        }
    } catch (e) {
        alert("复制复合字体出错: " + e.message);
    }
}

// 执行主函数
main();
