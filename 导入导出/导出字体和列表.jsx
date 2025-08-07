/**
 * 导出InDesign文档使用的字体列表
 * 支持复制字体文件到Document fonts文件夹
 */

// 获取桌面路径
function getDesktopPath() {
    return Folder.desktop.fsName;
}

// 获取文档中使用的字体列表
function getUsedFonts(doc) {
    var usedFonts = {};
    try {
        // 遍历所有文本框
        for(var i = 0; i < doc.textFrames.length; i++) {
            var textFrame = doc.textFrames[i];
            if(textFrame.contents.length > 0) {
                var texts = textFrame.texts;
                for(var j = 0; j < texts.length; j++) {
                    var font = texts[j].appliedFont;
                    if(font && font.fullNameNative) {
                        usedFonts[font.fullNameNative] = font;
                    }
                }
            }
        }
    } catch(e) {
        alert("获取字体时出错：" + e);
    }
    return usedFonts;
}

// 保存字体列表到文件
function saveFontList(fonts, filePath) {
    try {
        var file = new File(filePath);
        file.encoding = "UTF-8";
        file.open("w");
        
        var fontNames = [];
        for(var name in fonts) {
            fontNames.push(name);
        }
        
        file.write(fontNames.join("\r\n"));
        file.close();
        return true;
    } catch(e) {
        alert("保存字体列表时出错：" + e);
        return false;
    }
}

// 复制字体文件到Document fonts文件夹
function copyFontsToDocumentFonts(fonts) {
    try {
        var destFolder = new Folder(getDesktopPath() + "/Document fonts");
        if(!destFolder.exists) {
            destFolder.create();
        }
        
        for(var name in fonts) {
            var font = fonts[name];
            if(font.location) {
                var sourceFile = new File(font.location);
                var destFile = new File(destFolder + "/" + sourceFile.name);
                sourceFile.copy(destFile);
            }
        }
        return true;
    } catch(e) {
        alert("复制字体文件时出错：" + e);
        return false;
    }
}

// 主函数
function main() {
    if(app.documents.length === 0) {
        alert("请先打开一个文档！");
        return;
    }
    
    var doc = app.activeDocument;
    var fonts = getUsedFonts(doc);
    
    // 保存字体列表
    var docName = doc.name.replace(/\.[^\.]+$/, ""); // 去掉扩展名
    var txtPath = getDesktopPath() + "/" + docName + "_字体列表.txt";
    if(saveFontList(fonts, txtPath)) {
        alert("字体列表已保存到桌面：字体列表.txt");
        
        // 询问是否复制字体文件
        if(confirm("是否将字体文件复制到桌面的Document fonts文件夹？")) {
            if(copyFontsToDocumentFonts(fonts)) {
                alert("字体文件已复制完成！");
            }
        }
    }
}

// 执行脚本
main();
