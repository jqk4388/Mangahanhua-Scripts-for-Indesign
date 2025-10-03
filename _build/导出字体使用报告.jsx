/**
 * 导出字体使用报告脚本
 * 遍历所有打开的文档，统计每个文档中使用的字体及使用次数
 * 将结果保存为JSON文件到桌面
 */

// 主函数
function main() {
    try {
        // 检查是否有打开的文档
        if (app.documents.length === 0) {
            alert("没有打开的文档");
            return;
        }

        // 获取字体使用统计
        var fontReport = getFontUsageReport();

        // 保存报告到桌面
        var saved = saveReportToDesktop(fontReport);

        if (saved) {
            alert("字体使用报告已保存到桌面");
        } else {
            alert("保存报告时发生错误");
        }
    } catch (e) {
        alert("执行脚本时发生错误: " + e);
    }
}

/**
 * 获取所有打开文档的字体使用报告
 * @return {Array} 字体使用统计数组
 */
function getFontUsageReport() {
    var fontStats = {};

    try {
        // 遍历所有打开的文档
        for (var docIndex = 0; docIndex < app.documents.length; docIndex++) {
            var doc = app.documents[docIndex];
            
            // 遍历文档中的所有文本框（包括编组中的文本框）
            processTextFramesInDocument(doc, fontStats);
        }
    } catch (e) {
        alert("收集字体信息时发生错误: " + e);
    }

    // 将统计结果转换为数组格式
    var report = [];
    for (var fontName in fontStats) {
        report.push({
            "font.name": fontName,
            "usageCount": fontStats[fontName]
        });
    }

    return report;
}

/**
 * 处理文档中的文本框，统计字体使用情况
 * @param {Document} doc - 要处理的文档
 * @param {Object} fontStats - 字体统计对象
 */
function processTextFramesInDocument(doc, fontStats) {
    try {
        // 处理文档中的所有文本框
        processTextFrames(doc.textFrames, fontStats);
        
        // 处理所有组中的文本框
        processGroups(doc.groups, fontStats);
    } catch (e) {
        alert("处理文档 " + doc.name + " 时发生错误: " + e);
    }
}

/**
 * 处理文本框集合
 * @param {TextFrames} textFrames - 文本框集合
 * @param {Object} fontStats - 字体统计对象
 */
function processTextFrames(textFrames, fontStats) {
    for (var i = 0; i < textFrames.length; i++) {
        try {
            var textFrame = textFrames[i];
            processTextFrame(textFrame, fontStats);
        } catch (e) {
            // 忽略单个文本框的错误，继续处理下一个
            continue;
        }
    }
}

/**
 * 处理单个文本框中的文本
 * @param {TextFrame} textFrame - 要处理的文本框
 * @param {Object} fontStats - 字体统计对象
 */
function processTextFrame(textFrame, fontStats) {
    // 检查文本框是否有内容
    if (textFrame.contents !== null && textFrame.contents !== undefined) {
        var contents = textFrame.contents + ""; // 转换为字符串
        // 检查文本框是否包含实际内容（去除空白字符后）
        if (contents.length > 0) {
            try {
                // 遍历文本框中的所有文本
                var textObjects = textFrame.texts;
                for (var j = 0; j < textObjects.length; j++) {
                    var text = textObjects[j];
                    if (text.appliedFont && text.appliedFont.name) {
                        var fontName = text.appliedFont.name + ""; // 转换为字符串
                        // 更新字体统计
                        if (fontStats[fontName]) {
                            fontStats[fontName] = fontStats[fontName] + 1;
                        } else {
                            fontStats[fontName] = 1;
                        }
                    }
                }
            } catch (e) {
                // 忽略文本处理错误
            }
        }
    }
}

/**
 * 处理组中的文本框
 * @param {Groups} groups - 组集合
 * @param {Object} fontStats - 字体统计对象
 */
function processGroups(groups, fontStats) {
    for (var i = 0; i < groups.length; i++) {
        try {
            var group = groups[i];
            // 递归处理组中的文本框
            processTextFrames(group.textFrames, fontStats);
            // 处理嵌套组
            processGroups(group.groups, fontStats);
        } catch (e) {
            // 忽略单个组的错误，继续处理下一个
            continue;
        }
    }
}

/**
 * 保存报告到桌面
 * @param {Array} reportData - 报告数据
 * @return {Boolean} 是否保存成功
 */
function saveReportToDesktop(reportData) {
    try {
        // 获取桌面路径
        var desktopPath = getDesktopPath();
        
        // 生成基于文档名称的文件名
        var docName = "文档";
        if (app.documents.length > 0) {
            docName = app.activeDocument.name;
            // 去除文件扩展名
            var lastDotIndex = docName.lastIndexOf(".");
            if (lastDotIndex > 0) {
                docName = docName.substring(0, lastDotIndex);
            }
        }
        
        var fileName = docName + "_字体报告.json";
        var fullPath = desktopPath + "/" + fileName;

        // 在Mac和Windows上处理路径分隔符
        if (File.fs === "Windows") {
            fullPath = desktopPath + "\\" + fileName;
        }

        // 创建文件对象
        var reportFile = new File(fullPath);
        
        // 设置编码
        reportFile.encoding = "UTF-8";
        
        // 打开文件准备写入
        if (reportFile.open("w")) {
            // 写入JSON数据
            var jsonData = formatJSON(reportData);
            reportFile.write(jsonData);
            reportFile.close();
            return true;
        } else {
            return false;
        }
    } catch (e) {
        alert("保存文件时发生错误: " + e);
        return false;
    }
}

/**
 * 获取桌面路径（兼容Mac和Windows）
 * @return {String} 桌面路径
 */
function getDesktopPath() {
    try {
        // 使用系统提供的桌面路径
        return Folder.desktop.fsName;
    } catch (e) {
        // 备用方法
        if (File.fs === "Windows") {
            return Folder("~").fsName + "\\Desktop";
        } else {
            return Folder("~/Desktop").fsName;
        }
    }
}

/**
 * 格式化JSON数据（避免使用JSON.stringify）
 * @param {Object|Array} obj - 要格式化的对象
 * @return {String} 格式化后的JSON字符串
 */
function formatJSON(obj) {
    if (obj instanceof Array) {
        var result = "[";
        for (var i = 0; i < obj.length; i++) {
            if (i > 0) {
                result += ",";
            }
            result += formatJSON(obj[i]);
        }
        result += "]";
        return result;
    } else if (typeof obj === "object" && obj !== null) {
        var result = "{";
        var first = true;
        for (var key in obj) {
            if (!first) {
                result += ",";
            }
            first = false;
            result += "\"" + key + "\":" + formatJSON(obj[key]);
        }
        result += "}";
        return result;
    } else if (typeof obj === "string") {
        // 转义字符串中的特殊字符
        var escaped = obj.replace(/\\/g, "\\\\").replace(/"/g, "\\\"");
        return "\"" + escaped + "\"";
    } else if (typeof obj === "number") {
        return String(obj);
    } else if (typeof obj === "boolean") {
        return obj ? "true" : "false";
    } else {
        return "null";
    }
}

// 执行主函数
main();