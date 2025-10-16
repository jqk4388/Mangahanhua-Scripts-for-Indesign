#include "../Library/KTUlib.jsx"
var kDebugSampleScriptName = "链接图隐藏图层 ネーム.jsx";
// 主函数
function main() {
    try {
        // 检查是否有打开的文档
        if (app.documents.length === 0) {
            alert("请先打开一个文档");
            return;
        }
        
        // 获取脚本文件名以确定要隐藏的图层名
        var scriptName = getScriptName();
        var layerName = getLayerNameFromScriptName(scriptName);
        $.writeln("图层名：" + layerName);
        
        var doc = app.activeDocument;
        var selectionType = getUserSelection();
        
        if (selectionType == null) {
            return; // 用户取消操作
        }
        
        var links = getLinksToProcess(doc, selectionType);
        
        if (links.length == 0) {
            alert("没有找到符合条件的链接图片");
            return;
        }
        
        // 处理链接图像
        var processedCount = 0;
        for (var i = 0; i < links.length; i++) {
            if (processLink(links[i], layerName)) {
                processedCount++;
            }
        }
        
        alert("处理完成，共处理了 " + processedCount + " 个链接图像。");
        
    } catch (e) {
        alert("脚本执行出错: " + e.message);
    }
}

// 获取脚本文件名
function getScriptName() {
    try {
        return decodeURIComponent(app.activeScript.name);
    } catch (error) {
        return kDebugSampleScriptName;
    }
    
}

// 从脚本文件名提取图层名
function getLayerNameFromScriptName(scriptName) {
    // 移除文件扩展名
    var nameWithoutExtension = scriptName.replace(/\.[^\.]*$/, "");
    // 以空格为分隔符拆分文件名
    var nameParts = nameWithoutExtension.split(" ");
    // 获取最后一个元素作为图层名
    return nameParts[nameParts.length - 1];
}

// 检查链接是否为PSD格式
function isPSDLink(link) {
    var filePath = link.filePath;
    if (!filePath) return false;
    
    filePath = filePath.toLowerCase();
    return filePath.substr(-4) === ".psd";
}

// 处理单个链接
function processLink(link, layerName) {
    try {
        // 检查链接状态
        if (link.status !== LinkStatus.NORMAL) {
            return false;
        }
        
        var GraphicLayers = link['parent']['graphicLayerOptions']['graphicLayers'];
        var GraphicLayer;
        for (var i = 0; i < GraphicLayers.length; i++) {
            if (GraphicLayers[i]['name'] === layerName) {
                GraphicLayer = GraphicLayers[i];
                if (GraphicLayers['0']['currentVisibility'] === false) {
                    GraphicLayers['0']['currentVisibility'] = true;
                } else {
                    GraphicLayers['0']['currentVisibility'] = false;
                }
                break;
            }
        }
        
        // 更新链接
        link.update();
        
        return true;
    } catch (e) {
        $.writeln("处理链接时出错: " + e.message);
        return false;
    }
}

// 执行主函数
KTUDoScriptAsUndoable(function() { main(); }, "显示/隐藏图层");