var kDebugSampleScriptName = "链接图隐藏图层 ネーム.jsx";
// 主函数
function main() {
    try {
        // 检查是否有打开的文档
        if (app.documents.length === 0) {
            alert("请先打开一个文档");
            return;
        }
        
        var myDocument = app.activeDocument;
        
        // 获取脚本文件名以确定要隐藏的图层名
        var scriptName = getScriptName();
        var layerName = getLayerNameFromScriptName(scriptName);
        $.writeln("图层名：" + layerName);
        
        // 创建对话框询问用户处理范围
        var scope = getUserScopeSelection();
        if (scope === null) {
            return; // 用户取消操作
        }
        
        // 获取链接图像
        var links = getLinksToProcess(myDocument, scope);
        
        if (links.length === 0) {
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

// 获取用户选择的处理范围
function getUserScopeSelection() {
    var dialog = new Window("dialog", "选择处理范围");
    dialog.orientation = "column";
    dialog.alignChildren = "left";
    
    var text = dialog.add("statictext", undefined, "请选择处理范围:");
    
    var rg = dialog.add("group");
    rg.orientation = "row";
    var currentPageRadio = rg.add("radiobutton", undefined, "仅当前页面");
    var allDocumentRadio = rg.add("radiobutton", undefined, "整个文档");
    currentPageRadio.value = true;
    
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    var okBtn = buttonGroup.add("button", undefined, "确定");
    var cancelBtn = buttonGroup.add("button", undefined, "取消");
    
    var result = null;
    
    okBtn.onClick = function() {
        result = currentPageRadio.value ? "current" : "document";
        dialog.close();
    };
    
    cancelBtn.onClick = function() {
        result = null;
        dialog.close();
    };
    
    dialog.show();
    
    return result;
}

// 获取需要处理的链接
function getLinksToProcess(doc, scope) {
    var links = [];
    
    if (scope == "current") {
        // 处理当前页面
        var page = doc.layoutWindows[0].activePage;
        var pageItems = page.allPageItems;
        
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i].hasOwnProperty("images") && pageItems[i].images.length > 0) {
                var image = pageItems[i].images[0];
                if (image.hasOwnProperty("itemLink")) {
                    // 只获取PSD链接
                    var link = image.itemLink;
                    if (isPSDLink(link)) {
                        links.push(link);
                    }
                }
            }
        }
    } else {
        // 处理整个文档
        var allLinks = doc.links;
        for (var i = 0; i < allLinks.length; i++) {
            // 只获取PSD链接
            if (isPSDLink(allLinks[i])) {
                links.push(allLinks[i]);
            }
        }
    }
    
    return links;
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
main();