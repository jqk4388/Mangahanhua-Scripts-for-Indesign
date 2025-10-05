// 主函数
function convertBitmapToGrayscale() {
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
    
    processLinks(links);
}

// 获取用户选择的处理范围
function getUserSelection() {
    var dialog = new Window("dialog", "选择处理范围");
    dialog.orientation = "column";
    dialog.alignChildren = "left";
    
    var text = dialog.add("statictext", undefined, "请选择处理范围:");
    
    var rg = dialog.add("group");
    rg.orientation = "row";
    var currentPageRadio = rg.add("radiobutton", undefined, "仅当前页面");
    var allDocumentRadio = rg.add("radiobutton", undefined, "整个文档");
    var selectedLinksRadio = rg.add("radiobutton", undefined, "当前选中的链接");
    currentPageRadio.value = true;
    
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    var okBtn = buttonGroup.add("button", undefined, "确定");
    var cancelBtn = buttonGroup.add("button", undefined, "取消");
    
    var result = null;
    
    okBtn.onClick = function() {
        if (currentPageRadio.value) {
            result = "current";
        } else if (allDocumentRadio.value) {
            result = "document";
        } else if (selectedLinksRadio.value) {
            result = "selected";
        }
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
function getLinksToProcess(doc, selectionType) {
    var links = [];
    
    if (selectionType == "current") {
        // 处理当前页面
        var page = doc.layoutWindows[0].activePage;
        var pageItems = page.allPageItems;
        
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i].hasOwnProperty("images") && pageItems[i].images.length > 0) {
                var image = pageItems[i].images[0];
                if (image.hasOwnProperty("itemLink") && isSupportedImage(image.itemLink)) {
                    links.push(image.itemLink);
                }
            }
        }
    } else if (selectionType == "selected") {
        // 处理当前选中的链接
        links = getSelectedLinksFromClipboard(doc);
    } else {
        // 处理整个文档
        var allLinks = doc.links;
        for (var i = 0; i < allLinks.length; i++) {
            if (isSupportedImage(allLinks[i])) {
                links.push(allLinks[i]);
            }
        }
    }
    
    return links;
}

// 通过剪贴板获取选中的链接
function getSelectedLinksFromClipboard(doc) {
    var links = [];
    
    // 复制选中的链接信息
    try {
        var action = app.menuActions.itemByName("复制选定链接的信息");
        if (!action.isValid) action = app.menuActions.itemByName("Copy Link Info");
        if (action.isValid) action.invoke();
        else {
            alert("找不到菜单命令");
            return links;
        }
    } catch (error) {
        alert("请先在链接窗口中选中链接：" + error);
    }
    
    // 等待剪贴板更新
    sleepWithEvents(1000);
    
    // 获取剪贴板内容
    var clipboardText = getClipboardText();
    
    if (!clipboardText) {
        alert("无法获取剪贴板内容");
        return links;
    }
    
    // 解析剪贴板内容，提取文件名
    var filenames = parseClipboardContent(clipboardText);
    
    if (filenames.length === 0) {
        alert("未从剪贴板中解析到文件名");
        return links;
    }
    
    // 根据文件名查找对应的链接
    var allLinks = doc.links;
    for (var i = 0; i < filenames.length; i++) {
        var filename = filenames[i];
        for (var j = 0; j < allLinks.length; j++) {
            var link = allLinks[j];
            if (isSupportedImage(link)) {
                var linkFilename = getFileNameFromPath(link.filePath);
                if (linkFilename === filename) {
                    links.push(link);
                    break;
                }
            }
        }
    }
    
    return links;
}

// 从剪贴板内容中解析文件名
function parseClipboardContent(content) {
    var filenames = [];
    var lines = content.split('\n');
    
    // 跳过标题行
    var startIndex = 1;
    if (lines.length > 1 && lines[0].indexOf("名称") !== -1 && lines[0].indexOf("状态") !== -1) {
        startIndex = 1;
    } else {
        startIndex = 0;
    }
    
    for (var i = startIndex; i < lines.length; i++) {
        var line = lines[i];
        if (line) {
            // 提取第一列作为文件名
            var columns = line.split('\t');
            if (columns.length > 0 && columns[0]) {
                filenames.push(columns[0]);
            }
        }
    }
    
    return filenames;
}

// 获取路径中的文件名
function getFileNameFromPath(filePath) {
    if (!filePath) return "";
    
    var parts = filePath.split(/[\/\\]/);
    var filename = parts[parts.length - 1];
    
    return filename;
}

// 获取剪贴板文本内容
function getClipboardText() {
    try {
        // 创建临时文件来获取剪贴板内容
        var tempFile = new File(Folder.temp.absoluteURI + "/clipboard_temp.txt");
        if (tempFile.exists) {
            tempFile.remove();
        }
        
        // 使用系统命令获取剪贴板内容（Windows）
        if (Folder.fs === "Windows") {
            var scriptFile = new File(Folder.temp.absoluteURI + "/get_clipboard.vbs");
            var scriptContent = 'Set objHTML = CreateObject("htmlfile")\n' +
                               'strContent = objHTML.ParentWindow.ClipboardData.GetData("text")\n' +
                               'Set objFSO = CreateObject("Scripting.FileSystemObject")\n' +
                               'Set objFile = objFSO.CreateTextFile("' + tempFile.fsName.replace(/\\/g, "\\\\") + '", True)\n' +
                               'objFile.Write strContent\n' +
                               'objFile.Close\n';
            
            scriptFile.open("w");
            scriptFile.write(scriptContent);
            scriptFile.close();
            
            scriptFile.execute();
            sleepWithEvents(1000);
            
            if (tempFile.exists) {
                tempFile.open("r");
                var content = tempFile.read();
                tempFile.close();
                tempFile.remove();
                scriptFile.remove();
                return content;
            }
        }
        return "";
    } catch (e) {
        $.writeln("获取剪贴板内容时出错: " + e.message);
        return "";
    }
}

// 检查是否为支持的图片格式
function isSupportedImage(link) {
    if (!link) return false;
    
    var filePath = link.filePath;
    if (!filePath) return false;
    
    filePath = filePath.toLowerCase();
    
    // 检查是否为psd或tif格式
    if (filePath.substr(-4) == ".psd" || filePath.substr(-4) == ".tif" || 
        filePath.substr(-5) == ".tiff") {
        return true;
    }
    
    return false;
}

// 处理链接图片
function processLinks(links) {
    // 创建进度条窗口
    var progressWindow = new Window("palette", "处理进度");
    progressWindow.orientation = "column";
    progressWindow.alignChildren = "left";
    
    var progressText = progressWindow.add("statictext", undefined, "正在处理图片...");
    progressText.preferredSize.width = 300;
    
    var progressBar = progressWindow.add("progressbar", undefined, 0, links.length);
    progressBar.preferredSize.width = 300;
    
    var progressInfo = progressWindow.add("statictext", undefined, "0 / " + links.length);
    progressInfo.preferredSize.width = 300;
    
    var cancelButton = progressWindow.add("button", undefined, "取消");
    
    var cancelled = false;
    cancelButton.onClick = function() {
        cancelled = true;
    };
    
    progressWindow.show();
    
    var processedCount = 0;
    
    for (var i = 0; i < links.length; i++) {
        // 更新进度条
        progressBar.value = i;
        progressInfo.text = (i + 1) + " / " + links.length;
        progressWindow.update();
        
        // 检查是否取消
        if (cancelled) {
            break;
        }
        
        if (processSingleLink(links[i])) {
            processedCount++;
        }
        
        // 处理事件队列，保持界面响应
        app.scriptPreferences.enableRedraw = true;
    }
    
    // 关闭进度条窗口
    progressWindow.close();
    
    if (cancelled) {
        alert("用户取消操作。已完成 " + processedCount + " 个文件的处理。");
    } else {
        alert("处理完成！共处理了 " + processedCount + " 个文件。");
    }
}

// 处理单个链接
function processSingleLink(link) {
    try {
        var originalPath = link.filePath;
        if (!originalPath) return false;
        
        // 首先在InDesign中检查图片是否为位图模式
        if (!isBitmapImage(link)) {
            // 如果不是位图，直接返回true表示处理完成（无需处理）
            return true;
        }
        
        // 通过BridgeTalk调用Photoshop处理图片
        if (processImageInPhotoshop(originalPath)) {
            // 更新InDesign中的链接
            var newPath = originalPath;
            
            // 等待文件系统更新
            sleepWithEvents(2000);
            
            // 更新链接
            try {
                link.relink(new File(newPath));
                link.update();
                return true;
            } catch (e) {
                // 如果更新失败，可能是文件还在被占用，再等待一下
                sleepWithEvents(3000);
                try {
                    link.relink(new File(newPath));
                    link.update();
                    return true;
                } catch (e2) {
                    $.writeln("更新链接失败: " + e2.message);
                    return false;
                }
            }
        }
    } catch (e) {
        $.writeln("处理链接时出错: " + e.message);
        return false;
    }
    
    return false;
}

// 检查图片是否为位图模式
function isBitmapImage(link) {
    try {
        // 获取链接对应的图片对象
        var image = link.parent;
        if (image && image.hasOwnProperty('imageTypeName')) {
            // 如果能直接获取到imageTypeName，可以检查是否为Bitmap
            // 注意：这可能需要根据实际API调整
            return image['space'] == "Bitmap"||image['space'] == "黑白";
        }
        // 如果无法直接判断，返回true以保持原有逻辑
        return true;
    } catch (e) {
        $.writeln("检查图片模式时出错: " + e.message);
        // 出错时返回true，保持原有逻辑
        return true;
    }
}

// 在Photoshop中处理图片
function processImageInPhotoshop(imagePath) {
    try {
        if (!BridgeTalk.isRunning("photoshop")) {
            alert("请先启动Photoshop");
            return false;
        }
        
        var bt = new BridgeTalk();
        bt.target = "photoshop";
        
        var psScript = createPhotoshopScript(imagePath);
        
        bt.body = psScript;
        
        var result = false;
        bt.onResult = function(resObj) {
            result = resObj.body == "true";
        };
        
        bt.onError = function(errObj) {
            result = false;
        };
        
        bt.send();
        
        sleepWithEvents(4000);
        
        return result;
    } catch (e) {
        $.writeln("调用Photoshop处理图片时出错: " + e.message);
        return false;
    }
};

// Photoshop脚本函数
function photoshopScriptFunction(fileRef, originalPath) {
    if (fileRef.exists) {
        var doc = app.open(fileRef);
        if (doc.mode == DocumentMode.BITMAP) {
            doc.changeMode(ChangeMode.GRAYSCALE);
            
            // 检查原始文件是否为TIFF格式，如果是则保持TIFF格式并使用LZW压缩
            var ext = originalPath.toLowerCase().split('.').pop();
            var isTiff = (ext === 'tif' || ext === 'tiff');
            if (isTiff) {
                var tiffSaveOptions = new TiffSaveOptions();
                tiffSaveOptions.embedColorProfile = true;
                tiffSaveOptions.alphaChannels = true;
                tiffSaveOptions.layers = true;
                tiffSaveOptions.imageCompression = TIFFEncoding.TIFFLZW;
                doc.saveAs(fileRef, tiffSaveOptions, true, Extension.LOWERCASE);
            } else {
                // PSD文件使用PSD保存选项
                var psdSaveOptions = new PhotoshopSaveOptions();
                psdSaveOptions.embedColorProfile = true;
                psdSaveOptions.alphaChannels = true;
                psdSaveOptions.layers = true;
                doc.saveAs(fileRef, psdSaveOptions, true, Extension.LOWERCASE);
            }
            
            doc.close(SaveOptions.DONOTSAVECHANGES);
            return true;
        } else {
            doc.close(SaveOptions.DONOTSAVECHANGES);
            return true;
        }
    } else {
        return false;
    }
}

// 创建Photoshop脚本
function createPhotoshopScript(imagePath) {
    // 构造函数字符串，替换文件路径和保存路径
    var scriptFunctionStr = photoshopScriptFunction.toString()
        .replace(/this\.savePath/g, "'" + getSavePath(imagePath).replace(/'/g, "\\'") + "'");
    
    var fullScript = 
        "var fileRef = new File('" + imagePath.replace(/'/g, "\\'") + "');\n" +
        "var originalPath = '" + imagePath.replace(/'/g, "\\'") + "';\n" +
        scriptFunctionStr + ";\n" +
        "photoshopScriptFunction(fileRef, originalPath);";
    
    return fullScript;
}

// 获取保存路径
function getSavePath(originalPath) {
    var path = originalPath.toLowerCase();
    if (path.substr(-4) == ".psd") {
        return originalPath;
    } else if (path.substr(-4) == ".tif" || path.substr(-5) == ".tiff") {
        // 对于TIFF文件，保持原格式
        return originalPath;
    }
    return originalPath;
}

// 带事件处理的延迟函数
function sleepWithEvents(milliseconds) {
    var start = new Date().getTime();
    var now = start;
    while (now - start < milliseconds) {
        // 处理事件队列，保持界面响应
        app.scriptPreferences.enableRedraw = true;
        now = new Date().getTime();
    }
}

// 执行主函数
convertBitmapToGrayscale();