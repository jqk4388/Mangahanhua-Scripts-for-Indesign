/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
MANGA LETTERING AUTOMATION LIBRARY
Paul Starr
October 2019
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

// Set up some convenient variables
var theDoc    = app.activeDocument
var thePages  = app.activeDocument.pages
var theLayers = app.activeDocument.layers
var theMaster = app.activeDocument.masterSpreads.itemByName('A-Master')

// 将脚本作为单个可撤销步骤执行
// 参数：1：要执行的脚本，2：一个向用户描述脚本操作的字符串
// 返回：可能是所执行脚本的返回值。

function KTUDoScriptAsUndoable(theScript, scriptDesc) {
    if (parseFloat(app.version) < 6) {
        return theScript() // execute script
    } else {
        app.doScript(theScript, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, scriptDesc)
    }
}

// 在每个页面上执行脚本
// 参数：1：要执行的脚本，2：向用户描述该脚本作用的字符串
// 作为参数传递给此脚本的函数应接受单个Page对象作为其参数，并对该页面或其中包含的对象进行一些转换
// 返回：可能是所执行脚本的返回值。

function KTUDoEveryPage(theScript, scriptDesc) {
    if (parseFloat(app.version) < 6) {
        return theScript(thePages[i]) // execute script on current page, pre CS6
    } else {
        for (var i = 0; i < thePages.length; i++) { // execute script on current page, post CS6
            app.doScript(theScript(thePages[i]), ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, scriptDesc)
        }
    }
}


// MATCH PAGE TO BLEED SIZE

function KTUMatchFrameToBleedSize(theFrame) {
    // set the measurement units to Points, so our math lower down will work out
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
    // set the ruler to "spread", again so our math works out.
    oldOrigin = theDoc.viewPreferences.rulerOrigin // save old ruler origin
    theDoc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN

    if (theDoc.documentPreferences.pageBinding == PageBindingOptions.leftToRight) { // if the book's laid out left-to-right
        if (theFrame.parentPage.index % 2 == 0) { // if we’re on a left-side page
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - 9, // Same, but for right-side pages
                theFrame.parentPage.bounds[1] - 9, 
                theFrame.parentPage.bounds[2] + 9, 
                theFrame.parentPage.bounds[3] + 0];
        } else { // we must be on a right-side page
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - 9, // Adjust the dimensions to give 1/8" bleed on right-side pages
                theFrame.parentPage.bounds[1] - 0, // 
                theFrame.parentPage.bounds[2] + 9, 
                theFrame.parentPage.bounds[3] + 9];
        }
       
        } 
    else { // if the book is laid out right-to-left
        if (theFrame.parentPage.index % 2 == 0) { // if we’re on a right-side page
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - 9, // Adjust the dimensions to give 1/8" bleed
                theFrame.parentPage.bounds[1] + 0, // on right-side pages
                theFrame.parentPage.bounds[2] + 9, 
                theFrame.parentPage.bounds[3] + 9];
        } else { // we must be on a left-side page
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - 9, // Same, but for right-side pages
                theFrame.parentPage.bounds[1] - 9, 
                theFrame.parentPage.bounds[2] + 9, 
                theFrame.parentPage.bounds[3] + 0];
        }
    }
    // return ruler and measuremeant prefs to previous values
    theDoc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN
    app.scriptPreferences.measurementUnit = AutoEnum.AUTO_VALUE
    return theFrame
}

// TOGGLE BINDING DIRECTION
// returns documentPreferences.pageBinding.PageBindingOptions
function KTUToggleBindingDirection(aDocument) {
    aDocument.documentPreferences.pageBinding = 
        aDocument.documentPreferences.pageBinding == PageBindingOptions.LEFT_TO_RIGHT ? 
            PageBindingOptions.RIGHT_TO_LEFT : 
            PageBindingOptions.LEFT_TO_RIGHT
    return aDocument.documentPreferences.pageBinding
}

// CHECK BINDING FOR RIGHT-TO-LEFT SETTING
function KTUIsBindingCorrect(aDocument) {
    if (aDocument.documentPreferences.pageBinding == PageBindingOptions.RIGHT_TO_LEFT) {
        return true
    } else {
        return false
    }
}

// LOCK ALL ITEMS IN A DOCUMENT
function KTULockAllItems(aDocument) {
        aDocument.pageItems.everyItem().locked = true
}

// UNLOCK ALL ITEMS
function KTUUnLockAllItems(aDocument) {
        aDocument.pageItems.everyItem().locked = false
}

// APPLY MASTER TO PAGE
// Takes a page and a master page as arguments, and applies the given master page to the given page
function KTUApplyMasterToPage(aPage,aMaster) {
    aPage.appliedMaster = aMaster
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
        if (action.isValid && action.enabled) action.invoke();
        else {
            alert("找不到菜单命令或者无法点击。");
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
        
        if (Folder.fs === "Windows") {
            // 使用系统命令获取剪贴板内容（Windows）
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
        } else if (Folder.fs === "Macintosh") {
            // 使用AppleScript获取剪贴板内容（Mac）
            var scriptFile = new File(Folder.temp.absoluteURI + "/get_clipboard.scpt");
            var scriptContent = 'set clipboardText to the clipboard as string\n' +
                               'set fileRef to open for access POSIX file "' + tempFile.fsName + '" with write permission\n' +
                               'set eof fileRef to 0\n' +
                               'write clipboardText to fileRef\n' +
                               'close access fileRef';
            
            scriptFile.open("w");
            scriptFile.write(scriptContent);
            scriptFile.close();
            
            // 执行AppleScript
            var result = app.doScript("run script (read file \"" + scriptFile.fsName + "\" as alias)", 
                                      ScriptLanguage.APPLESCRIPT_LANGUAGE);
            
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

// Polyfills to make scripting more like ES6 syntax
Array.prototype.filter = function (callback) {
  var filtered = [];
  for (var i = 0; i < this.length; i++)
    if (callback(this[i], i, this)) filtered.push(this[i]);
  return filtered;
};
Array.prototype.forEach = function (callback) {
  for (var i = 0; i < this.length; i++) callback(this[i], i, this);
};
Array.prototype.includes = function (item) {
  for (var i = 0; i < this.length; i++) if (this[i] == item) return true;
  return false;
};
