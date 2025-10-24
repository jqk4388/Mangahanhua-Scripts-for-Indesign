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
/**
 * 隐藏未选中的元素（仅影响选中文本框所在的页面）
 */
function hideUnselectedElements(selectedFrames) {
    if (!selectedFrames || selectedFrames.length === 0) return;

    // 以第一个选中的文本框为准，获取图层 A 和页面 B
    var targetFrame = selectedFrames[0];
    var layerA = targetFrame.itemLayer;
    var pageB = targetFrame.parentPage;

    // 隐藏所有图层，除了图层 A
    for (var i = 0; i < theLayers.length; i++) {
        try {
            theLayers[i].visible = (theLayers[i].id === layerA.id);
        } catch (e) {
            // 忽略无法设置可见性的对象
        }
    }

    // 如果无法确定页面，则不做页面层面的隐藏
    if (!pageB) return;

    // 构建一个快速查找表，标记需要保留可见性的选中文本框（按 id）
    var keepIds = {};
    for (var j = 0; j < selectedFrames.length; j++) {
        var f = selectedFrames[j];
        if (f && f.isValid) keepIds[f.id] = true;
    }

    // 隐藏页面 B 上的所有元素，除了选中的文本框
    var pageItems = pageB.pageItems;
    for (var k = 0; k < pageItems.length; k++) {
        var item = pageItems[k];
        try {
            item.visible = !!keepIds[item.id];
        } catch (e) {
            // 忽略无法设置可见性的对象
        }
    }
}
/**
 * 保存所有图层及对象的可见状态（支持分组）
 */
function saveLayerStates() {
    var states = [];

    for (var i = 0; i < theLayers.length; i++) {
        var layer = theLayers[i];
        var layerState = {
            id: layer.id,                 // 记录图层 ID
            visible: layer.visible,
            items: []
        };

        // 保存图层中对象的可见性（递归支持分组）
        var items = layer.pageItems;
        for (var j = 0; j < items.length; j++) {
            layerState.items.push(collectItemState(items[j]));
        }

        states.push(layerState);
    }

    return states;
}
/**
 * 恢复所有图层及对象的可见状态
 */
function restoreLayerStates(states, tempFrame) {
    if (!states || !states.length) return;
    var doc = app.activeDocument;

    // 取文本框所在的页面A和图层B（如果有提供 tempFrame）
    var pageA = null, layerB = null;
    if (tempFrame && tempFrame.isValid) {
        try {
            layerB = tempFrame.itemLayer;
            pageA = tempFrame.parentPage;
        } catch (e) {
            layerB = null;
            pageA = null;
        }
    }

    for (var i = 0; i < states.length; i++) {
        var layerState = states[i];
        var layer = getByID(doc.layers, layerState.id);
        if (!layer) continue;

        try {
            // 恢复图层可见性
            try { layer.visible = layerState.visible; } catch(e) {}

            // 对于目标图层B，先恢复图层可见性
            try { layer.visible = layerState.visible; } catch(e) {}

            // 如果无法确定页面A，则不遍历具体元素（保持图层整体恢复）
            if (!pageA) continue;

            // 构建页面A在该图层上的顶层对象快速查找表（仅顶层 pageItems）
            var pageItemMap = {};
            try {
                var topItems = layer.pageItems;
                for (var p = 0; p < topItems.length; p++) {
                    var ti = topItems[p];
                    // 确保该顶层对象位于页面A
                    try {
                        if (ti.parentPage && ti.parentPage.id === pageA.id) {
                            pageItemMap[ti.id] = true;
                        }
                    } catch(e) {}
                }
            } catch(e) {
                // 若获取 pageItems 失败，则跳过具体恢复
                continue;
            }

            // 仅遍历并恢复属于页面A 的那些状态（避免遍历整层所有元素状态）
            for (var j = 0; j < layerState.items.length; j++) {
                var itemState = layerState.items[j];
                if (!itemState) continue;
                if (!pageItemMap[itemState.id]) continue; // 不是页面A上的顶层对象，跳过
                try {
                    restoreItemState(itemState, layer);
                } catch(e) {
                    // 忽略单项恢复错误
                }
            }

        } catch (e) {
            // 忽略单个图层恢复错误，继续处理下一个图层
        }
    }

    // 最后统一隐藏临时框
    if (tempFrame && tempFrame.isValid) {
        try { tempFrame.visible = false; } catch(e) {}
    }
}
/**
 * 递归收集对象状态（含分组）
 */
function collectItemState(item) {
    var state = {
        id: item.id,
        visible: item.visible,
        type: item.constructor.name,
        children: []
    };

    // 如果是编组对象，递归保存内部子对象状态
    if (item instanceof Group) {
        var subItems = item.pageItems;
        for (var i = 0; i < subItems.length; i++) {
            state.children.push(collectItemState(subItems[i]));
        }
    }

    return state;
}

/**
 * 递归恢复单个对象状态
 */
function restoreItemState(state, parentContainer) {
    var item = getByID(parentContainer.pageItems, state.id);
    if (!item) return; // 被删除则忽略

    try {
        item.visible = state.visible;
    } catch(e) {}

    if (state.children && state.children.length > 0 && item instanceof Group) {
        for (var i = 0; i < state.children.length; i++) {
            restoreItemState(state.children[i], item);
        }
    }
}
/**
 * 根据 ID 从集合中查找对象
 */
function getByID(collection, id) {
    for (var i = 0; i < collection.length; i++) {
        if (collection[i].id == id) return collection[i];
    }
    return null;
}

/**
 * 提取字符串中的前N个汉字
 */
function extractChineseCharacters(str, count) {
    var chineseChars = "";
    for (var i = 0; i < str.length && chineseChars.length < count; i++) {
        var charCode = str.charCodeAt(i);
        // 判断是否为汉字（基本汉字范围）
        if (charCode >= 0x4e00 && charCode <= 0x9fff) {
            chineseChars += str.charAt(i);
        }
    }
    return chineseChars;
}

/**
 * 导出页面为AI文件
 */
function exportToAI(page, filePath) {
    try {
        // 使用PDF导出预设"高质量打印"
        var pdfExportPresets = app.pdfExportPresets;
        var preset = null;
        
        // 尝试查找"高质量打印"预设
        try {
            preset = pdfExportPresets.itemByName("高质量打印");
            if (!preset.isValid) {
                preset = pdfExportPresets.itemByName("[高质量打印]");
            }
        } catch(e) {
            // 预设不存在，使用默认预设
        }
        
        // 如果找不到特定预设，使用第一个可用预设
        if (!preset || !preset.isValid) {
            preset = pdfExportPresets.firstItem();
        }
        
        // 设置导出参数
        app.pdfExportPreferences.pageRange = page.name;
        
        // 执行导出
        theDoc.exportFile(ExportFormat.PDF_TYPE, filePath, false, preset);
        
    } catch (err) {
        throw new Error("导出PDF时出错: " + err.description);
    }
}
/**
 * 将AI文件置入到文档中
 */
function placeAIFile(page, filePath) {
    try {
        // 在相同页面上置入AI文件
        app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_BLEED
        var placedAsset = page.place(new File(filePath.fsName))[0];
        app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_CONTENT_VISIBLE_LAYERS
        placedAsset.parent.locked = true;
    } catch (err) {
        alert("置入AI文件时出错: " + err.description);
    }
}
function removeOutlined(obj) {
    if (obj == null) return;

    // 如果对象有 remove 方法，直接删除
    if (typeof obj.remove === "function") {
        obj.remove();
    } 
    // 如果是数组或类数组
    else if (obj instanceof Array || (typeof obj === "object" && "length" in obj)) {
        for (var i = 0; i < obj.length; i++) {
            removeOutlined(obj[i]); // 递归删除每个元素
        }
    } 
    // 如果是单个对象但没有 remove 方法（比如某些特殊返回类型）
    else {
        // 可以尝试通过 parent 删除
        if (obj.hasOwnProperty("parent") && obj.parent) {
            obj.parent.remove();
        }
    }
}
// 获取字号
function getFontSize(textFrame) {
    try {
        return textFrame['parentStory']['pointSize'];
    } catch (e) {
        return 12; // 默认值
    }
}
/**
 * 修改文本框的 Tracking（字距），并如果有 overset（溢流）则 fit
 * @param {TextFrame} tf
 * @param {Number} trackingValue - 文字跟踪（追踪），单位 InDesign 跟踪值
 */
function applyTrackingAndFit(tf, trackingValue) {
    try {
        if (!tf || !tf.texts || tf.texts.length === 0) return;
        var t = tf.texts[0];
        // 对全文字符设置追踪（tracking）
        try {
            t.tracking = trackingValue;
        } catch (e) {
            // 如果追踪属性不可用，则按字符遍历
            for (var ci = 0; ci < t.characters.length; ci++) {
                try {
                    t.characters[ci].tracking = trackingValue;
                } catch (ee) {}
            }
        }

        // 适合文本框：如果溢流则 fit
        try {
            if (tf.overflows) {
                tf.fit(FitOptions.FRAME_TO_CONTENT);
            }
        } catch (e) {
        }
    } catch (e) {
        throw e;
    }
}