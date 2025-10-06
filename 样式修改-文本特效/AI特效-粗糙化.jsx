#include "../Library/KTUlib.jsx"
/**
 * 主函数：实现AI粗糙化效果
 */
function aiRoughenEffect() {
    try {
        // 检查是否有选中的文本框
        if (app.selection.length === 0) {
            alert("请先选中一个文本框");
            return;
        }

        // 检查文档是否已保存
        if (theDoc.fullName == null || theDoc.fullName.toString() === "") {
            alert("请先保存文档");
            return;
        }

        // 获取文档所在文件夹路径
        var docFolder = theDoc.filePath;
        
        // 验证选中的对象是否都是文本框
        var textFrames = [];
        for (var i = 0; i < app.selection.length; i++) {
            if (app.selection[i] instanceof TextFrame) {
                textFrames.push(app.selection[i]);
            }
        }
        
        if (textFrames.length > 1) {
            alert("只选择一个文本框！");
            return;
        }
        
        // 记录当前图层状态
        var layerStates = saveLayerStates();
        
        // 隐藏不需要的元素
        hideUnselectedElements(textFrames);
        
        // 处理每个选中的文本框
        for (var j = 0; j < textFrames.length; j++) {
            var tempFrame = processTextFrame(textFrames[j], docFolder, j)
        }
        // 还原图层状态
        restoreLayerStates(layerStates, tempFrame);

    } catch (err) {
        alert("发生错误，请先保存文件！ \n" + err.description);
        // 尝试还原图层状态
        try {
            restoreLayerStates(layerStates);
        } catch(e) {
            // 忽略还原状态时的错误
        }
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
/**
 * 保存图层状态
 */
function saveLayerStates() {
    var states = [];
    for (var i = 0; i < theLayers.length; i++) {
        var layer = theLayers[i];
        states.push({
            layer: layer,
            visible: layer.visible,
            items: []
        });
        
        // 保存当前图层中各项目的可见性状态
        var items = layer.pageItems;
        for (var j = 0; j < items.length; j++) {
            states[i].items.push({
                item: items[j],
                visible: items[j].visible
            });
        }
    }
    return states;
}

/**
 * 隐藏未选中的元素（仅影响选中文本框所在的页面）
 */
function hideUnselectedElements(selectedFrames) {
    // 隐藏所有图层
    for (var i = 0; i < theLayers.length; i++) {
        theLayers[i].visible = false;
    }

    // 记录已处理过的页面，避免重复
    var processedPages = {};

    // 显示包含选中文本框的图层，并隐藏该页面中非选中对象
    for (var j = 0; j < selectedFrames.length; j++) {
        var frame = selectedFrames[j];
        var layer = frame.itemLayer;
        var page = frame.parentPage;

        // 显示该图层
        layer.visible = true;

        // 避免重复处理同一页
        var pageKey = layer.name + "_" + page.name;
        if (processedPages[pageKey]) continue;
        processedPages[pageKey] = true;

        // 仅处理该页面上的对象
        var pageItems = page.pageItems;
        for (var k = 0; k < pageItems.length; k++) {
            var item = pageItems[k];
            if (item.itemLayer != layer) continue; // 只处理当前图层的对象

            // 判断是否为选中框
            var isSelected = false;
            for (var l = 0; l < selectedFrames.length; l++) {
                if (item == selectedFrames[l]) {
                    isSelected = true;
                    break;
                }
            }

            item.visible = isSelected;
        }
    }
}


/**
 * 还原图层状态
 */
function restoreLayerStates(states,tempFrame) {
    if (!states) return;
    
    for (var i = 0; i < states.length; i++) {
        var state = states[i];
        state.layer.visible = state.visible;
        
        for (var j = 0; j < state.items.length; j++) {
            try {
                state.items[j].item.visible = state.items[j].visible;
                tempFrame.visible = false;
            } catch(e) {
                // 忽略单项恢复错误
            }
        }
    }
}

/**
 * 处理单个文本框
 */
function processTextFrame(frame, folder, index) {
    try {
        // 检查文本框是否有内容
        if (frame.contents == null || frame.contents.toString().length === 0) {
            alert("文本框为空，跳过处理");
            return;
        }
        if (frame.overflows) {
            frame.fit(FitOptions.FRAME_TO_CONTENT);
        }
        
        // 获取文本框所在的页面
        var page = frame.parentPage;
        if (!page) {
            page = frame.parent instanceof MasterSpread ? null : frame.parentPage;
            if (!page) {
                alert("无法确定文本框所在页面");
                return;
            }
        }
        
        // 提取前5个汉字作为文件名
        var fileName = extractChineseCharacters(frame.contents.toString(), 5);
        if (fileName.length === 0) {
            // 如果没有汉字，使用索引作为文件名
            fileName = "文本框_" + (index + 1);
        }
        
        // 添加时间戳避免重名
        var timestamp = new Date().getTime();
        var docName = app.activeDocument.name.replace(/\.[^\.]+$/, ""); 
        fileName = docName + "_" + fileName + "_" + timestamp;
        
        // 定义文件路径
        var filePath = Folder(folder + "/" + fileName + ".ai");
        //复制一个frame的副本
        var tempFrame = frame.duplicate();
        tempFrame.visible = false;        
        // 将文本框内容转曲
        var outlinedGroup = frame.createOutlines();
        // 导出为PDF（AI格式）
        exportToAI(page, filePath);
        // 删除转曲后的对象
        removeOutlined(outlinedGroup);
        // 调用Illustrator进行粗糙化处理
        roughenInIllustrator(filePath);        
        // 将处理后的AI文件置入到文档中
        placeAIFile(page, filePath);
        return tempFrame;
    } catch (err) {
        alert("处理文本框时出错: " + err.description);
    }
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
 * 对 AI 文件中所有剪切组内的复合路径应用“粗糙化”效果
 * @param {string} filePath - Illustrator 文件路径
 * @param {number} sizePercent - 粗糙化大小（百分比值，例如 1 表示 1%）
 * @param {number} detailPerInch - 每英寸细节数
 * @param {boolean} smooth - 是否使用平滑点（true=平滑，false=尖角）
 */
function illustratorRoughenScript(filePath, sizePercent, detailPerInch) { 
    var doc = app.open(new File(filePath));

    function applyEffectToGroup(group) {
        if (group.clipped) {
            for (var j = 0; j < group.compoundPathItems.length; j++) {
                try {
                    LE_Roughen(group.compoundPathItems[j], {
                        amount: sizePercent,              // 大小：百分比
                        absoluteness: 0,        // 相对
                        segmentsPerInch: detailPerInch,    // 细节
                        smoothness: 1           // 平滑
                    });
                } catch (e) {
                    $.writeln("跳过复合路径：" + e);
                }
            }
        }

        // 若组中还有子组，则递归
        for (var k = 0; k < group.groupItems.length; k++) {
            applyEffectToGroup(group.groupItems[k]);
        }
    }

    function LE_Roughen(item, options) {
    try {
        var defaults = {
            amount: 5,
            absoluteness: 0,       /* 0 or false = relative, 1 or true = absolute */
            segmentsPerInch: 10,
            smoothness: 0,         /* 0 = corners, 1 = smooth, 0.5 = halfway, 1.5 = too much? */
            expandAppearance: false
        }
        var o = defaultsObject(item, defaults, options, arguments.callee)
        var xml = '<LiveEffect name="Adobe Roughen"><Dict data="R asiz #1 R size #2 R absoluteness #3 R dtal #4 R roundness #5 "/></LiveEffect>'
            .replace(/#1/, o.amount)
            .replace(/#2/, o.amount)
            .replace(/#3/, Number(o.absoluteness))
            .replace(/#4/, o.segmentsPerInch)
            .replace(/#5/, o.smoothness);
        applyEffect(item, xml, o.expandAppearance);
    } catch (error) {
        alert('LE_Roughen failed: ' + error);
    }
}

    /**
     * 合并默认参数与用户选项
     * @param {Object} defaults - 默认参数
     * @param {Object} [options] - 传入的选项
     * @param {String} [funcName] - 当前函数名（用于报错）
     * @returns {Object} - 合并后的参数对象
     */
    function defaultsObject(item, defaults, options, funcName) {
        try {
            if (!defaults && !options) return {};
            if (!defaults) return options;
            if (!options) return defaults;
            for (var key in options) {
                if (options.hasOwnProperty(key)) {
                    defaults[key] = options[key];
                }
            }
            return defaults;
        } catch (error) {
            throw new Error(funcName + ' failed to parse options object. ' + error);
        }
    }


    /**
     * 对单个或多个 Illustrator 对象应用 LiveEffect
     * @param {(PageItem|Array)} item - 可是单个对象或多个对象数组
     * @param {String} xml - LiveEffect 的 XML 字符串
     * @param {Boolean} [expand=false] - 是否执行“扩展外观”
     */
    function applyEffect(item, xml, expand) {
        if (item == undefined) throw new Error('applyEffect failed: No item provided.');

        // 判断是单个对象还是数组
        var items = [];
        if (item instanceof Array) {
            items = item;
        } else if (item.typename != undefined) {
            items = [item];
        } else {
            throw new Error('applyEffect failed: Unexpected item type.');
        }

        // 对每个项目应用效果
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it || !it.applyEffect) continue;
            try {
                it.applyEffect(xml);
                if (expand) app.executeMenuCommand('expandStyle');
            } catch (e) {
                $.writeln('⚠️ applyEffect error on item ' + i + ': ' + e);
            }
        }
    }

    // 遍历所有图层
    if (doc && doc.layers.length > 0) {
        for (var l = 0; l < doc.layers.length; l++) {
            var layer = doc.layers[l];
            for (var g = 0; g < layer.groupItems.length; g++) {
                applyEffectToGroup(layer.groupItems[g]);
            }
        }
    }

    doc.save();
    doc.close();
}

/**
 * 调用Illustrator进行粗糙化处理
 */
function roughenInIllustrator(filePath) {
    try {
        // 检查BridgeTalk是否可用
        if (typeof BridgeTalk === "undefined") {
            alert("BridgeTalk不可用，无法调用Illustrator");
            return;
        }
        
        // 检查Illustrator是否可用
        if (!BridgeTalk.isRunning("illustrator")) {
            alert("Illustrator未运行，请先启动Illustrator");
            return;
        }
        
        // 获取illustratorRoughenScript函数的字符串表示
        var scriptFunction = illustratorRoughenScript.toString();
        
        // 构造发送给Illustrator的完整脚本
        var script = 
            "var filePath = '" + filePath.fsName.replace(/\\/g, "\\\\") + "';\n" +
            scriptFunction + ";\n" +
            "illustratorRoughenScript(filePath,0.15,50);";
        
        // 创建BridgeTalk消息
        var bt = new BridgeTalk();
        bt.target = "illustrator";
        bt.body = script;

        var result = false;
        bt.onResult = function(resObj) {
            result = resObj.body == "true";
        };
        
        bt.onError = function(errObj) {
            result = false;
        };
        
        
        // 发送消息并等待处理完成
        bt.send();
        
        // 等待一段时间确保处理完成
        sleepWithEvents(3000);
        return result;
    } catch (err) {
        alert("调用Illustrator处理时出错: " + err.description + "\n将继续执行后续操作");
        return false;
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

/**
 * 将AI文件置入到文档中
 */
function placeAIFile(page, filePath) {
    try {
        // 在相同页面上置入AI文件
        var placedAsset = page.place(new File(filePath.fsName))[0];
        placedAsset.parent.locked = true;
    } catch (err) {
        alert("置入AI文件时出错: " + err.description);
    }
}

// 执行主函数
aiRoughenEffect();