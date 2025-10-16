#include "../Library/KTUlib.jsx"
/**
 * 主函数：实现AI黑底雕效果
 */
function aiBlackWEffect() {
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
 * 恢复所有图层及对象的可见状态
 */
function restoreLayerStates(states, tempFrame) {
    if (!states || !states.length) return;

    for (var i = 0; i < states.length; i++) {
        var layerState = states[i];
        var layer = getByID(app.activeDocument.layers, layerState.id);
        if (!layer) continue;

        try {
            layer.visible = layerState.visible;
        } catch(e) {}

        // 恢复该图层中对象状态
        for (var j = 0; j < layerState.items.length; j++) {
            restoreItemState(layerState.items[j], layer);
        }
    }

    // 最后统一隐藏临时框
    if (tempFrame && tempFrame.isValid) {
        tempFrame.visible = false;
    }
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
        applyTrackingAndFit(frame, getFontSize(frame)*14);
        
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
        // 删除转曲后的对象！！注意再次调用frame就会崩溃！！
        removeOutlined(outlinedGroup);
        // 调用Illustrator进行粗糙化处理
        blackWInIllustrator(filePath,getFontSize(tempFrame));   
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
// 获取字号
function getFontSize(textFrame) {
    try {
        return textFrame['parentStory']['pointSize'];
    } catch (e) {
        return 12; // 默认值
    }
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
 * 对 AI 文件中所有剪切组内的复合路径应用黑底雕效果
 * @param {string} filePath - Illustrator 文件路径
 * @param {number} fontSize - 文本框文字大小
 */
function illustratorBlackWordScript(filePath, fontSize) {
    var doc = app.open(new File(filePath));
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
    function get(type, parent, deep) {
    if (arguments.length == 1 || !parent) {
        parent = app.activeDocument;
        deep = true;
    }
    var result = [];
    if (!parent[type]) return [];
    for (var i = 0; i < parent[type].length; i++) {
        result.push(parent[type][i]);
        if (parent[type][i][type] && deep)
        result = [].concat(result, get(type, parent[type][i], deep));
    }
    return result;
    }
    var LE = {
        functionName: 'LE',
        testMode: false,
        debug: false,
        defaults: {},
        testResults: [], // 补充原代码中使用的testResults属性
        transformPoints: [
            Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
            Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
            Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
        ]
    };

    /**
     * Combines defaults and user-specified options along with a little admin work
     * @param {(PageItem|PageItems)} item - A PageItem or collection or array of PageItems
     * @param {Object} defaults - Object with default properties
     * @param {Object} options - Object with user supplied properties
     * @param {Function} func - The LE_Function
     */
    function LE_defaultsObject(item, defaults, options, func) {
        LE.functionName = func.name;
        try {
            if (defaults == undefined && options == undefined) return {};
            if (defaults == undefined) return options;
            if (options == undefined) return defaults;
            if (options.debug) LE.debug = true;
            for (var key in options) {
                defaults[key] = options[key];
            }
            LE_defaults = defaults;
            return defaults;
        } catch (error) {
            throw new Error(func.name + ' failed to parse options object. ' + error)
        }
    }


    /**
     * Applies the Live Effect, unless in test mode
     * @param {(PageItem|PageItems)} item - A PageItem or collection or array of PageItems
     * @param {String} xml - The Live Effect XML to apply
     * @param {Boolean} expand - Perform Expand Appearance
     */
    function LE_applyEffect(item, xml, expand) {
        if (LE.testMode) {
            LE.testResults.push({ timestamp: new Date(), functionName: LE.functionName, xml: xml });
            return xml;
        } else {
            // work out whether item is single item or multiple
            var items;
            if (item == undefined) {
                throw new Error(LE.functionName + ' failed. No item available.');
            } else if (item[0] == undefined && item.typename != undefined) {
                // a single item
                items = [item];
            }
            if (items.length == undefined) throw new Error(LE.functionName + ' failed. Unexpected item type. [1]');
            // applyEffect to each item
            for (var i = 0; i < items.length; i++) {
                if (items[i].typename == undefined) throw new Error(LE.functionName + ' failed. Unexpected item type. [2]');
                items[i].applyEffect(xml);
                if (expand) LE_expandAppearance(items[i]);
            }
        }
    }

    /**
     * Handles error
     * @param {Error} error - a javascript Error
     */
    function LE_handleError(error) {
        $.writeln(error.message);
    }

    /**
     * Performs Expand Appearance
     * @param {PageItem} item - a PageItem
     */
    function LE_expandAppearance(item) {
        app.redraw();
        app.activeDocument.selection = [item];
        app.executeMenuCommand('expandStyle');
        item = app.activeDocument.selection[0];
    }
    function LE_Transform(item, options) {
        try {
            var defaults = {
                scaleHorzPercent: 100,
                scaleVertPercent: 100,
                moveHorzPts: 0,
                moveVertPts: 0,
                rotateDegrees: 0,
                randomize: false,
                numberOfCopies: 0,
                transformPoint: Transformation.CENTER,  /* must be a Transformation constant, eg. Transformation.BOTTOMRIGHT */
                scaleStrokes: false,
                transformPatterns: true,
                transformObjects: true,
                reflectX: false,
                reflectY: false,
                expandAppearance: false
            }
            var o = LE_defaultsObject(item, defaults, options, arguments.callee)
            o.transformIndex = 4;
            for (var i = 0; i < LE.transformPoints.length; i++) {
                if (o.transformPoint === LE.transformPoints[i]) {
                    o.transformPointIndex = i;
                    break;
                }
            }
            var xml = '<LiveEffect name="Adobe Transform"><Dict data="R scaleH_Percent #1 R scaleV_Percent #2 R scaleH_Factor #3 R scaleV_Factor #4 R moveH_Pts #5 R moveV_Pts #6 R rotate_Degrees #7 R rotate_Radians #8 I numCopies #9 I pinPoint #10 B scaleLines #11 B transformPatterns #12 B transformObjects #13 B reflectX #14 B reflectY #15 B randomize #16 "/></LiveEffect>'
                .replace(/#1/, o.scaleHorzPercent)
                .replace(/#2/, o.scaleVertPercent)
                .replace(/#3/, o.scaleHorzPercent / 100)
                .replace(/#4/, o.scaleVertPercent / 100)
                .replace(/#5/, o.moveHorzPts)
                .replace(/#6/, -o.moveVertPts)
                .replace(/#7/, o.rotateDegrees)
                .replace(/#8/, o.rotateDegrees * Math.PI / 180)
                .replace(/#9/, o.numberOfCopies)
                .replace(/#10/, o.transformPointIndex)
                .replace(/#11/, o.scaleStrokes ? 1 : 0)
                .replace(/#12/, o.transformPatterns ? 1 : 0)
                .replace(/#13/, o.transformObjects ? 1 : 0)
                .replace(/#14/, o.reflectX ? 1 : 0)
                .replace(/#15/, o.reflectY ? 1 : 0)
                .replace(/#16/, o.randomize ? 1 : 0);
            LE_applyEffect(item, xml, o.expandAppearance);
        } catch (error) {
            LE_handleError(error);
        }
    }

    function LE_OffsetPath(item, options) {
        try {
            var defaults = {
                offset: 10,
                joinType: 2,  // joinTypes: 0 = Round, 1 = Bevel , 2 = Miter
                miterLimit: 4,
                expandAppearance: false
            }
            var o = LE_defaultsObject(item, defaults, options, arguments.callee)
            var xml = '<LiveEffect name="Adobe Offset Path"><Dict data="R ofst #1 I jntp #2 R mlim #3 "/></LiveEffect>'
                .replace(/#1/, o.offset)
                .replace(/#2/, o.joinType)
                .replace(/#3/, o.miterLimit);
            LE_applyEffect(item, xml, o.expandAppearance);
        } catch (error) {
            LE_handleError(error);
        }
    }

    function applyMultiAppearance(item,fontSize) {
        var doc = app.activeDocument;

        // 第一层：黑色偏移
        var layer1 = item.duplicate();
        var a = get("pathItems",layer1)
        a.forEach(function(item){item.filled = true;});
        var c1 = new CMYKColor(); c1.black = 100;
        a.forEach(function(item){item.fillColor = c1;});
        var option1 = {
                offset: 0.0085*fontSize,
                joinType: 2,  // joinTypes: 0 = Round, 1 = Bevel , 2 = Miter
                miterLimit: 4,
                expandAppearance: false
            }
        LE_OffsetPath(layer1, option1);

        // 第二层：白色变换
        var layer2 = item.duplicate();
        var b = get("pathItems",layer2)
        b.forEach(function(item){item.filled = true;});
        var c2 = new CMYKColor(); c2.black = 0;
        b.forEach(function(item){item.fillColor = c2;});
        var option2 = {
                scaleHorzPercent: 100,
                scaleVertPercent: 100,
                moveHorzPts: 0.0068*fontSize,
                moveVertPts: 0.0068*fontSize,
                rotateDegrees: 0,
                randomize: false,
                numberOfCopies: Math.ceil(fontSize),
                transformPoint: Transformation.CENTER,  /* must be a Transformation constant, eg. Transformation.BOTTOMRIGHT */
                scaleStrokes: false,
                transformPatterns: true,
                transformObjects: true,
                reflectX: false,
                reflectY: false,
                expandAppearance: false
            }
        LE_Transform(layer2, option2);

        // 第三层：黑描边变换
        var layer3 = item.duplicate();
        var c = get("pathItems",layer3)
        c.forEach(function(item){item.stroked = true;});
        c.forEach(function(item){item.strokeWidth = 0.073*fontSize;});
        var c3 = new CMYKColor(); c3.black = 100;
        c.forEach(function(item){item.strokeColor = c3;});
        c.forEach(function(item){item.strokeMiterLimit = 4;});
        c.forEach(function(item){item.filled = false;});
        LE_Transform(layer3, option2);

        // 组起来
        var grp = doc.groupItems.add();
        layer3.move(grp, ElementPlacement.INSIDE);
        layer2.move(grp, ElementPlacement.INSIDE);
        layer1.move(grp, ElementPlacement.INSIDE);

        // 删除原始
        item.remove();

        return grp;
    }
    function applyEffectToGroup(group) {
        if (group.clipped) {
            for (var j = 0; j < group.compoundPathItems.length; j++) {
                try {
                    //找到路径的填色，在填色中加入偏移路径。
                    var compoundPath = group.compoundPathItems[j];
                        applyMultiAppearance(compoundPath,fontSize);
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
 * 调用Illustrator进行处理
 */
function blackWInIllustrator(filePath,fontSize) {
    try {
        // 检查BridgeTalk是否可用
        if (typeof BridgeTalk === "undefined") {
            alert("BridgeTalk不可用，无法调用Illustrator");
            return;
        }
        
        // 检查Illustrator是否可用
        if (!BridgeTalk.isRunning("illustrator")) {
            alert("Illustrator未运行，请先启动Illustrator，或启动方式不对请重启Illustrator不要打开任何文件！");
            return;
        }
        var scriptFunction = illustratorBlackWordScript.toString();
        var script = 
            "var filePath = '" + filePath.fsName.replace(/\\/g, "\\\\") + "';\n" +
            scriptFunction + "\n" +
            "illustratorBlackWordScript(filePath,"+fontSize+");";
        
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
        sleepWithEvents(5000);
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
        app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_BLEED
        var placedAsset = page.place(new File(filePath.fsName))[0];
        app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_CONTENT_VISIBLE_LAYERS
        placedAsset.parent.locked = true;
    } catch (err) {
        alert("置入AI文件时出错: " + err.description);
    }
}

// 执行主函数
KTUDoScriptAsUndoable(function() { aiBlackWEffect(); }, "AI黑底雕特效");