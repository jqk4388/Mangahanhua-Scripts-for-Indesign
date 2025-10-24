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
// 执行主函数
KTUDoScriptAsUndoable(function() { aiBlackWEffect(); }, "AI黑底雕特效");