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
        
        // 获取文本框所在的页面
        var page = textFrames[0].parentPage;
        
        // 记录当前图层状态
        var layerStates = saveLayerStates(page);
        
        // 隐藏不需要的元素
        hideUnselectedElements(textFrames);
        
        // 处理每个选中的文本框
        for (var j = 0; j < textFrames.length; j++) {
            var tempFrame = processTextFrame(textFrames[j], docFolder, j)
        }
        // 还原图层状态
        restoreLayerStates(layerStates, tempFrame, page);

    } catch (err) {
        alert("发生错误，请先保存文件！ \n" + err.description);
        // 尝试还原图层状态
        try {
            restoreLayerStates(layerStates, null, page);
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
            alert("Illustrator未运行，请先启动Illustrator，或启动方式不对请重启Illustrator不要打开任何文件！");
            return;
        }
        
        // 获取illustratorRoughenScript函数的字符串表示
        var scriptFunction = illustratorRoughenScript.toString();
        
        // 构造发送给Illustrator的完整脚本
        var script = 
            "var filePath = '" + filePath.fsName.replace(/\\/g, "\\\\") + "';\n" +
            scriptFunction + "\n" +
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

// 执行主函数
//弹出提示框，点击确定才继续，取消退出脚本
var userConfirm = confirm("在超过30页的文档中运行此脚本会非常慢，请把要做的特效文字复制到新文档！\n即将对选中的文本框应用AI粗糙化特效。\n请确保已保存当前文档，并且Illustrator已启动且未打开任何文件。\n点击“确定”继续，或“取消”退出脚本。");
if (userConfirm) {
KTUDoScriptAsUndoable(function() { aiRoughenEffect(); }, "AI粗糙化特效");
} else {
    // 用户取消操作，退出脚本
    alert("操作已取消，脚本退出。");
}