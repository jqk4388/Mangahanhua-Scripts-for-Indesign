//#target indesign
//作者：几千块
//时间：20250402
//描述：将选中对象的黑色和白色互换，包括文本框内的文字颜色、描边颜色、填充颜色、外发光效果等
//版本：1.0.0
/**
 * 主函数 - 处理用户选择对象
 */
function main() {
    try {
        // 检查是否有选中对象
        if (!app.documents.length || !app.selection || !app.selection.length) {
            throw new Error("请先选中需要处理的对象");
        }

        // 遍历所有选中对象
        for (var i = 0; i < app.selection.length; i++) {
            processItem(app.selection[i]);
        }

        // alert("处理完成，共处理 " + app.selection.length + " 个对象");
    } catch (e) {
        alert("错误: " + e.message);
    }
}

/**
 * 递归处理对象
 * @param {Object} item - InDesign对象
 */
function processItem(item) {
    try {
        // 处理组对象
        if (item.constructor.name === "Group") {
            for (var j = 0; j < item.pageItems.length; j++) {
                processItem(item.pageItems[j]);
            }
            return;
        }

        // 处理文本框
        var a =item.constructor.name;
        if (item.constructor.name === "TextFrame" || item.constructor.name === "PageItem") {
            processTextFrame(item);
            return;
        }

        // 处理普通对象
        processNormalObject(item);
    } catch (e) {
        // 错误继续执行不中断
        $.writeln("处理对象时出错: " + e.message);
    }
}

/**
 * 处理普通对象
 * @param {Object} obj - 页面对象
 */
function processNormalObject(obj) {
    // 处理外发光效果
    processEffects(obj);

    // 处理填充颜色
    if (obj.fillColor && isValidColor(obj.fillColor)) {
        obj.fillColor = getReversedColor(obj.fillColor);
    }

    // 处理描边颜色
    if (obj.strokeColor && isValidColor(obj.strokeColor)) {
        obj.strokeColor = getReversedColor(obj.strokeColor);
    }
}

/**
 * 处理文本框特殊逻辑
 * @param {TextFrame} textFrame - 文本框对象
 */
function processTextFrame(textFrame) {
    // 处理文本框自身属性
    processNormalObject(textFrame);

    // 处理文本内容颜色
    var textRanges = textFrame.texts;
    for (var i = 0; i < textRanges.length; i++) {
        var textRange = textRanges[i];
        processTextColor(textRange);
    }
}

/**
 * 处理文本颜色
 * @param {Text} textRange - 文本范围
 */
function processTextColor(textRange) {
    // 文本填充色
    if (textRange.fillColor && isValidColor(textRange.fillColor)) {
        textRange.fillColor = getReversedColor(textRange.fillColor);
    }

    // 文本描边色
    if (textRange.strokeColor && isValidColor(textRange.strokeColor)) {
        textRange.strokeColor = getReversedColor(textRange.strokeColor);
    }
}

/**
 * 处理对象所有外发光效果
 * @param {Object} obj
 */
function processEffects(obj) {
    try {

        // 1.处理对象自身效果
        if (obj.transparencySettings && obj.transparencySettings.isValid) {
            
            var objEffects = obj.transparencySettings;
            if (objEffects['outerGlowSettings']['applied']) {
                var outerGlow = objEffects.outerGlowSettings;
                if (outerGlow.effectColor && isValidColor(outerGlow.effectColor)) {
                    outerGlow.effectColor = getReversedColor(outerGlow.effectColor);
                }
            }
        }

        // 2.处理填充效果
        if (obj.fillTransparencySettings && obj.fillTransparencySettings.isValid) {
            var fillEffects = obj.fillTransparencySettings;
            if (fillEffects['outerGlowSettings']['applied']) {
                var outerGlow = fillEffects.outerGlowSettings;
                if (outerGlow.effectColor && isValidColor(outerGlow.effectColor)) {
                    outerGlow.effectColor = getReversedColor(outerGlow.effectColor);
                }
            }
        }

        // 3.处理描边效果
        if (obj.strokeTransparencySettings && obj.strokeTransparencySettings.isValid) {
            var strokeEffects = obj.strokeTransparencySettings;
            if (strokeEffects['outerGlowSettings']['applied']) {
                var outerGlow = strokeEffects.outerGlowSettings;
                if (outerGlow.effectColor && isValidColor(outerGlow.effectColor)) {
                    outerGlow.effectColor = getReversedColor(outerGlow.effectColor);
                }
            }
        }

        // 4.处理内容效果
        if (obj.contentTransparencySettings && obj.contentTransparencySettings.isValid) {
            var contentEffects = obj.contentTransparencySettings;
            if (contentEffects['outerGlowSettings']['applied']) {
                var outerGlow = contentEffects.outerGlowSettings;
                if (outerGlow.effectColor && isValidColor(outerGlow.effectColor)) {
                    outerGlow.effectColor = getReversedColor(outerGlow.effectColor);
                }
            }
        }
    } catch (e) {
        $.writeln("处理外发光效果时出错: " + e.message);
    }
}


/**
 * 验证是否为黑白颜色
 * @param {Color} color - 颜色对象
 * @returns {Boolean}
 */
function isValidColor(color) {
    return color.name === "Black" || color.name === "Paper";
}

/**
 * 获取反转颜色
 * @param {Color} color - 原始颜色
 * @returns {Color}
 */
function getReversedColor(color) {
    return color.name === "Black" 
        ? app.activeDocument.colors.item("Paper") 
        : app.activeDocument.colors.item("Black");
}

// 执行主程序
main();