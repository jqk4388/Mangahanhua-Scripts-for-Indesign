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
    // doc.close();
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



var filePath   = "m:\\桌面\\未命名-2_但是但是但_1759687535217.ai"
illustratorRoughenScript(
    filePath,
    0.3,    // 大小 1%
    50,   // 细节 50 /Inch
);
