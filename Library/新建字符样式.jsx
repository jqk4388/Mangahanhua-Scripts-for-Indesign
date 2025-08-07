/**
 * 获取所有日文/繁体中文/简体中文字体
 * @returns {Array} 字体对象数组
 */
function getCJKFonts() {
    var fonts = app.fonts;
    var cjkFonts = [];
    // 汉字和假名的正则表达式
    var cjkRegex = /[\u4E00-\u9FFF\u3040-\u30FF]/;
    for (var i = 0; i < fonts.length; i++) {
        var font = fonts[i];
        // 跳过无效字体,当font['postscriptName']有值才继续
        if (!font['postscriptName']) continue;
        var fullName = font.fullNameNative;
        // 排除包含“华文”的字体
        if (cjkRegex.test(fullName) && fullName.indexOf("华文") === -1) {
            cjkFonts.push(font);
            $.writeln(fullName);
        }
    }
    return cjkFonts;
}

/**
 * 按厂牌和家族分组字体
 * @param {Array} fonts 字体对象数组
 * @returns {Object} {厂牌名: {家族名: [字体对象, ...]}}
 */
function groupFontsByFoundryAndFamily(fonts) {
    var groups = {};
    // 常见风格词列表
    var styleWords = [
        "Regular", "Bold", "Medium", "Light", "Heavy", "Book", "Thin", "Black", "Italic", "Oblique", "Semibold", "Extrabold", "Ultralight", "Condensed", "Expanded" , "Demi","DemiBold", "Extrabold", "Ultrabold", "Extralight", "Ultralight", "Extra", "Ultra", "Narrow", "Wide", "Narrow", "Wide", "Narrow", "Wide", "Narrow", "Wide", "Narrow", "Wide", "Narrow", "Wide", "Narrow", "Wide", "Narrow", "Wide", "Narrow", "Wide", "Narrow", "Wide", "粗体", "细体", "斜体", "半粗体", "半细体", "超细体", "超粗体", "超细体", "超粗体", "超细体" ,"GBK","mono" ,"serif","W3","W5","W7","W9","W11","W12","W13","W14","W15","W16","W17","W18","W19","W20","E","U","H","M","B","D","R","UB","EB","L","EL","T","UL","DB","-"," "
    ];
    for (var i = 0; i < fonts.length; i++) {
        var font = fonts[i];
        var name = font.fullNameNative;
        var foundry = name.substr(0, 2);
        var rest = name.substr(2);
        var family = "";
        // 去除前导空格
        while (rest.length > 0 && rest[0] === " ") {
            rest = rest.substr(1);
        }
        // 截取到第一个空格或风格词前
        var splitIndex = rest.length;
        for (var j = 0; j < styleWords.length; j++) {
            var idx = rest.indexOf(styleWords[j]);
            if (idx !== -1 && idx < splitIndex) {
                splitIndex = idx;
            }
        }
        var spaceIdx = rest.indexOf(" ");
        if (spaceIdx !== -1 && spaceIdx < splitIndex) {
            splitIndex = spaceIdx;
        }
        family = rest.substr(0, splitIndex);
        // 如果家族名为空，则归为厂牌下
        if (!groups[foundry]) groups[foundry] = {};
        if (family) {
            if (!groups[foundry][family]) groups[foundry][family] = [];
            groups[foundry][family].push(font);
        } else {
            if (!groups[foundry]["__nofamily__"]) groups[foundry]["__nofamily__"] = [];
            groups[foundry]["__nofamily__"].push(font);
        }
    }
    return groups;
}

/**
 * 获取字体厂牌名
 * - 如果第一个字符是中文，取前两个字符
 * - 如果第一个字符是英文，取第一个单词
 * @param {Font} font 字体对象
 * @returns {String} 厂牌名
 */
function getFontFoundry(font) {
    var name = font.name;
    if (!name || name.length === 0) return "";
    var firstChar = name[0];
    // 判断是否为中文字符
    if (/[\u4e00-\u9fa5]/.test(firstChar)) {
        return name.substr(0, 2);
    } else if (/[a-zA-Z]/.test(firstChar)) {
        var match = name.match(/^[a-zA-Z]+/);
        return match ? match[0] : firstChar;
    } else {
        return firstChar;
    }
}

/**
 * 获取字重名
 * @param {Font} font 字体对象
 * @returns {String} 字重名
 */
function getFontWeight(font) {
    // 使用font.styleName属性
    return font.fontStyleName;
}

/**
 * 新建字符样式
 * @param {Document} doc 当前文档
 * @param {String} styleName 样式名
 * @param {Font} font 字体对象
 * @param {Object} parentGroup 父组（可选）
 */
function createCharacterStyle(doc, styleName, font, parentGroup) {
    try {
        var baseStyle = doc.characterStyles.itemByName("[无]");
        var newStyle;
        if (parentGroup) {
            newStyle = parentGroup.characterStyles.add({ name: styleName, basedOn: baseStyle });
            $.writeln("新建样式成功：" + styleName + " 在组 " + parentGroup.name);
        } else {
            newStyle = doc.characterStyles.add({ name: styleName, basedOn: baseStyle });
            $.writeln("新建样式成功："+ styleName);
        }
        newStyle.appliedFont = font;
        newStyle.fontStyle = getFontWeight(font);
    } catch (e) {
        // 样式已存在或其他错误，忽略
    }
}

/**
 * 新建字符样式组
 * @param {Document} doc 当前文档
 * @param {String} groupName 组名
 * @returns {Object} 样式组对象
 */
function createCharacterStyleGroup(doc, groupName) {
    try {
        return doc.characterStyleGroups.add({ name: groupName });        
    } catch (e) {
        // 组已存在或其他错误，返回已存在的组
        return doc.characterStyleGroups.itemByName(groupName);
    }
}

/**
 * 主流程：遍历字体并新建样式
 */
function main() {
    if (app.documents.length === 0) {
        alert("请先打开一个文档。");
        return;
    }
    var doc = app.activeDocument;
    var cjkFonts = getCJKFonts();
    var fontGroups = groupFontsByFoundryAndFamily(cjkFonts);

    var debugMsg = "分组情况：\n";
    for (var foundry in fontGroups) {
        debugMsg += "厂牌[" + foundry + "]\n";
        for (var family in fontGroups[foundry]) {
            debugMsg += "  家族[" + family + "] 字体数：" + fontGroups[foundry][family].length + "\n";
        }
    }
    $.writeln(debugMsg);

    var styleCount = 0;
    for (var foundry in fontGroups) {
        var foundryGroup = createCharacterStyleGroup(doc, foundry);
        $.writeln("新建厂牌组：" + foundry);
        for (var family in fontGroups[foundry]) {
            var fonts = fontGroups[foundry][family];
            // 家族字体数量大于2才新建家族组，否则直接放厂牌组下
            if (family !== "__nofamily__" && fonts.length > 2) {
                var familyGroup = createCharacterStyleGroup(foundryGroup, family);
                $.writeln("  新建家族组：" + family);
                for (var i = 0; i < fonts.length; i++) {
                    var font = fonts[i];
                    var styleName = font.fullNameNative;
                    createCharacterStyle(doc, styleName, font, familyGroup);
                    styleCount++;
                }
            } else {
                // 不新建家族组，直接放在厂牌组下
                for (var i = 0; i < fonts.length; i++) {
                    var font = fonts[i];
                    var styleName = font.fullNameNative;
                    createCharacterStyle(doc, styleName, font, foundryGroup);
                    styleCount++;
                }
            }
        }
    }
    $.writeln("总共新建样式数：" + styleCount);
}

// 执行主流程
try {
    main();
} catch (e) {
    alert("脚本执行出错：" + e);
}
