var version = "1.1";

// 主入口
function main() {
    try {
        // 1. 检查样式数量
        var styleCheck = checkStyleCount();
        if (!styleCheck.ok) {
            alert("字符样式或段落样式数量小于2，请导入样式模板。");
            return;
        }

        // 2. 读取配置
        var config = readConfigFile();

        // 3. 查找所有页面的文本框
        var textFrames = [];
        var doc = app.activeDocument;
        for (var i = 0; i < doc.pages.length; i++) {
            var page = doc.pages[i];
            var frames = getAllTextFrames(page);
            for (var j = 0; j < frames.length; j++) {
            textFrames.push(frames[j]);
            }
        }

        // 4. 提取所有花括号内的字体和字号
        var fontList = [];
        var sizeList = [];
        extractFontSizeList(textFrames, fontList, sizeList);
        if (fontList.length === 0 && sizeList.length === 0) {
            alert("未检测到花括号内的字体或字号内容，脚本退出。");
            return;
        }

        // 5. 构建UI
        var ui = buildUI(fontList, sizeList, config, styleCheck.charStyleNames, styleCheck.paraStyleNames);

        // 6. UI事件绑定
        bindUIEvents(ui, fontList, sizeList, config, textFrames, styleCheck.charStyle, styleCheck.paraStyle);

        // 7. 显示UI
        ui.win.show();

    } catch (e) {
        alert("脚本发生错误：" + e.message);
    }
}

// 检查字符样式和段落样式数量，并输出样式名称列表
function checkStyleCount() {
    var doc = app.activeDocument;
    var charStyleCount = 0;
    var paraStyleCount = 0;
    var charStyleNames = [];
    var paraStyleNames = [];
    var charStyle = [];
    var paraStyle = [];

    // 递归统计字符样式，带分组
    function countCharStyles(styles, groupName) {
        for (var i = 0; i < styles.length; i++) {
            if (styles[i].constructor.name === "CharacterStyleGroup") {
                var gName = styles[i].name;
                countCharStyles(styles[i].characterStyles, gName);
                countCharStyles(styles[i].characterStyleGroups, gName);
            } else if (styles[i].constructor.name === "CharacterStyle") {
                    charStyleCount++;
                    var name = styles[i].name;
                    var CStyle = styles[i];
                    var group = groupName && groupName !== "" ? groupName : "";
                    if (group !== "") {
                        name = name + "(" + group + ")";
                    }
                    charStyleNames.push(name);
                    // 存对象，包含样式和分组名
                    charStyle.push({ style: CStyle, groupName: group });
            }
        }
    }



    // 递归统计段落样式，带分组
    function countParaStyles(styles, groupName) {
        for (var i = 0; i < styles.length; i++) {
            if (styles[i].constructor.name === "ParagraphStyleGroup") {
                var gName = styles[i].name;
                countParaStyles(styles[i].paragraphStyles, gName);
                countParaStyles(styles[i].paragraphStyleGroups, gName);
            } else if (styles[i].constructor.name === "ParagraphStyle") {
                if (styles[i].name !== "[基本段落]" && styles[i].name !== "[无]") {
                    paraStyleCount++;
                    var name = styles[i].name;
                    var PStyle = styles[i];
                    var group = groupName && groupName !== "" ? groupName : "";
                    if (group !== "") {
                        name = name + "(" + group + ")";
                    }
                    paraStyleNames.push(name);
                    paraStyle.push({style: PStyle, groupName: group});
                }
            }
        }
    }

    // 统计组外样式
    countCharStyles(doc.characterStyles, "");
    countCharStyles(doc.characterStyleGroups, "");
    countParaStyles(doc.paragraphStyles, "");
    countParaStyles(doc.paragraphStyleGroups, "");

    return {
        ok: (charStyleCount >= 2 && paraStyleCount >= 2),
        charStyleNames: charStyleNames,
        paraStyleNames: paraStyleNames,
        charStyle: charStyle,
        paraStyle: paraStyle
    };
}

// 读取配置文件
function readConfigFile() {
var dirSeparator = $.os.search(/windows/i) === -1 ? '/' : '\\';
var dataPath = Folder.appData.fsName + dirSeparator + "style_replace";
var dataFolder = new Folder(dataPath);
    var config = {};
    try {
        var folder = Folder(dataPath + dirSeparator);
        if (!folder.exists) {
            folder.create();
        }
        var file = File(dataPath + dirSeparator + "style_replace.ini");
        if (file.exists) {
            file.open("r");
            while (!file.eof) {
                var line = file.readln();
                var eq = line.indexOf("=");
                if (eq > 0) {
                    var key = line.substring(0, eq);
                    var val = line.substring(eq + 1);
                    config[key] = val;
                }
            }
            file.close();
        }
    } catch (e) {
        // 读取失败返回空对象
        config = {};
    }
    return config;
}

// 保存配置文件
function saveConfigFile(config, filePath) {
var dirSeparator = $.os.search(/windows/i) === -1 ? '/' : '\\';
var dataPath = Folder.appData.fsName + dirSeparator + "style_replace";
var dataFolder = new Folder(dataPath);
if (!dataFolder.exists) {
    if (!dataFolder.create()) {
        dataPath = Folder.temp.fsName;
    }
}
    try {
        var file;
        if (filePath) {
            file = File(filePath);
        } else {
            var folder = Folder(dataPath + dirSeparator);
            if (!folder.exists) {
                folder.create();
            }
            file = File(dataPath + dirSeparator + "style_replace.ini");
        }
        file.encoding = "UTF-8";
        file.open("w");
        for (var key in config) {
            if (config.hasOwnProperty(key)) {
                file.writeln(key + "=" + config[key]);
            }
        }
        file.close();
    } catch (e) {
        alert("保存配置文件失败：" + e.message);
    }
}

// 获取页面中所有文本框，包括编组中的文本框
function getAllTextFrames(page) {
    var textFrames = [];

    function collectTextFrames(item) {
        if (item.constructor.name === "TextFrame") {
            textFrames.push(item);
        } else if (item.constructor.name === "Group") {
            for (var i = 0; i < item.allPageItems.length; i++) {
                collectTextFrames(item.allPageItems[i]);
            }
        }
    }

    for (var i = 0; i < page.allPageItems.length; i++) {
        collectTextFrames(page.allPageItems[i]);
    }

    // 过滤掉文字转曲后的路径
    var validFrames = [];
    for (var i = 0; i < textFrames.length; i++) {
        if (textFrames[i].contents !== null && String(textFrames[i].contents).replace(/^\s+|\s+$/g, '') !== "") {
            validFrames.push(textFrames[i]);
        }
    }
    return validFrames;
}

// 提取所有花括号内的字体和字号
function extractFontSizeList(textFrames, fontList, sizeList) {
    try {
        for (var i = 0; i < textFrames.length; i++) {
            var story = textFrames[i].parentStory;
            if (story) {
                var txt = story.contents;
                // 查找{字体：xxx}
                var idx = 0;
                while (true) {
                    var start = txt.indexOf("{字体：", idx);
                    if (start < 0) break;
                    var end = txt.indexOf("}", start);
                    if (end < 0) break;
                    var font = txt.substring(start + 4, end);
                    if (!inArray(fontList, font)) fontList.push(font);
                    idx = end + 1;
                }
                // 查找{字号：xxx}
                idx = 0;
                while (true) {
                    var start2 = txt.indexOf("{字号：", idx);
                    if (start2 < 0) break;
                    var end2 = txt.indexOf("}", start2);
                    if (end2 < 0) break;
                    var size = txt.substring(start2 + 4, end2);
                    if (!inArray(sizeList, size)) sizeList.push(size);
                    idx = end2 + 1;
                }
            }
        }
    } catch (e) {}
}

// 判断数组中是否存在某元素
function inArray(arr, val) {
    for (var i = 0; i < arr.length; i++) {
        if (arr[i] == val) return true;
    }
    return false;
}

// 构建UI
function buildUI(fontList, sizeList, config, charStyleNames, paraStyleNames) {
    var win = new Window("dialog", "样式匹配脚本"+version, undefined, {closeButton: true});
    win.preferredSize = [800, 600];

    // 主分组
    var mainGroup = win.add("group");
    mainGroup.orientation = "row";
    mainGroup.maximumSize.height = 600;
    mainGroup.alignment = "left";

    // 左侧panel
    var leftPanel = mainGroup.add("panel", undefined, "字体-字符样式");
    leftPanel.orientation = "column";
    leftPanel.alignment = ['left', 'top'];
    leftPanel.alignChildren = "fill";
    leftPanel.preferredSize = [250, 550];    
    var leftAddBtn = leftPanel.add("button", undefined, "+");
    var leftexcel = leftPanel.add ("panel", undefined, "");
    leftexcel.orientation = "row";
    var leftListPanel = leftexcel.add("panel");
    leftListPanel.preferredSize = [230, 500];
    leftListPanel.alignChildren = "top";
    leftListPanel.alignment = ['left', 'top'];
    leftListPanel.orientation = "column";
    leftListPanel.margins = [0,0,0,0];
    leftListPanel.spacing = 2;

    // 右侧panel
    var rightPanel = mainGroup.add("panel", undefined, "字号-段落样式");
    rightPanel.orientation = "column";
    rightPanel.alignment = ['right', 'top'];
    rightPanel.alignChildren = "fill";
    rightPanel.preferredSize = [250, 550];
    var rightAddBtn = rightPanel.add("button", undefined, "+");
    var rightexcel = rightPanel.add ("panel", undefined, "");
    rightexcel.orientation = "row";
    var rightListPanel = rightexcel.add("panel");
    rightListPanel.preferredSize = [230, 500];
    rightListPanel.alignChildren = "top";
    rightListPanel.alignment = ['right', 'top'];
    rightListPanel.orientation = "column";
    rightListPanel.margins = [0,0,0,0];
    rightListPanel.spacing = 2;

    // 行数据存储
    var leftRows = [];
    var rightRows = [];

    // 添加左侧一行
    function addLeftRow(fontName, styleName) {
        var g = leftListPanel.add("group");
        g.orientation = "row";
        g.alignChildren = "center";
        g.preferredSize = [220, 24];
        var cb = g.add("checkbox", undefined, "");
        cb.value = true;
        var edit = g.add("edittext", undefined, fontName || "");
        edit.preferredSize = [180, 22];
        var label = g.add("statictext", undefined, "匹配");
        var dropdown = g.add("dropdownlist", undefined, charStyleNames);
        dropdown.preferredSize = [150, 22];
        // 按config匹配
        if (styleName && charStyleNames.length > 0) {
            for (var i = 0; i < charStyleNames.length; i++) {
                if (charStyleNames[i] == styleName) {
                    dropdown.selection = i;
                    break;
                }
            }
        } else {
            dropdown.selection = null;
        }
        leftRows.push({cb: cb, edit: edit, dropdown: dropdown, group: g});
        leftListPanel.layout.layout(true);
    }
    // 添加右侧一行
    function addRightRow(sizeName, styleName) {
        var g = rightListPanel.add("group");
        g.orientation = "row";
        g.alignChildren = "center";
        g.preferredSize = [220, 24];
        var cb = g.add("checkbox", undefined, "");
        cb.value = true;
        var edit = g.add("edittext", undefined, sizeName || "");
        edit.preferredSize = [70, 22];
        var label = g.add("statictext", undefined, "匹配");
        var dropdown = g.add("dropdownlist", undefined, paraStyleNames);
        dropdown.preferredSize = [80, 22];
        // 自动匹配最接近的段落样式
        if (sizeName && paraStyleNames.length > 0) {
            var num = parseFloat(sizeName);
            var minDiff = 99999, selIdx = -1;
            for (var i = 0; i < paraStyleNames.length; i++) {
                var psNum = parseFloat(paraStyleNames[i]);
                if (!isNaN(psNum) && !isNaN(num)) {
                    var diff = Math.abs(psNum - num);
                    if (diff < minDiff) {
                        minDiff = diff;
                        selIdx = i;
                    }
                }
            }
            if (selIdx >= 0) dropdown.selection = selIdx;
        }
        rightRows.push({cb: cb, edit: edit, dropdown: dropdown, group: g});
        rightListPanel.layout.layout(true);
    }

    // 填充左侧
    for (var i = 0; i < fontList.length; i++) {
        var styleName = config["font_style_" + i] || null;
        addLeftRow(fontList[i], styleName);
    }
    // 在填充左侧列表后添加如下代码
    var leftbar = leftexcel.add("scrollbar", [0,0,20,700]);
    // 计算内容总高度和可见区域高度
    var totalHeight = leftRows.length * 24; // 每行高24
    var visibleHeight = 500; // leftListPanel的preferredSize.height
    // 设置滚动条最大值为(总高度-可见区域)
    leftbar.maxvalue = Math.max(0, totalHeight - visibleHeight);
    leftbar.stepdelta = 24; // 设置为单行高度
    leftbar.jumpdelta = visibleHeight; // 设置翻页距离为可见区域高度
    leftbar.onChanging = function () {
        leftListPanel.location.y = -1 * this.value;
    }
    leftbar.alignment = ['right', 'top'];
    // 填充右侧
    for (var j = 0; j < sizeList.length; j++) {
        var styleName = config["size_style_" + j] || null;
        addRightRow(sizeList[j], styleName);
    }
    var rightbar = rightexcel.add("scrollbar", [0,0,20,700]);
    var totalHeightR = rightRows.length * 24;
    var visibleHeightR = 500;
    rightbar.maxvalue = Math.max(0, totalHeightR - visibleHeightR);
    rightbar.stepdelta = 24;
    rightbar.jumpdelta = visibleHeightR;
    rightbar.onChanging = function () {
        rightListPanel.location.y = -1 * this.value;
    }
    rightbar.alignment = ['right', 'top'];

    // 按钮区
    var btnGroup = mainGroup.add("group");
    btnGroup.orientation = "column";
    btnGroup.alignment = ['right', 'top'];
    btnGroup.alignChildren = "fill";
    var okBtn = btnGroup.add("button", undefined, "确定");
    var cancelBtn = btnGroup.add("button", undefined, "取消");
    var importBtn = btnGroup.add("button", undefined, "导入配置");
    var exportBtn = btnGroup.add("button", undefined, "导出配置");

    // 加号按钮事件
    leftAddBtn.onClick = function () {
        addLeftRow("", "");
        leftListPanel.layout.layout(true);
        win.layout.layout(true);
        win.center();
    };
    rightAddBtn.onClick = function () {
        addRightRow("", "");
        rightListPanel.layout.layout(true);
        win.layout.layout(true);
        win.center();
    };

    // 返回UI对象
    return {
        win: win,
        leftRows: leftRows,
        rightRows: rightRows,
        leftAddBtn: leftAddBtn,
        rightAddBtn: rightAddBtn,
        ListListPanel: leftListPanel,
        rightListPanel: rightListPanel,
        leftbar: leftbar,
        rightbar: rightbar,
        okBtn: okBtn,
        cancelBtn: cancelBtn,
        importBtn: importBtn,
        exportBtn: exportBtn,
        charStyleNames: charStyleNames,
        paraStyleNames: paraStyleNames
    };
}

// 绑定UI事件
function bindUIEvents(ui, fontList, sizeList, config, textFrames, charStyle, paraStyle) {
    // 确定按钮
    ui.okBtn.onClick = function () {
        // 保存配置
        for (var i = 0; i < ui.leftRows.length; i++) {
            var row = ui.leftRows[i];
            if (row.cb.value && row.edit.text) {
                config["font_" + i] = row.edit.text;
                config["font_style_" + i] = row.dropdown.selection ? row.dropdown.selection.text : "";
            }
        }
        for (var j = 0; j < ui.rightRows.length; j++) {
            var row = ui.rightRows[j];
            if (row.cb.value && row.edit.text) {
                config["size_" + j] = row.edit.text;
                config["size_style_" + j] = row.dropdown.selection ? row.dropdown.selection.text : "";
            }
        }
        saveConfigFile(config);
        // 应用样式
        applyStyles(textFrames, config, charStyle, paraStyle);
        //用Grep查找清除花括号{}中的内容
        ClearBrackets(textFrames);
        
        ui.win.close();
    };
    // 取消按钮
    ui.cancelBtn.onClick = function () {
        ui.win.close();
    };
    // 导入配置
    ui.importBtn.onClick = function () {
        var file = File.openDialog("选择配置文件", "*.ini");
        if (file && file.exists) {
            var newConfig = {};
            file.open("r");
            while (!file.eof) {
                var line = file.readln();
                var eq = line.indexOf("=");
                if (eq > 0) {
                    var key = line.substring(0, eq);
                    var val = line.substring(eq + 1);
                    newConfig[key] = val;
                }
            }
            file.close();
            // 更新UI
            updateUIFromConfig(ui, newConfig);
        }
    };
    // 导出配置
    ui.exportBtn.onClick = function () {
        var file = File.saveDialog("导出配置为", "ini文件:*.ini");
        if (file) {
            var exportConfig = {};
            for (var i = 0; i < ui.leftRows.length; i++) {
                var row = ui.leftRows[i];
                if (row.cb.value && row.edit.text) {
                    exportConfig["font_" + i] = row.edit.text;
                    exportConfig["font_style_" + i] = row.dropdown.selection ? row.dropdown.selection.text : "";
                }
            }
            for (var j = 0; j < ui.rightRows.length; j++) {
                var row = ui.rightRows[j];
                if (row.cb.value && row.edit.text) {
                    exportConfig["size_" + j] = row.edit.text;
                    exportConfig["size_style_" + j] = row.dropdown.selection ? row.dropdown.selection.text : "";
                }
            }
            saveConfigFile(exportConfig, file.fsName);
        }
    };
}

// 根据配置更新UI
function updateUIFromConfig(ui, config) {
    try {
        // 清空
        for (var i = ui.leftRows.length - 1; i >= 0; i--) {
            ui.leftRows[i].group.parent.remove(ui.leftRows[i].group);
        }
        ui.leftRows.length = 0;
        for (var j = ui.rightRows.length - 1; j >= 0; j--) {
            ui.rightRows[j].group.parent.remove(ui.rightRows[j].group);
        }
        ui.rightRows.length = 0;
        // 恢复
        var i = 0;
        while (config["font_" + i]) {
            var fontName = config["font_" + i];
            var styleName = config["font_style_" + i];
            var row = null;
            // add row
            if (typeof ui.leftAddBtn.onClick === "function") {
                ui.leftAddBtn.onClick();
                row = ui.leftRows[ui.leftRows.length - 1];
                row.edit.text = fontName;
                if (styleName && row.dropdown.items.length > 0) {
                    for (var k = 0; k < row.dropdown.items.length; k++) {
                        if (row.dropdown.items[k].text == styleName) {
                            row.dropdown.selection = k;
                            break;
                        }
                    }
                }
            }
            i++;
        }
        updateScrollbar(ui.leftListPanel, ui.leftRows, ui.leftbar, 500);
        var j = 0;
        while (config["size_" + j]) {
            var sizeName = config["size_" + j];
            var styleName = config["size_style_" + j];
            var row = null;
            if (typeof ui.rightAddBtn.onClick === "function") {
                ui.rightAddBtn.onClick();
                row = ui.rightRows[ui.rightRows.length - 1];
                row.edit.text = sizeName;
                if (styleName && row.dropdown.items.length > 0) {
                    for (var k = 0; k < row.dropdown.items.length; k++) {
                        if (row.dropdown.items[k].text == styleName) {
                            row.dropdown.selection = k;
                            break;
                        }
                    }
                }
            }
            j++;
        }
        updateScrollbar(ui.rightListPanel, ui.rightRows, ui.rightbar, 500);
    } catch (e) {}
}

// 应用样式
function applyStyles(textFrames, config, charStyleNames, paraStyleNames) {
    try {
        var doc = app.activeDocument;
        var charStyles = charStyleNames;
        var paraStyles = paraStyleNames;
        // 遍历文本框
        for (var i = 0; i < textFrames.length; i++) {
            var story = textFrames[i].parentStory;
            if (!story) continue;
            var txt = story.contents;

            // 字体-字符样式
            var fontIdx = 0;
            while (config["font_" + fontIdx]) {
            var fontName = config["font_" + fontIdx];
            var styleName = config["font_style_" + fontIdx];
            var cs = getStyleByName(charStyles, styleName);
            if (cs) {
                // 如果文本中包含{字体：xxx}，则对整个文本框应用字符样式
                if (txt.indexOf(fontName) !== -1) {
                try {
                    textFrames[i].texts[0].appliedCharacterStyle = cs;
                } catch (e) {}
                }
            }
            fontIdx++;
            }

            // 字号-段落样式
            var sizeIdx = 0;
            while (config["size_" + sizeIdx]) {
            var sizeName = config["size_" + sizeIdx];
            var styleName = config["size_style_" + sizeIdx];
            var ps = getStyleByName(paraStyles, styleName);
            // 若找不到则自动匹配
            if (!ps) ps = getClosestParaStyle(paraStyles, sizeName);
            if (ps) {
                // 如果文本中包含{字号：xxx}，则对整个文本框应用段落样式
                if (txt.indexOf("{字号：" + sizeName + "}") !== -1) {
                try {
                    textFrames[i].texts[0].appliedParagraphStyle = ps;
                } catch (e) {}
                }
            }
            sizeIdx++;
            }
        }
    } catch (e) {
        alert("应用样式时发生错误：" + e.message);
    }
}

// 根据名字获取样式
function getStyleByName(styles, name) {
    for (var i = 0; i < styles.length; i++) {
        var styleName = styles[i]['style']['name']+'('+ styles[i]['groupName']+')';
        if (styleName == name) return styles[i]['style'];
    }
    return null;
}

// 获取最接近的段落样式
function getClosestParaStyle(paraStyles, sizeStr) {
    var minDiff = 99999;
    var best = null;
    var target = parseFloat(sizeStr);
    for (var i = 0; i < paraStyles.length; i++) {
        var psName = paraStyles[i].name;
        var psNum = parseFloat(psName);
        if (!isNaN(psNum)) {
            var diff = Math.abs(psNum - target);
            if (diff < minDiff) {
                minDiff = diff;
                best = paraStyles[i];
            }
        }
    }
    return best;
}

function ClearBrackets(textFrames) {
    try {
        // 保存当前查找/替换设置
        var savedFindPrefs = app.findGrepPreferences.properties;
        var savedChangePrefs = app.changeGrepPreferences.properties;

        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;

        app.findGrepPreferences.findWhat = "\\{.*?\\}";
        app.changeGrepPreferences.changeTo = "";

        for (var i = 0; i < textFrames.length; i++) {
            var story = textFrames[i].parentStory;
            if (!story) continue;
            story.changeGrep();
        }

        // 恢复查找/替换设置
        app.findGrepPreferences.properties = savedFindPrefs;
        app.changeGrepPreferences.properties = savedChangePrefs;
    } catch (e) {}
}

//更新滚动条长度
function updateScrollbar(panel, rows, scrollbar, visibleHeight) {
    var totalHeight = rows.length * 24; // 每行高24
    scrollbar.maxvalue = Math.max(0, totalHeight - visibleHeight);
    scrollbar.stepdelta = 24;
    scrollbar.jumpdelta = visibleHeight;
    panel.location.y = 0; // 重置面板位置
    scrollbar.value = 0; // 重置滚动条位置
}

// 启动脚本
main();