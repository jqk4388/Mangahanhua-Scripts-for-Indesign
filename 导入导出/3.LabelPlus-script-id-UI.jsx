// LP翻译稿处理工具
// 该脚本用于处理LP翻译稿，将文本插入到InDesign文档中，并根据分组应用样式
// 作者：几千块
// 日期：202511123
var version = "2.4";
#include "../Library/KTUlib.jsx"
// 声明全局变量
var totalPages = 0;
var doc = app.activeDocument;
var selectedPages =[];// 保存选中的页面
var startPage = 1;// 设置起始页码
var singleLineMode = false; // 默认单行断句模式
var multiLineRadio = true; // 默认多行断句模式
var pageOffset = 0; // 页码偏移，允许正/负偏移，默认0
var matchByNumber = false; // 按页码匹配选择的文本
var fromStartoEnd = true; // 默认从头导入所有文本
var replacements = {
    "！？": "?!",
    "！": "!",
    "？": "?",
    "——": "—",
    "—？": "—?",
    "…？": "…?",
    "~": "～",
    "诶": "欸",
    "干嘛": "干吗",
    "混蛋": "浑蛋",
    "叮铃": "丁零",
    "":""

};//替换文本

var objectStyleMatchText = [
    { label: "默认匹配", checked: true },
    { label: "*", checked: false },
    { label: "※", checked: false },
    { label: "＊", checked: false },
    { label: "", checked: false },
    { label: "", checked: false },
    { label: "", checked: false },
    { label: "", checked: false },
    { label: "", checked: false },
    { label: "", checked: false },
    { label: "", checked: false },
];//匹配对象样式
var styleRules = [];      // 样式匹配规则

// 读取文本基础，输出行和内容
function readAndParseTxtFileBase(filePath,replacements) {
    var file = new File(filePath);
    if (!file.exists) {
        return { lines: [], content: "" };
    }
    file.open("r");
    var content = file.read();
    replacements = $.global['replacements']
    content = replaceMultipleWords(content, replacements);
    file.close();

    var lines = content.split("\n");
    return { lines: lines, content: content };
}
// 解析文本行，收集台词
function parseLines(lines, isSingleLineMode) {
    var entries = [];
    totalPages = 0; // 重置页面总数

    for (var i = 0; i < lines.length; i++) {
        // 匹配页面名
        var pageMatch = lines[i].match(/>>>>>>>>\[(.*)\]<<<<<<<<$/);
        if (pageMatch) {
            var pageName = pageMatch[1];
            var pageNameArr = [pageName];
            totalPages++; // 每遇到一个新页面时，页面总数增加
            continue;
        }

        // 匹配台词编号、位置和分组信息
        var textMatch = lines[i].match(/----------------\[(\d+)\]----------------\[(.*?)\,(.*?)\,(.*?)\]/);
        if (textMatch) {
            var pageNumber = parseInt(textMatch[1]);
            var baseX = parseFloat(textMatch[2]);
            var baseY = parseFloat(textMatch[3]);
            var group = parseInt(textMatch[4]);

            // 收集文本内容，直到遇到下一个编号或下一页
            var textContent = [];
            for (var j = i + 1; j < lines.length; j++) {
                if (lines[j].match(/----------------\[\d+\]----------------/) || lines[j].match(/>>>>>>>>\[(.*)\]<<<<<<<<$/)) {
                    break;
                }
                var line = lines[j];
                if (line) {
                    textContent.push(line);
                }
            }

            // 根据模式处理文本内容
            if (isSingleLineMode || textContent.length === 1) {
                try {
                    //从textContent[0]中提取大括号{}内的内容，包括大括号本身，包括多个大括号连用
                        var braceMatch="";
                        braceMatch = textContent[0].match(/\{.*?\}/g);
                } catch (error) {
                    
                }
                // 单行模式拆分成多个文本框
                for (var k = 0; k < textContent.length; k++) {
                    if (braceMatch){
                        if (k==0) {
                            var entry = {
                                pageImage: extractPageNumbers(pageNameArr)[0],
                                pageNumber: pageNumber,
                                position: [baseX - k * 0.03, baseY + k * 0.03],
                                group: group,
                                text: textContent[k]
                            };
                            entries.push(entry);
                        }                    
                        if (k>0) {
                            var entry = {
                                pageImage: extractPageNumbers(pageNameArr)[0],
                                pageNumber: pageNumber,
                                position: [baseX - k * 0.03, baseY + k * 0.03],
                                group: group,
                                text: braceMatch + textContent[k]
                            };
                            entries.push(entry);}
                    }else {
                        var entry = {
                            pageImage: extractPageNumbers(pageNameArr)[0],
                            pageNumber: pageNumber,
                            position: [baseX - k * 0.03, baseY + k * 0.03],
                            group: group,
                            text: textContent[k]
                        };
                        entries.push(entry);
                    }
                }
            } else {
                // 多行模式合并成一个文本框
                var entry = {
                    pageImage: extractPageNumbers(pageNameArr)[0],
                    pageNumber: pageNumber,
                    position: [baseX, baseY],
                    group: group,
                    text: textContent.join("\n")
                };
                entries.push(entry);
            }

            i += textContent.length; // 更新索引以跳过已处理的文本行
        }
    }

    return entries;
}
// 读取和解析文本文件
function readAndParseTxtFilelines(filePath) {
    var result = readAndParseTxtFileBase(filePath);
    var lines = result.lines;
    return parseLines(lines, false);
}
// 读取和解析文本文件（单行模式）
function readAndParseTxtFileoneline(filePath) {
    var result = readAndParseTxtFileBase(filePath);
    var lines = result.lines;
    return parseLines(lines, true); 
}
// 从文件名提取页码
function extractPageNumbers(pageNames) {
    var pageNumbers = [];
    var pageNumberPattern = /(\d+)(?=[^\d]*$)/; // 匹配末尾数字的正则

    for (var i = 0; i < pageNames.length; i++) {
        var name = pageNames[i];
        var matches = name.match(pageNumberPattern);
        var num = 0;

        // 优先取末尾数字
        if (matches) {
            num = parseInt(matches[0], 10);
        } 
        // 如果没有末尾数字，寻找最长数字段
        else {
            var allNumbers = name.match(/\d+/g) || [];
            if (allNumbers.length > 0) {
                // 按数字长度降序排序
                allNumbers.sort(function(a, b) {
                    return b.length - a.length;
                });
                num = parseInt(allNumbers[0], 10);
            }
        }

        if (!isNaN(num)) {
            pageNumbers.push(num);
        }
    }

    return pageNumbers;
}
// 根据列表替换多个词
function replaceMultipleWords(text, replacements) {
    // 遍历替换字典中的每个键值对，并替换对应的词
    for (var key in replacements) {
        if (replacements.hasOwnProperty(key)) {
            var regex = new RegExp(key, 'g'); // 使用正则表达式全局替换
            text = text.replace(regex, replacements[key]);
        }
    }

    // 删除末尾的换行符（包括多个换行符）
    while (text.length > 0 && (text.substr(text.length - 1) === "\n" || text.substr(text.length - 1) === "\r")) {
        text = text.substr(0, text.length - 1);
    }

    // 检测文本最后一个字符是否为双引号，如果是则删除
    if (text.length > 0 && text.substr(text.length - 1) === "\"") {
        text = text.substr(0, text.length - 1);
    }

    return text;
}
// 确保页面数足够
function ensureEnoughPages(requiredPageCount) {
    var currentPageCount = doc.pages.length;
    
    if (currentPageCount < requiredPageCount) {
        var pagesToAdd = requiredPageCount - currentPageCount;
        for (var i = 0; i < pagesToAdd; i++) {
            doc.pages.add();
        }
    }
}
// 将百分比坐标转换为绝对坐标
function convertPercentToAbsoluteCoordinates(page, xPercent, yPercent) {
    var pageWidth = page.bounds[3] - page.bounds[1]; // 页面宽度
    var pageHeight = page.bounds[2] - page.bounds[0]; // 页面高度
    var x = pageWidth * parseFloat(xPercent);
    var y = pageHeight * parseFloat(yPercent);
    return [x, y];
}
// 根据分组应用样式
function applyStylesBasedOnGroup(textFrame, group) {
    var paragraphStyleName = "";
    var characterStyleName = "";
    var objectStyleName = "";

    // 根据不同分组设置样式名称
    switch (group) {
        case 1:
            paragraphStyleName = "框内";
            characterStyleName = "字符样式1";
            objectStyleName = "文本框居中-垂直-自动缩框";
            break;
        case 2:
            paragraphStyleName = "框外";
            characterStyleName = "字符样式2";
            objectStyleName = "文本框居中-垂直-自动缩框";
            break;
        // 根据需求添加更多分组和样式
        default:
            paragraphStyleName = "[基本段落]";
            characterStyleName = "[无]";
            objectStyleName = "文本框居中-垂直-自动缩框";
    }

    // 应用段落样式
    try {
        var paragraphStyle = doc.paragraphStyles.item(paragraphStyleName);
        if (paragraphStyle.isValid) {
            textFrame.paragraphs[0].appliedParagraphStyle = paragraphStyle;
        }
    } catch (e) {
        alert("Paragraph style not found: " + paragraphStyleName);
    }

    // 应用字符样式
    try {
        var characterStyle = doc.characterStyles.item(characterStyleName);
        if (characterStyle.isValid) {
            for (var i = 0; i < textFrame.characters.length; i++) {
                textFrame.characters[i].appliedCharacterStyle = characterStyle;
            }
        }
    } catch (e) {
        alert("Character style not found: " + characterStyleName);
    }

    // 应用对象样式
    try {
        var objectStyle = doc.objectStyles.item(objectStyleName);
        if (objectStyle.isValid) {
            textFrame.appliedObjectStyle = objectStyle;
        }
    } catch (e) {
        alert("Object style not found: " + objectStyleName);
    }
}

// 单页操作，插入文本到页面
function insertTextOnPageByTxtEntry(entry) {
    var pageIndex = parseInt(entry.pageImage,10) - 1; // 假设pageImage类似001.tif
    var page = doc.pages[pageIndex];
    
    // 如果页面不存在，则新建页面
    if (!page) {
        page = doc.pages.add();
    }

    // 将坐标百分比转换为绝对坐标
    var coordinates = convertPercentToAbsoluteCoordinates(page, entry.position[0], entry.position[1]);
    var textFrame = page.textFrames.add();
    textFrame_x = 10;//设置文本框大小
    textFrame_y = 25;
    textFrame.geometricBounds = [coordinates[1], coordinates[0]-textFrame_x/2, coordinates[1] + textFrame_y, coordinates[0] + textFrame_x];
    textFrame.contents = entry.text;
    textFrame.parentStory.storyPreferences.storyOrientation = StoryHorizontalOrVertical.VERTICAL;
    if (textFrame.overflows) {
        textFrame.fit(FitOptions.FRAME_TO_CONTENT);
    }


    // 根据分组应用样式
    // applyStylesBasedOnGroup(textFrame, entry.group);
    // 应用对象样式
    applySelectedObjectStyles(textFrame, styleRules)
    
}
// 遍历每一页
function processTxtEntries(txtEntries) {
    // 确保页面数足够
    ensureEnoughPages(totalPages);
    // 遍历每个条目并插入到页面中
    for (var i = 0; i < txtEntries.length; i++) {
        insertTextOnPageByTxtEntry(txtEntries[i], doc);
    }
}
// 创建第二个用户界面匹配模式
function showSecondInterface() {
    var dialog = new Window("dialog", "LP翻译稿导入工具"+version);

    // 顶部：文件选择与单行/多行断句选项（原第一个界面功能合并到这里）
    var topGroup = dialog.add("group");
    topGroup.orientation = "column";
    var modeGroup = topGroup.add("group");
    modeGroup.orientation = "row";
    var singleLineRadio = modeGroup.add("radiobutton", undefined, "单行不断句\n文本一行为一个气泡");
    singleLineRadio.preferredSize = [200, 40];
    var multiLineRadioBtn = modeGroup.add("radiobutton", undefined, "多行断句\n文本多行为一个气泡");
    multiLineRadioBtn.preferredSize = [200, 40];
    singleLineRadio.value = true;

    var fileGroupTop = topGroup.add("group");
    fileGroupTop.orientation = "row";
    fileGroupTop.add("statictext", undefined, "打开lptxt:");
    var filePathInput = fileGroupTop.add("edittext", undefined, "");
    filePathInput.characters = 35;
    var browseButtonTop = fileGroupTop.add("button", undefined, "浏览");
    browseButtonTop.onClick = function () {
        var txtFile = File.openDialog("请选择一个LP翻译稿", "*.txt");
        if (txtFile) {
            filePathInput.text = txtFile.fsName;
            // 选择文件后自动加载并刷新列表
            try { loadAndFillFromPath(filePathInput.text); } catch (e) {}
        }
    };

    // 左侧选项部分
    var leftGroup = dialog.add("group");
    leftGroup.orientation = "column";
    
    var matchOptionGroup = leftGroup.add("panel", undefined, "匹配方式");
    matchOptionGroup.orientation = "column";
    matchOptionGroup.alignChildren = "left";
    var fromStartoEndCheckbox = matchOptionGroup.add("radiobutton", undefined, "匹配页码导入所有页面的文本");
    fromStartoEndCheckbox.value = true; // 默认选中
    var matchByNumberCheckbox = matchOptionGroup.add("radiobutton", undefined, "匹配页码导入选定页面的文本");
    
    var startFromPageCheckbox = matchOptionGroup.add("radiobutton", undefined, "从文档的第X页导入选定页面的文本");;
    var startPageInput = matchOptionGroup.add("edittext", undefined, "1");
    startPageInput.characters = 3;
    startPageInput.enabled = false;
    startPageInput.enabled = startFromPageCheckbox.value;
    startFromPageCheckbox.onClick = function () {
        startPageInput.enabled = startFromPageCheckbox.value;
    };
    // 页码偏移设置
    matchOptionGroup.add("statictext", undefined, "页码偏移(填整数, 可为负):");
    var offsetInput = matchOptionGroup.add("edittext", undefined, "0");
    offsetInput.characters = 4;

    // 右侧文件列表部分
    var rightGroup = dialog.add("group");
    rightGroup.orientation = "column";

    rightGroup.add("statictext", undefined, "LPtxt翻译稿对应页码");
    var pageNames = [];
    var pageNumbers = [];
    // 创建listbox（初始为空，用户选择文件或点击刷新时加载）
    var fileList = rightGroup.add("listbox", undefined, undefined, {
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: ["文件名", "页码"],
        multiselect: true
    });
    fileList.maximumSize.height = 300;
    fileList.preferredSize = [400, 300];
    fileList.alignment = "fill";

    function loadAndFillFromPath(path) {
        pageNames = [];
        pageNumbers = [];
        fileList.removeAll();
        if (!path) return;
        var res = readAndParseTxtFileBase(path, replacements);
        var lines = res.lines || [];
        for (var li = 0; li < lines.length; li++) {
            var pageMatch = lines[li].match(/>>>>>>>>\[(.*)\]<<<<<<<<$/);
            if (pageMatch) {
                pageNames.push(pageMatch[1]);
            }
        }
        pageNumbers = extractPageNumbers(pageNames);
        for (var ii = 0; ii < pageNames.length; ii++) {
            var it = fileList.add("item", pageNames[ii]);
            if (pageNumbers[ii] !== undefined) it.subItems[0].text = pageNumbers[ii];
            it.selected = true;
        }
        dialog.selectedPages = [];
    }

    // 选中的文件列表
    dialog.selectedPages = [];

    // 更新选中的文件列表
    fileList.onChange = function() {
        dialog.selectedPages = [];
        for (var i = 0; i < fileList.items.length; i++) {
            if (fileList.items[i].selected) {
                dialog.selectedPages.push(pageNames[i]);
            }
        }
    };

    // 添加一个刷新按钮以在用户选择文件后刷新列表
    var refreshBtn = rightGroup.add("button", undefined, "刷新列表");
    refreshBtn.onClick = function () {
        loadAndFillFromPath(filePathInput.text);
    };

    // 按钮组
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var cancelButton = buttonGroup.add("button", undefined, "取消");
    var confirmButton = buttonGroup.add("button", undefined, "确定");

    // 取消按钮点击事件
    cancelButton.onClick = function () {
        dialog.close(0);
    };

    // 确定按钮点击事件
    confirmButton.onClick = function () {
        if (fromStartoEndCheckbox.value) {
            fromStartoEnd = true;
            matchByNumber = false;
            startFromPage = false;
        } else if(matchByNumberCheckbox.value){
            fromStartoEnd = false;
            matchByNumber = true;
            startFromPage = false;
        } else if(startFromPageCheckbox.value){
            fromStartoEnd = false;
            matchByNumber = false;
            startFromPage = true;
        }

        // 如果未选择文件或列表为空，提示并返回
        if (!filePathInput.text || fileList.items.length === 0) {
            alert('请先选择一个 LP 文本文件。');
            return;
        }

        // 获取用户选择的页码列表,纯数字（先重置）
        selectedPages = [];
        for (var i = 0; i < fileList.items.length; i++) {
            if (fileList.items[i].selected) {
                selectedPages.push(pageNumbers[i]);
            }
        }

        // 读取并解析页码偏移值
        pageOffset = parseInt(offsetInput.text, 10);
        if (isNaN(pageOffset)) {
            pageOffset = 0;
        }
        if (pageOffset !== 0) {
            for (var si = 0; si < selectedPages.length; si++) {
                selectedPages[si] = selectedPages[si] + pageOffset;
            }
        }

        // 获取用户输入的起始页码
        startPage = parseInt(startPageInput.text, 10);
        if (isNaN(startPage) || startPage < 1) {
            alert("请输入有效的起始页码！");
            return;
        }

        dialog.close(1);
    };
    var result = dialog.show(); 
    if (result === 0) {
        return null;
    }
    if (result === 1) {
        showThirdInterface(filePathInput.text);
    }
}

// 创建第三个用户界面替换选项
function showThirdInterface(filePathInput) {
    var dialog = new Window("dialog", "文本替换与样式匹配");

    // 主分组
    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = "fill";

    // 左侧：文字替换
    var leftPanel = mainGroup.add("panel", undefined, "文字替换");
    leftPanel.alignChildren = "fill";
    leftPanel.preferredSize = [340, 350];
    var leftAddBtn = leftPanel.add("button", undefined, "+");
    var leftListPanel = leftPanel.add("panel");
    leftListPanel.preferredSize = [320, 300];
    leftListPanel.alignChildren = "top";
    leftListPanel.orientation = "column";
    leftListPanel.margins = [0,0,0,0];
    leftListPanel.spacing = 2;
    leftListPanel.maximumSize.height = 800;

    // 右侧：对象样式匹配
    var rightPanel = mainGroup.add("panel", undefined, "对象样式匹配");
    rightPanel.alignChildren = "fill";
    rightPanel.preferredSize = [340, 350];
    var rightAddBtn = rightPanel.add("button", undefined, "+");
    var rightListPanel = rightPanel.add("panel");
    rightListPanel.preferredSize = [320, 300];
    rightListPanel.alignChildren = "top";
    rightListPanel.orientation = "column";
    rightListPanel.margins = [0,0,0,0];
    rightListPanel.spacing = 2;
    rightListPanel.maximumSize.height = 800;

    // 行数据存储
    var leftRows = [];
    var rightRows = [];

    // 左侧添加一行
    function addLeftRow(fromText, toText, checked) {
        var g = leftListPanel.add("group");
        g.orientation = "row";
        g.alignChildren = "center";
        g.preferredSize = [310, 24];
        var cb = g.add("checkbox", undefined, "");
        cb.value = (typeof checked === "boolean") ? checked : true;
        var fromEdit = g.add("edittext", undefined, fromText || "");
        fromEdit.preferredSize = [90, 22];
        var label = g.add("statictext", undefined, "→");
        var toEdit = g.add("edittext", undefined, toText || "");
        toEdit.preferredSize = [90, 22];
        leftRows.push({cb: cb, fromEdit: fromEdit, toEdit: toEdit, group: g});
        leftListPanel.layout.layout(true);
    }

    // 右侧添加一行
    function addRightRow(matchText, styleName, checked) {
        var g = rightListPanel.add("group");
        g.orientation = "row";
        g.alignChildren = "center";
        g.preferredSize = [310, 24];
        var cb = g.add("checkbox", undefined, "");
        cb.value = (typeof checked === "boolean") ? checked : true;
        if (matchText!=='默认匹配') {
            var matchEdit = g.add("edittext", undefined, matchText || "");
            matchEdit.preferredSize = [90, 22];
        }else {
            var matchEdit = g.add("statictext", undefined, "默认匹配");
        }
        
        var label = g.add("statictext", undefined, "应用样式");
        var docObjectstyles = getDocumentObjectStyles();
        var styleNames = [];
        for (var j = 0; j < docObjectstyles.length; j++) {
            styleNames.push(docObjectstyles[j].name);
        }
        var dropdown = g.add("dropdownlist", undefined, styleNames);
        if (matchText=="*"||matchText=="※"||matchText=="＊") {
            if(dropdown.children.length > 5){
                dropdown.selection = 4;
            }else {
                dropdown.selection = 3;
            }
        }else if(dropdown.children.length > 5){
            dropdown.selection = 5;
            }else {
                dropdown.selection = 3;
            }
        dropdown.preferredSize = [120, 22];
        rightRows.push({cb: cb, matchEdit: matchEdit, dropdown: dropdown, group: g});
        rightListPanel.layout.layout(true);
    }

    // 填充左侧
    for (var key in replacements) {
        if (replacements.hasOwnProperty(key)) {
            var checked = !(key === "" || key === "？");
            addLeftRow(key, replacements[key], checked);
        }
    }
    // 填充右侧
    for (var i = 0; i < objectStyleMatchText.length; i++) {
        var opt = objectStyleMatchText[i];
        addRightRow(opt.label, opt.style, opt.checked);
    }

    // 加号按钮事件
    leftAddBtn.onClick = function () {
        addLeftRow("", "", true);
        leftListPanel.layout.layout(true);
        dialog.layout.layout(true);
        dialog.center();
    };
    rightAddBtn.onClick = function () {
        addRightRow("", "", true);
        rightListPanel.layout.layout(true);
        dialog.layout.layout(true);
        dialog.center();
    };

    // 底部按钮组
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var cancelButton = buttonGroup.add("button", undefined, "取消");
    var confirmButton = buttonGroup.add("button", undefined, "确定");

    // 取消按钮点击事件
    cancelButton.onClick = function () {
        dialog.close(0);
    };

    // 确定按钮点击事件
    confirmButton.onClick = function () {
        // 获取左侧文字替换内容
        replacements = {};
        for (var i = 0; i < leftRows.length; i++) {
            var row = leftRows[i];
            if (row.cb.value && row.fromEdit.text) {
                replacements[row.fromEdit.text] = row.toEdit.text;
            }
        }
        // 获取右侧样式匹配内容
        styleRules = [];
        for (var j = 0; j < rightRows.length; j++) {
            var row = rightRows[j];
            if (row.cb.value && row.matchEdit.text && row.dropdown.selection) {
                styleRules.push({
                    match: row.matchEdit.text,
                    style: row.dropdown.selection.text
                });
            }
        }
        dialog.close(1);
    };
    var result = dialog.show(); 
    if (result === 0) {
        return null;
    }
    if (result === 1) {
        //如果有名字为Text的图层，则选中
        try {
            var textLayer = doc.layers.itemByName("Text");
            if (textLayer.isValid) {
                doc.activeLayer = textLayer;
                textLayer.locked = false;
                textLayer.visible = true;
            }
        } catch (e) {}
        KTUDoScriptAsUndoable(function() {processStart(filePathInput)}, "LP翻译稿导入");
    }
}

// 根据放置选项开始放入台词
function processStart(filePathInput) {
    //获取台词文本，获取的同时已经替换掉替换列表中匹配的文本
    var txtEntries = processFile(filePathInput);
        if (fromStartoEnd){
            processTxtEntries(txtEntries);
        }else if (matchByNumber) {
            //选择列表里的最大值
            var maxOFselectedPageNumber = selectedPages[selectedPages.length - 1];
            totalPages = parseInt(maxOFselectedPageNumber,10);
            //获取选定的台词文本
            var filteredEntries = txtEntriesSelected(selectedPages, txtEntries);
            processTxtEntries(filteredEntries);
        }else if (startPage) {
            totalPages = startPage + selectedPages.length;
            var filteredEntries = txtEntriesSelected(selectedPages, txtEntries);
            filteredEntries = assignPageNumbers(filteredEntries, startPage);
            processTxtEntries(filteredEntries);
        }

    }

// 页码分配函数
function assignPageNumbers(filteredEntries, startPage) {
    // 创建页码分组映射表
    var pageGroups = {};
    for (var i = 0; i < filteredEntries.length; i++) {
        var entry = filteredEntries[i];
        var originalPage = entry.pageImage;
        if (!pageGroups[originalPage]) {
            pageGroups[originalPage] = [];
        }
        pageGroups[originalPage].push(entry);
    }

    // 按分组更新页码
    var groupIndex = 0;
    for (var pageKey in pageGroups) {
        if (pageGroups.hasOwnProperty(pageKey)) {
            var group = pageGroups[pageKey];
            var newPage = startPage + groupIndex;
            
            // 更新整组页码
            for (var j = 0; j < group.length; j++) {
                group[j].pageImage = newPage;
            }
            
            groupIndex++;
        }
    }
    return filteredEntries;
}

// 从文章中筛选选中的页码
function txtEntriesSelected(selectedPages, txtEntries) {
    // 过滤匹配的条目
    var filteredEntries = [];
    for (var j = 0; j < txtEntries.length; j++) {
        var entry = txtEntries[j];
        // 检查当前条目的页码是否在选中页码中
        var isValid = false;
        for (var k = 0; k < selectedPages.length; k++) {
            if (entry.pageImage === selectedPages[k]) {
                isValid = true;
                break;
            }
        }
        if (isValid) {
            filteredEntries.push(entry);
        }
    }
    
    return filteredEntries;
}

// 应用对象样式（根据文本内容匹配规则）
function applySelectedObjectStyles(textFrame, styleRules) {
    var defaultStyleName = $.global['styleRules']['0']['style']; // 默认样式名称
    var applied = false; // 标记是否已应用样式
    
    try {
        // 遍历所有样式规则
        for (var i = 1; i < styleRules.length; i++) {
            var rule = styleRules[i];
            
            // 检查文本内容是否包含匹配字符
            if (textFrame.contents.indexOf(rule.match) !== -1) {
                var style = doc.objectStyles.item(rule.style);
                if (style.isValid) {
                    textFrame.appliedObjectStyle = style;
                    applied = true;
                    break; // 匹配到第一条规则后立即退出循环
                }
            }
        }
        
        // 未匹配任何规则时应用默认样式
        if (!applied) {
            var defaultStyle = doc.objectStyles.item(defaultStyleName);
            if (defaultStyle.isValid) {
                textFrame.appliedObjectStyle = defaultStyle;
            }
        }
    } catch (e) {
        alert("样式应用错误: " + e);
    }
}

// 获取当前文档的对象样式
function getDocumentObjectStyles() {
    var styles = [];
    for (var i = 0; i < doc.objectStyles.length; i++) {
        styles.push({ name: doc.objectStyles[i].name });
    }
    return styles;
}


// 处理文件+改单位原点
function processFile(filePath) {
    if (singleLineMode) {
        var txtEntries = readAndParseTxtFileoneline(filePath);
    }else if (multiLineRadio) {
        var txtEntries = readAndParseTxtFilelines(filePath);
    }

    // 设置标尺原点和单位
    doc.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;
    doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
    doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
    doc.zeroPoint = [0, 0];

return txtEntries;
}

// 直接从第二个界面开始运行
showSecondInterface();
//调试用
// var filePathInput = "M:\\汉化\\PS_PNG\\output_数据.txt"
// // 调用第二界面
// var SecondInterface = showSecondInterface(filePathInput);