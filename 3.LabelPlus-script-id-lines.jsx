// 声明全局变量
var totalPages = 0;

function readAndParseTxtFile(filePath) {
    var file = new File(filePath);
    if (!file.exists) {
        alert("File not found: " + filePath);
        return [];
    }
    file.open("r");
    var content = file.read();
    var replacements = {
        "！": "!",
        "？": "?",
        "——": "—"
    };
    
    var content = replaceMultipleWords(content, replacements);
    file.close();

    var lines = content.split("\n");
    var entries = [];

    totalPages = 0; // 重置页面总数

    for (var i = 0; i < lines.length; i++) {
        // 匹配页面名
        var pageMatch = lines[i].match(/>>>>>>>>\[(.*)\]<<<<<<<<$/);
        if (pageMatch) {
            totalPages++; // 每遇到一个新页面时，页面总数增加
            continue;
        }

        // 匹配台词编号、位置和分组信息
        var textMatch = lines[i].match(/----------------\[(\d+)\]----------------\[(.*?)\,(.*?)\,(.*?)\]/);
        if (textMatch) {
            var entry = {
                pageImage: totalPages, // 当前条目所处的页码
                pageNumber: parseInt(textMatch[1]), // 当前台词在页面上的编号（例如：1）
                position: [parseFloat(textMatch[2]), parseFloat(textMatch[3])], // 坐标百分比
                group: parseInt(textMatch[4]) // 分组
            };

            // 收集文本内容，直到遇到下一个编号或下一页
            var textContent = [];
            for (var j = i + 1; j < lines.length; j++) {
                if (lines[j].match(/----------------\[\d+\]----------------/) || lines[j].match(/>>>>>>>>\[(.*)\]<<<<<<<<$/)) {
                    break; // 遇到下一个编号或下一页，停止收集
                }
                textContent.push(String(lines[j])); // 确保加入的是字符串
            }
            entry.text = textContent.join("\n"); // 将多行文本合并为一个字符串，保留换行符
            entries.push(entry);
        }
    }

    // alert("Total pages: " + totalPages); // 可以用 alert 或其他方式显示页面总数
    return entries;
}

function replaceMultipleWords(text, replacements) {
    // 遍历替换字典中的每个键值对，并替换对应的词
    for (var key in replacements) {
        if (replacements.hasOwnProperty(key)) {
            var regex = new RegExp(key, 'g'); // 使用正则表达式全局替换
            text = text.replace(regex, replacements[key]);
        }
    }
    return text;
}


function ensureEnoughPages(doc, requiredPageCount) {
    var currentPageCount = doc.pages.length;
    
    if (currentPageCount < requiredPageCount) {
        var pagesToAdd = requiredPageCount - currentPageCount;
        for (var i = 0; i < pagesToAdd; i++) {
            doc.pages.add();
        }
    }
}

function convertPercentToAbsoluteCoordinates(page, xPercent, yPercent) {
    var pageWidth = page.bounds[3] - page.bounds[1]; // 页面宽度
    var pageHeight = page.bounds[2] - page.bounds[0]; // 页面高度
    var x = pageWidth * parseFloat(xPercent);
    var y = pageHeight * parseFloat(yPercent);
    return [x, y];
}
function applyStylesBasedOnGroup(doc, textFrame, group) {
    var paragraphStyleName = "";
    var characterStyleName = "";
    var objectStyleName = "";

    // 根据不同分组设置样式名称
    switch (group) {
        case 1:
            paragraphStyleName = "框内";
            characterStyleName = "字符样式1";
            objectStyleName = "文本框居中-垂直";
            break;
        case 2:
            paragraphStyleName = "框外";
            characterStyleName = "字符样式2";
            objectStyleName = "对象样式 2";
            break;
        // 根据需求添加更多分组和样式
        default:
            paragraphStyleName = "[基本段落]";
            characterStyleName = "[无]";
            objectStyleName = "[基本文本框架]";
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


function insertTextOnPageByTxtEntry(entry, doc) {
    var pageIndex = parseInt(entry.pageImage,10) - 1; // 假设pageImage类似001.tif
    var page = doc.pages[pageIndex];
    
    // 如果页面不存在，则新建页面
    if (!page) {
        page = doc.pages.add();
    }

    // 将坐标百分比转换为绝对坐标
    var coordinates = convertPercentToAbsoluteCoordinates(page, entry.position[0], entry.position[1]);
    var textFrame = page.textFrames.add();
    textFrame_x = 18;//设置文本框大小
    textFrame_y = 23;
    textFrame.geometricBounds = [coordinates[1], coordinates[0]-textFrame_x/2, coordinates[1] + textFrame_y, coordinates[0] + textFrame_x];
    textFrame.contents = entry.text;
    textFrame.parentStory.storyPreferences.storyOrientation = StoryHorizontalOrVertical["VERTICAL"];


    // 根据分组应用样式

    applyStylesBasedOnGroup(doc, textFrame, entry.group);
    
}
// 处理台词
function processTxtEntries(txtEntries, doc) {
    // 确保页面数足够
    ensureEnoughPages(doc, totalPages);

    // 遍历每个条目并插入到页面中
    for (var i = 0; i < txtEntries.length; i++) {
        insertTextOnPageByTxtEntry(txtEntries[i], doc);
    }
}

// 用户选择文件
function selectAndReadFiletxt() {
    var txtFile = File.openDialog("请选择一个LP翻译稿", "*.txt");
    if (txtFile) {
        var txtEntries = readAndParseTxtFile(txtFile.fsName);
        // 所有台词装入txtEntries
        var doc = app.activeDocument;
        //改标尺的原点和单位，设置默认零点
        doc.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;
        doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters
        doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters
        doc.viewPreferences.typographicMeasurementUnits = MeasurementUnits.points
        doc.viewPreferences.textSizeMeasurementUnits = MeasurementUnits.points
        doc.viewPreferences.printDialogMeasurementUnits = MeasurementUnits.points
        doc.zeroPoint = [0, 0];
        // 确保页面数足够
        ensureEnoughPages(doc, totalPages);
        processTxtEntries(txtEntries, doc);
    } else {
        alert("未选择任何文件。");
    }
    return txtEntries;
}

// 调用选择文件的函数
selectAndReadFiletxt();


