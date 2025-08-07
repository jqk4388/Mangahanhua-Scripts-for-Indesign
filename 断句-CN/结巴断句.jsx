// 兼容性垫片（ES3环境）
if (!Array.prototype.indexOf) {
    Array.prototype.indexOf = function(item) {
        for (var i = 0; i < this.length; i++) {
            if (this[i] === item) return i;
        }
        return -1;
    };
}
// 主入口函数
function main() {
    try {
        // 显示UI界面
        var userOptions = showUI();
        if (!userOptions) {
            alert("操作已取消。");
            return;
        }

        // 根据用户选择的范围执行断句
        var textFrames = getTextFrames(userOptions.range);

        if (textFrames.length === 0) {
            alert("未找到任何文本框。");
            return;
        }

        DoalltextFrames(textFrames);//执行所有文本框

    
        // alert("断句完成！");
    } catch (e) {
        alert("发生错误: " + e.message);
    }
}
function DoalltextFrames(textFrames) {
    // 创建临时文件路径，兼容macOS和Windows
    var tempFolder = Folder.temp;
    var inputFile = new File(tempFolder.fsName + "/jieba_temp_input.txt");
    var outputFile = new File(tempFolder.fsName + "/jieba_temp_output.txt");

    // 打开文件准备写入
    inputFile.encoding = "UTF-8";
    inputFile.open("w");
    
    // 遍历所有文本框，写入内容
    for (var i = 0; i < textFrames.length; i++) {
        var textFrame = textFrames[i];
        var content = textFrame['parentStory']['contents'];
        content = content.replace(/[\r\n]/g, ""); // 移除换行符
        
        // 写入当前文本框内容并添加换行
        inputFile.write(content + "\n");
    }
    
    // 关闭文件
    inputFile.close();
    // 调用 jieba-pytojs 处理文本
    var jsxPath = File($.fileName).parent
    var pythonScript = new File(jsxPath.fsName + "/jieba_pytojs.pyw");
    pythonScript.execute("\"" + inputFile.fsName + "\" \"" + outputFile.fsName + "\"");

    // 等待用户确认脚本是否完成
    var w = new Window("dialog", "提示");
    w.orientation = "column";
    var msg = w.add("statictext", undefined, "!!!!请等待弹窗!!!!\n未弹窗请选择No。\n\n弹窗后点击“Yes”替换断句结果，点击“No”退出。", {multiline: true});
    msg.characters = 40;
    msg.graphics.font = ScriptUI.newFont(msg.graphics.font.name, msg.graphics.font.style, 24); // 放大字号
    var btnGroup = w.add("group");
    btnGroup.alignment = "right";
    var yesBtn = btnGroup.add("button", undefined, "Yes", {name: "ok"});
    var noBtn = btnGroup.add("button", undefined, "No", {name: "cancel"});
    var isDone = false;
    yesBtn.onClick = function() { isDone = true; w.close(); };
    noBtn.onClick = function() { isDone = false; w.close(); };
    w.show();
    if (!isDone) {
        // 删除临时文件
        inputFile.remove();
        outputFile.remove();
        throw new Error("用户取消，脚本终止。");
    }
    
    // 读取处理后的文本
    outputFile.encoding = "UTF-8";
    outputFile.open("r");
    var segmentedContents = outputFile.read();
    outputFile.close();
    // 将文本按行分割
    var contentLines = segmentedContents.split("\n");
    
    // 遍历每个文本框，更新内容
    for (var i = 0; i < textFrames.length && i < contentLines.length; i++) {

        // 将\r转换为真正的换行符
        var processedContent = contentLines[i].replace(/\\r/g, "\n");
        
        // 更新文本框内容
        textFrames[i]['parentStory']['contents'] = processedContent || '';
    }

    // // 删除临时文件
    // inputFile.remove();
    // outputFile.remove();
}
  
// 显示UI界面
function showUI() {
    var dialog = new Window("dialog", "自动断句");

    // 断句范围选项
    dialog.add("statictext", undefined, "断句范围：");
    var rangeGroup = dialog.add("group");
    var currentSelectionRadio = rangeGroup.add("radiobutton", undefined, "当前选中的文本框");
    var currentPageRadio = rangeGroup.add("radiobutton", undefined, "当页所有文本框");
    var entireDocumentRadio = rangeGroup.add("radiobutton", undefined, "文档中所有文本框");
    currentSelectionRadio.value = true; // 默认选中第一个选项

    // 确定和取消按钮
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    var okButton = buttonGroup.add("button", undefined, "确定", { name: "ok" });
    var cancelButton = buttonGroup.add("button", undefined, "取消", { name: "cancel" });

    // 处理用户交互
    if (dialog.show() === 1) {
        var range;
        if (currentSelectionRadio.value) {
            range = "currentSelection";
        } else if (currentPageRadio.value) {
            range = "currentPage";
        } else {
            range = "entireDocument";
        }

        return {
            range: range,

        };
    } else {
        return null;
    }
}

// 获取指定范围的文本框
function getTextFrames(range) {
    var textFrames = [];
    var doc = app.activeDocument;

    if (range === "currentSelection") {
        for (var i = 0; i < app.selection.length; i++) {
            if (app.selection[i].constructor.name === "TextFrame") {
                textFrames.push(app.selection[i]);
            }
        }
    } else if (range === "currentPage") {
        var currentPage = app.activeWindow.activePage;
        textFrames = currentPage.textFrames.everyItem().getElements();
    } else if (range === "entireDocument") {
        var allTextFrames = doc.textFrames.everyItem().getElements();
        var normalPageTextFrames = [];

        for (var i = 0; i < allTextFrames.length; i++) {
            var tf = allTextFrames[i];
            try {
                // 如果 textFrame 所在的 parentPage 存在，并且该页面不是主页
                if (tf.parentPage != null && tf.parentPage.parent.constructor.name != "MasterSpread") {
                    normalPageTextFrames.push(tf);
                }
            } catch (e) {
                alert("发生错误: " + e.message);
            }
        }
        textFrames = normalPageTextFrames;
    }

    return textFrames;
}

// 执行脚本
main();