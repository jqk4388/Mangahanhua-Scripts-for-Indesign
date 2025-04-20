#include "mergeJiebaLines.js"
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

        var waitTime = userOptions.waitTime; // 获取用户设置的等待时间

        // 根据用户选择的范围执行断句
        var textFrames = getTextFrames(userOptions.range);
        if (textFrames.length === 0) {
            alert("未找到任何文本框。");
            return;
        }

        // 创建进度条对话框
        var progressWin = new Window("palette", "处理进度");
        progressWin.progressBar = progressWin.add("progressbar", undefined, 0, textFrames.length);
        progressWin.progressBar.preferredSize.width = 300;
        progressWin.status = progressWin.add("statictext", undefined, "正在处理...");
        progressWin.center();
        progressWin.show();

        for (var i = 0; i < textFrames.length; i++) {
            // 更新进度条
            progressWin.progressBar.value = i + 1;
            progressWin.status.text = "正在处理第 " + (i + 1) + " 个文本框，共 " + textFrames.length + " 个";

            var textFrame = textFrames[i];
            var content = textFrame['parentStory']['contents'];
            content = content.replace(/[\r\n]/g, "");
            // 创建临时文件路径
            var tempFolder = Folder.temp;
            var inputFile = new File(tempFolder + "/jieba_temp_input.txt");
            var outputFile = new File(tempFolder + "/jieba_temp_output.txt");

            // 写入内容到临时文件
            inputFile.encoding = "UTF-8";
            inputFile.open("w");
            inputFile.write(content);
            inputFile.close();

            // 调用 jieba-pytojs 处理文本
            var jsxPath = File($.fileName).parent
            var pythonScript = new File(jsxPath.fsName + "/jieba_pytojs.pyw");
            pythonScript.execute("\"" + inputFile.fsName + "\" \"" + outputFile.fsName + "\"");
            
            if (i == 0) {
                $.sleep(waitTime * 1000 + 2000); // 第一个框建立词典缓存等待时间长
            } else {
                $.sleep(waitTime * 1000);
            }

            // 读取处理后的文本
            outputFile.encoding = "UTF-8";
            outputFile.open("r");
            var segmentedContent = outputFile.read();
            outputFile.close();

            //处理后的文本重新合并短句
            segmentedContent = mergeJiebaLines(segmentedContent);

            // 更新文本框内容
            textFrame['parentStory']['contents'] = segmentedContent;

            // 删除临时文件
            inputFile.remove();
            outputFile.remove();
        }

        // 关闭进度条
        progressWin.close();

        // alert("断句完成！");
    } catch (e) {
        alert("发生错误: " + e.message);
    }
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

    // 等待时间滑块
    dialog.add("statictext", undefined, "等待时间：性能不好的电脑建议调整成2秒以上");
    var waitTimeGroup = dialog.add("group");
    var waitTimeSlider = waitTimeGroup.add("slider", undefined, 1, 1, 6); // 默认1秒，范围1-8秒
    waitTimeSlider.preferredSize.width = 200;
    var waitTimeText = waitTimeGroup.add("statictext", undefined, "1 秒");
    waitTimeSlider.onChanging = function () {
        waitTimeText.text = Math.round(waitTimeSlider.value) + " 秒";
    };

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
            waitTime: Math.round(waitTimeSlider.value), // 返回用户设置的等待时间
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
        textFrames = doc.textFrames.everyItem().getElements();
    }

    return textFrames;
}

// 执行脚本
main();