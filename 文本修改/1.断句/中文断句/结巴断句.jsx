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
            // 等待1秒
            $.sleep(800);

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
        textFrames = doc.textFrames.everyItem().getElements();
    }

    return textFrames;
}

// 执行脚本
main();

// function mergeJiebaLines(text) {
//     // 停顿标点和句末助词
//     var PAUSE_PUNCT = ["……", "。", "！", "？", "?", "，", "；", "、", "—", "?!", "!!", "！！", "!!!", "!?", "？！"];
//     var END_PARTICLE = ["的", "了", "吗", "吧", "呢"];
//     // 绝对不能分开的词（可外挂词典）
//     var ABSOLUTE_UNBREAKABLE = ["……", "？！", "!!", "！！", "!!!", "?!", "!?", "——"];
//     //不可分割的词
//     var UNBREAKABLE_WORDS = ["会不会", "能不能", "要不要", "好不好", "对不对", "是不是", "有没有", "行不行", "该不该", "肯不肯", "愿不愿", "敢不敢", "愧不愧", "怪不得", "看样子", "要不然", "算起来", "说不定", "无所谓", "差不多", "没关系", "不管怎样", "不管怎么说", "不管怎么样", "不管怎么说", "不管怎样", "不管怎么说","还需要", "还应该", "还必须", "还能够", "还可以","拔刀斋","和平年代"];
//     // 若有外挂词典文件，可在此处加载并合并到ABSOLUTE_UNBREAKABLE
//         for (var u = 0; u < UNBREAKABLE_WORDS.length; u++) {
//             if (ABSOLUTE_UNBREAKABLE.indexOf(UNBREAKABLE_WORDS[u]) < 0) {
//                 ABSOLUTE_UNBREAKABLE.push(UNBREAKABLE_WORDS[u]);
//             }
//         }
//     // 1. 词组数组
//     var words = text.replace(/\r/g, '').split('\n');
//     // 2. 合并停顿标点到前一个词
//     var merged = [];
//     for (var i = 0; i < words.length; i++) {
//         var w = words[i];
//         // 合并标点到前一个词
//         if (PAUSE_PUNCT.indexOf(w) >= 0 && merged.length > 0) {
//             merged[merged.length - 1] += w;
//         } else if (w === "" && merged.length > 0) {
//             // 跳过空词
//             continue;
//         } else {
//             merged.push(w);
//         }
//     }
//     // 3. 合并句末助词
//     var merged2 = [];
//     for (var j = 0; j < merged.length; j++) {
//         var w2 = merged[j];
//         if (END_PARTICLE.indexOf(w2) >= 0 && merged2.length > 0) {
//             merged2[merged2.length - 1] += w2;
//         } else {
//             merged2.push(w2);
//         }
//     }
//     // 4. 统计总字数
//     var totalLen = 0;
//     for (var k = 0; k < merged2.length; k++) {
//         totalLen += merged2[k].length;
//     }
//     // 5. 计算行数和理想行长
//     var minLines = 1;
//     var maxLineLen = 6; // 漫画气泡常用最大行长
//     var lines = Math.max(minLines, Math.ceil(totalLen / maxLineLen));
//     var idealLen = Math.round(totalLen / lines);
//     // 6. 分行
//     var result = [];
//     var line = "";
//     var lineLen = 0;
//     for (var m = 0; m < merged2.length; m++) {
//         var word = merged2[m];
//         // 绝对不能分开的词，单独成段或与前后词合并时整体移动
//         if (ABSOLUTE_UNBREAKABLE.indexOf(word) >= 0) {
//             if (lineLen > 0) {
//                 result.push(line);
//                 line = "";
//                 lineLen = 0;
//             }
//             result.push(word);
//             continue;
//         }
//         // 如果加上当前词超出理想长度1字，且当前行不空，则换行
//         if (lineLen > 0 && (lineLen + word.length > idealLen + 1)) {
//             result.push(line);
//             line = "";
//             lineLen = 0;
//         }
//         // 长词组不能拆分
//         if (lineLen > 0 && word.length >= idealLen + 2) {
//             result.push(line);
//             result.push(word);
//             line = "";
//             lineLen = 0;
//             continue;
//         }
//         // 合并到当前行
//         line += word;
//         lineLen += word.length;
//         // 如果刚好到理想长度附近，且不是最后一个词，则换行
//         if (lineLen >= idealLen && m < merged2.length - 1) {
//             result.push(line);
//             line = "";
//             lineLen = 0;
//         }
//     }
//     if (lineLen > 0) {
//         result.push(line);
//     }
//     // 7. 处理标点后多余换行
//     for (var n = 0; n < result.length - 1; n++) {
//         var lastChar = result[n].charAt(result[n].length - 1);
//         // 如果以标点结尾，下一行不是标点，则用\r，否则用\n
//         if (PAUSE_PUNCT.join('').indexOf(lastChar) >= 0) {
//             result[n] += "\r";
//         } else {
//             result[n] += "\n";
//         }
//     }
//     // 8. 合并结果
//     return result.join('').replace(/[\r\n]+$/,"");
// }