// Adobe InDesign Script: 自动拼音脚本
// 主要功能：为选中文本框中的括号内容添加拼音注音
// 作者：几千块
// 日期：2025年3月30日
// 版本：1.0

// 主函数
main();

function main() {
    // 确保有文档打开
    if (app.documents.length === 0) {
        alert("请先打开一个文档！");
        return;
    }

    // 获取当前选中对象
    var selection = app.selection;
    if (selection.length === 0 || !(selection[0].constructor.name === "TextFrame")) {
        alert("请先选择文本框！");
        return;
    }

    try {
        // 处理每个选中的文本框
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].constructor.name === "TextFrame") {
                processTextFrame(selection[i]);
            }
        }
    } catch (e) {
        alert("发生错误：" + e);
    }
}

function processTextFrame(textFrame) {
    textFrame.fit(FitOptions.FRAME_TO_CONTENT);
    var content = textFrame.contents;
    var lines = content.split("\r");
    var globalBracketsPositions = []; // 用于记录括号在整个文本框中的全局位置
    var localBracketsPositions = [];  // 用于记录括号在每行中的相对位置

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var brackets = findBrackets(line);

        if (brackets.length > 0) {
            for (var j = 0; j < brackets.length; j++) {
                var bracketInfo = brackets[j];
                var targetText = line.substring(0, bracketInfo.start);
                var rubyText = line.substring(bracketInfo.start + 1, bracketInfo.end);

                // 获取默认字符数（当前行括号前的字符数）
                var defaultChars = targetText.length;

                // 显示对话框
                var prevLine = i > 0 ? lines[i - 1] : "";
                var userInput = showRubyDialog(targetText, rubyText, prevLine, defaultChars);

                if (userInput === null) {
                    return; // 用户取消操作
                }

                if (userInput > 0) {
                    try {
                        // 应用拼音，基于当前行的字符位置
                        applyRubyToLine(textFrame, i, bracketInfo.start - userInput, userInput, rubyText);
                    } catch (e) {
                        alert("应用拼音失败：" + e);
                    }

                    // 如果有前一行，调整行距
                    if (i > 0) {
                        adjustPreviousLineLeading(textFrame, i);
                    }

                    // 记录括号的全局位置
                    var currentPosition = 0;
                    // 计算当前行之前的所有字符数
                    for (var k = 0; k < i; k++) {
                        currentPosition += lines[k].length + 1; // +1 是为了包含换行符
                    }
                    globalBracketsPositions.push({
                        start: currentPosition + bracketInfo.start,
                        end: currentPosition + bracketInfo.end
                    });

                    // 记录括号的相对位置（行内位置）
                    localBracketsPositions.push({
                        lineIndex: i,
                        start: bracketInfo.start,
                        end: bracketInfo.end
                    });
            }
        }
    }
    }

    // 删除所有括号及其内容
    removeBracketsWithFindReplace(textFrame, globalBracketsPositions);
    textFrame.fit(FitOptions.FRAME_TO_CONTENT);
}

function findBrackets(text) {
    var results = [];
    var bracketPairs = [
        {open: '(', close: ')'}, 
        {open: '（', close: '）'}
    ];
    
    for (var i = 0; i < text.length; i++) {
        for (var j = 0; j < bracketPairs.length; j++) {
            if (text[i] === bracketPairs[j].open) {
                var closePos = text.indexOf(bracketPairs[j].close, i);
                if (closePos !== -1) {
                    results.push({
                        start: i,
                        end: closePos
                    });
                    i = closePos;
                    break;
                }
            }
        }
    }
    
    return results;
}

function showRubyDialog(targetText, rubyText, prevLine, defaultChars) {
    var dialog = app.dialogs.add({name: "拼音设置"});
    
    try {
        var column = dialog.dialogColumns.add({alignChildren: 'left'});
        
        // // 添加一个静态文本区域用于显示前一行内容
        // if (prevLine !== "") {
        //     var prevLineRow = column.dialogRows.add();
        //     prevLineRow.staticTexts.add({staticLabel: "前一行内容：" + prevLine});
        // }
        
        // 添加一个静态文本区域用于显示目标文本
        var targetTextRow = column.dialogRows.add();
        targetTextRow.staticTexts.add({staticLabel: "目标文本：" + targetText});
        
        // 添加一个静态文本区域用于显示拼音文本
        var rubyTextRow = column.dialogRows.add();
        rubyTextRow.staticTexts.add({staticLabel: "拼音文本：" + rubyText});
        
        // 添加输入框，放在最下方
        var inputRow = column.dialogRows.add({alignChildren: 'left'});
        inputRow.staticTexts.add({staticLabel: "注音字符数（0表示不注音）："});
        var charField = inputRow.integerEditboxes.add({
            editValue: defaultChars,
            minWidth: 100
        });
        
        if (dialog.show()) {
            var value = charField.editValue;
            if (value > targetText.length) {
                alert("输入的字符数超过了目标文本长度！");
                return 0;
            }
            return value;
        } else {
            return null;
        }
    } catch (e) {
        alert("对话框错误：" + e);
    } finally {
        dialog.destroy();
    }
}

function applyRubyToLine(textFrame, lineIndex, startPosition, charCount, rubyText) {
    try {
        // 获取目标行
        var targetLine = textFrame.lines[lineIndex];
        if (targetLine) {
            // 获取目标字符范围
            var targetRange = targetLine.characters.itemByRange(startPosition, startPosition + charCount - 1);
            if (targetRange.isValid) {
                targetRange.rubyFlag = true; // 启用拼音
                targetRange.rubyString = rubyText; // 设置拼音文本
                targetRange.rubyAlignment = RubyAlignments.RUBY_JIS; // 设置拼音对齐方式
            } else {
                alert("目标字符范围无效，无法应用拼音。");
            }
        } else {
            alert("目标行无效，无法应用拼音。");
        }
    } catch (e) {
        alert("应用拼音时发生错误：" + e);
    }
}

function adjustPreviousLineLeading(textFrame, position) {
    try {
        var prevLine = textFrame.lines[position - 1];
        if (prevLine) {
            var maxFontSize = 0;
            // 获取行中最大字号
            for (var i = 0; i < prevLine.characters.length; i++) {
                var fontSize = prevLine.characters[i].pointSize;
                if (fontSize > maxFontSize) {
                    maxFontSize = fontSize;
                }
            }
            prevLine.leading = maxFontSize * 1.44; // 144%的行距
        }
    } catch (e) {
        alert("调整行距时发生错误：" + e);
    }
}

// 删除括号
function removeBracketsWithFindReplace(textFrame, bracketsPositions) {
    if (bracketsPositions.length === 0) {
        return; // 如果没有括号位置，直接返回
    }
    try {
        // 按位置从后往前删除，避免索引错位
        for (var i = bracketsPositions.length - 1; i >= 0; i--) {
            var start = bracketsPositions[i].start;
            var end = bracketsPositions[i].end;

            // 验证范围是否有效
            var range = textFrame.characters.itemByRange(start, end);
            // alert("删除范围：" + start + " - " + end);
            if (range.isValid) {
                range.remove();
            }
        }
    } catch (e) {
        alert("删除括号时发生错误：" + e);
    }
}