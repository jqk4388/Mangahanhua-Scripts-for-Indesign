// Adobe InDesign Script: 自动拼音脚本
// 主要功能：为选中文本框中的括号内容添加拼音注音
// 作者：几千块
// 日期：2025年4月26日
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
    var content = textFrame['parentStory']['contents'];
    var lines = content.split(/\r|\n/);
    var globalBracketsPositions = [];
    var localBracketsPositions = [];

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var brackets = findBrackets(line);

        if (brackets.length > 0) {
            for (var j = 0; j < brackets.length; j++) {
                var bracketInfo = brackets[j];
                var targetText = line.substring(0, bracketInfo.start);
                var rubyText = line.substring(bracketInfo.start + 1, bracketInfo.end);

                // 检查是否有「」或『』包裹内容（允许后面紧跟括号）
                var quoteTypes = [
                    {open: "「", close: "」"},
                    {open: "『", close: "』"},
                    {open: "《", close: "》"},
                    {open: "【", close: "】"},
                    {open: "［", close: "］"},
                    {open: "“", close: "”"},
                    {open: "‘", close: "’"},
                    {open: "〈", close: "〉"}
                ];
                var quoteInfo = null;
                var quoteStart = -1, quoteEnd = -1, quoteInner = "", quoteInnerStart = 0;
                for (var q = 0; q < quoteTypes.length; q++) {
                    var openIdx = targetText.lastIndexOf(quoteTypes[q].open);
                    var closeIdx = targetText.indexOf(quoteTypes[q].close, openIdx + 1);
                    // 只要有一对引号包裹，且closeIdx紧跟在bracketInfo.start前或为末尾
                    if (openIdx !== -1 && closeIdx !== -1 && closeIdx === targetText.length - 1) {
                        quoteInfo = quoteTypes[q];
                        quoteStart = openIdx;
                        quoteEnd = closeIdx;
                        quoteInner = targetText.substring(quoteStart + 1, quoteEnd);
                        quoteInnerStart = quoteStart + 1;
                        break;
                    }
                }

                // 优化：查找整对引号，并判断是否与括号粘连，不粘连则忽略引号
                var quotePairFound = false;
                for (var q = 0; q < quoteTypes.length; q++) {
                    var openIdx = targetText.lastIndexOf(quoteTypes[q].open);
                    if (openIdx !== -1) {
                        var closeIdx = targetText.indexOf(quoteTypes[q].close, openIdx + 1);
                        if (closeIdx !== -1) {
                            // 检查引号是否和括号粘连
                            if (closeIdx === bracketInfo.start - 1) {
                                // 引号和括号粘连，使用引号内内容
                                quoteInfo = quoteTypes[q];
                                quoteStart = openIdx;
                                quoteEnd = closeIdx;
                                quoteInner = targetText.substring(quoteStart + 1, quoteEnd);
                                quoteInnerStart = quoteStart + 1;
                                quotePairFound = true;
                                break;
                            }
                        }
                    }
                }

                var defaultChars, maxChars, startPos;
                if (quotePairFound && quoteInner.length > 0) {
                    defaultChars = maxChars = quoteInner.length;
                    startPos = quoteInnerStart;
                } else {
                    // 没有粘连的引号，全部可加拼音
                    defaultChars = maxChars = targetText.length;
                    startPos = 0;
                }

                // 兼容“只有开引号”情况（如「蜥蜴战士）（测试加拼音）），且开引号和括号紧相连
                if (!quotePairFound &&!quoteInfo) {
                    for (var q = 0; q < quoteTypes.length; q++) {
                        var openIdx = targetText.lastIndexOf(quoteTypes[q].open);
                        if (openIdx !== -1 && openIdx < targetText.length - 1) {
                            // 判断开引号后所有内容是否和括号紧相连
                            if (bracketInfo.start === openIdx + 1 + targetText.substring(openIdx + 1).length) {
                                quoteInner = targetText.substring(openIdx + 1);
                                if (quoteInner.length > 0) {
                                    defaultChars = maxChars = quoteInner.length;
                                    startPos = openIdx + 1;
                                    isQuoteAndBracketAdjacent = true;
                                    break;
                                }
                            }
                        }
                    }
                }

                // 兼容“只有闭引号”情况
                if (!quotePairFound &&!quoteInfo) {
                    for (var q = 0; q < quoteTypes.length; q++) {
                        var closeIdx = targetText.indexOf(quoteTypes[q].close);
                        if (closeIdx !== -1 && closeIdx > 0) {
                            quoteInner = targetText.substring(0, closeIdx);
                            if (quoteInner.length > 0) {
                                defaultChars = maxChars = quoteInner.length;
                                startPos = 0;
                                break;
                            }
                        }
                    }
                }

                // 显示对话框
                var prevLine = i > 0 ? lines[i - 1] : "";
                var userInput = showRubyDialog(
                    quoteInner.length > 0 ? quoteInner : targetText,
                    rubyText,
                    prevLine,
                    defaultChars,
                    maxChars
                );

                if (userInput === null) {
                    return; // 用户取消操作
                }

                if (userInput > 0) {
                    try {
                        // 计算拼音起始位置
                        var applyStart = startPos + (maxChars - userInput);
                        // 若用户输入小于最大值，则从引号内末尾往前数
                        if (userInput !== maxChars) {
                            applyStart = startPos + (maxChars - userInput);
                        } else {
                            applyStart = startPos;
                        }
                        applyRubyToLine(textFrame, i, applyStart, userInput, rubyText);
                    } catch (e) {
                        alert("应用拼音失败：" + e);
                    }

                    if (i > 0) {
                        adjustPreviousLineLeading(textFrame, i);
                    }

                    var currentPosition = 0;
                    for (var k = 0; k < i; k++) {
                        currentPosition += lines[k].length + 1;
                    }
                    globalBracketsPositions.push({
                        start: currentPosition + bracketInfo.start,
                        end: currentPosition + bracketInfo.end
                    });

                    localBracketsPositions.push({
                        lineIndex: i,
                        start: bracketInfo.start,
                        end: bracketInfo.end
                    });
                }
            }
        }
    }

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

function showRubyDialog(targetText, rubyText, prevLine, defaultChars, maxChars) {
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
        var rubyTextRow = column.dialogRows.add();
        rubyTextRow.staticTexts.add({staticLabel: "拼音文本：" + rubyText});
        var inputRow = column.dialogRows.add({alignChildren: 'left'});
        inputRow.staticTexts.add({staticLabel: "注音字符数（0表示不注音）："});
        var charField = inputRow.integerEditboxes.add({
            editValue: defaultChars,
            minWidth: 100
        });
        
        if (dialog.show()) {
            var value = charField.editValue;
            if (value > maxChars) {
                alert("输入的字符数超过了可加拼音的文字长度！");
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