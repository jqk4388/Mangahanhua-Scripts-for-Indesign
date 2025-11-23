// Adobe InDesign Script: 自动加着重号
// 功能：对选中的文本框内括号中的文字加着重号，并调整相关行距
// 作者：几千块
// 日期：2025年11月23日
// 版本：1.3
// 加点位置右上
#include "../Library/KTUlib.jsx"
QUOTES_star = "\u3010"; //  【
QUOTES_end = "\u3011"; // 】
// 主函数
function main() {
    try {
        // 检查是否有选中的文本框
        if (app.selection.length === 0) {
            alert("请先选择一个或多个文本框！");
            return;
        }

        // 遍历选中的对象
        for (var i = 0; i < app.selection.length; i++) {
            var selectedObject = app.selection[i];
            if (selectedObject.constructor.name === "TextFrame") {
                processTextFrame(selectedObject);
            } else if (selectedObject.constructor.name === "Text"|| selectedObject.constructor.name === "Character"|| selectedObject.constructor.name === "TextStyleRange"|| selectedObject.constructor.name === "Word") { 
                processText(selectedObject);
                var textFrame = selectedObject['parentTextFrames'][0];
                processTextFrame(textFrame);
            } else {
                alert("选中的对象不是文本框或文本！");
            }
        }
    } catch (error) {
        alert("发生错误：" + error.message);
    }
}

function processText(text) {
    try {
        var contents = QUOTES_star + text.contents + QUOTES_end;
        text.contents = contents;
        return;
    } catch (error) {
        alert("加括号发生错误：" + error.message);
    }
}
// 处理单个文本框
function processTextFrame(textFrame) {
    textFrame.fit(FitOptions.FRAME_TO_CONTENT);
    try {
        var text = textFrame['parentStory']['contents'];
        var matches = findBrackets(text);

        if (matches.length === 0) {
            alert("未找到括号内容！");
            return;
        }

        var globalPositions = [];
        var linePositions = [];

        for (var j = 0; j < matches.length; j++) {
            var match = matches[j];
            var startIndex = match.start;
            var endIndex = match.end;

            // 记录全局位置和行位置
            globalPositions.push({ start: startIndex, end: endIndex });
            linePositions.push(getLinePosition(textFrame, startIndex));

            // 添加着重号
            addEmphasis(textFrame, startIndex + 1, endIndex - 1);

            // 调整行距
            adjustLeading(textFrame, startIndex, endIndex);
        }

        // 删除括号
        deleteBrackets(textFrame, globalPositions);

        //使框架适合内容
        if (textFrame.overflows) {
            textFrame.fit(FitOptions.FRAME_TO_CONTENT);
        }

    } catch (error) {
        alert("处理文本框时发生错误：" + error.message);
    }
}

// 查找括号内容
function findBrackets(text) {
    var results = [];
    var regex = /(\(|（|\[|【|｛|{|<)(.*?)(\)|）|\]|】|｝|}|>)/g;
    var match;

    while ((match = regex.exec(text)) !== null) {
        results.push({
            start: match.index,
            end: match.index + match[0].length
        });
    }

    return results;
}

// 获取字符所在的行位置
function getLinePosition(textFrame, charIndex) {
    try {
        var lines = textFrame.lines;
        for (var i = 0; i < lines.length; i++) {
            if (charIndex >= lines[i].index && charIndex < lines[i].index + lines[i].length) {
                return i + 1; // 行号从1开始
            }
        }
    } catch (error) {
        alert("获取行位置时发生错误：" + error.message);
    }
    return -1; // 未找到
}

// 添加着重号
function addEmphasis(textFrame, startIndex, endIndex) {
    try {
        for (var i = startIndex; i <= endIndex; i++) {
            var character = textFrame.characters[i];
            character.kentenKind = KentenCharacter.KENTEN_SMALL_BLACK_CIRCLE;
            character.kentenPosition = RubyKentenPosition.ABOVE_RIGHT; // 位置：右上
        }
    } catch (error) {
        alert("添加着重号时发生错误：" + error.message);
    }
}

// 调整行距
function adjustLeading(textFrame, startIndex, endIndex) {
    try {
        var startLine = getLinePosition(textFrame, startIndex);
        var endLine = getLinePosition(textFrame, endIndex);

        if (startLine === -1 || endLine === -1) {
            // alert("无法获取括号所在的行位置！");
            return;
        }

        // 确保行号从小到大
        var minLine = Math.min(startLine, endLine);
        var maxLine = Math.max(startLine, endLine);

        // 遍历从起始行的前一行到结束行的所有行
        for (var lineIndex = minLine - 1; lineIndex <= maxLine - 1; lineIndex++) {
            if (lineIndex > 0) {
                var currentLine = textFrame.lines[lineIndex - 1];
                var maxFontSize = getMaxFontSize(currentLine);
                currentLine.leading = maxFontSize * 1.44;
            }
        }
    } catch (error) {
        alert("调整行距时发生错误：" + error.message);
    }
}

// 获取一行中最大的字符字号
function getMaxFontSize(line) {
    var maxFontSize = 0;
    try {
        for (var i = 0; i < line.characters.length; i++) {
            var fontSize = line.characters[i].pointSize;
            if (fontSize > maxFontSize) {
                maxFontSize = fontSize;
            }
        }
    } catch (error) {
        alert("获取最大字号时发生错误：" + error.message);
    }
    return maxFontSize;
}

// 删除括号
function deleteBrackets(textFrame, positions) {
    try {
        for (var i = positions.length - 1; i >= 0; i--) {
            var position = positions[i];
            // 删除右括号
            textFrame.characters[position.end - 1].remove();
            // 删除左括号
            textFrame.characters[position.start].remove();
        }
    } catch (error) {
        alert("删除括号时发生错误：" + error.message);
    }
}

// 执行主函数
KTUDoScriptAsUndoable(function() {main()}, "A自动加点");