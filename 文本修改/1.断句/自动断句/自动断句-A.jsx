/* 自动断句脚本 - 模块化版本 */

try {
    if (app.selection.length === 0) {
        alert("请至少选中一个文本框！");
    } else {
        // 遍历所有选中对象
        for (var i = 0; i < app.selection.length; i++) {
            var mySelection = app.selection[i];
            if (mySelection instanceof TextFrame) {
                var myText = mySelection.contents;
                
                // 调用模块化函数处理文本
                myText = removeLineBreaks(myText);
                myText = addLineBreaksAfterPunctuation(myText);
                myText = addLineBreaksAfterCommonCharacters(myText);
                myText = splitTextIfSingleLine(myText);
                myText = analyzeAndSplitLongSentences(myText);

                // 更新文本框内容
                mySelection.contents = myText;
            }
        }
    }
} catch (e) {
    alert("错误: " + e);
}

/**
 * 去除所有换行符
 * @param {string} text - 输入文本
 * @returns {string} - 处理后的文本
 */
function removeLineBreaks(text) {
    return text.replace(/(\r|\n)+/g, "");
}

/**
 * 在中文句末标点处添加换行符
 * @param {string} text - 输入文本
 * @returns {string} - 处理后的文本
 */
function addLineBreaksAfterPunctuation(text) {
    var punctuationMarks = ["……", "—", "，", "。", "！", "!"];
    // Use for loop instead of forEach for better compatibility with ExtendScript
    for (var i = 0; i < punctuationMarks.length; i++) {
        var mark = punctuationMarks[i];
        var regex = new RegExp("(" + mark + ")(?!$)", "g");
        text = text.replace(regex, "$1\r");
    };
    return text;
}

/**
 * 在常见的断句汉字处添加换行符
 * @param {string} text - 输入文本
 * @returns {string} - 处理后的文本
 */
function addLineBreaksAfterCommonCharacters(text) {
    var commonCharacters = ["的", "了", "吗", "吧", "呢"];
    var exceptions = [
        "的话", "的确", "的说", "的是",
        "了不起", "了然", "了得", "了解",
        "吗啡",
        "吧台", "吧主",
        "呢喃", "呢子"
    ];
    
    // First replace exceptions with temporary markers
    for (var i = 0; i < exceptions.length; i++) {
        var exception = exceptions[i];
        var tempMarker = "#TEMP" + i + "#";
        text = text.replace(new RegExp(exception, "g"), tempMarker);
    }

    // Add line breaks after common characters
    for (var i = 0; i < commonCharacters.length; i++) {
        var character = commonCharacters[i];
        var regex = new RegExp("([^，。！\r])(" + character + ")(?![，。！的了吗吧呢])", "g");
        text = text.replace(regex, "$1$2\r");
    }

    // Restore exceptions
    for (var i = 0; i < exceptions.length; i++) {
        var tempMarker = "#TEMP" + i + "#";
        text = text.replace(new RegExp(tempMarker, "g"), exceptions[i]);
    }

    return text;
}

/**
 * 如果处理后没有断行，则根据字数选择断为多行
 * @param {string} text - 输入文本
 * @returns {string} - 处理后的文本
 */
function splitTextIfSingleLine(text) {
    if (text.indexOf("\r") === -1) {
        var len = text.length;
        var lineCount;

        // 根据字符长度确定行数
        if (len <= 5) {
            lineCount = 2;
        } else if (len <= 12) {
            lineCount = 3;
        } else if (len <= 20) {
            lineCount = 4;
        } else {
            lineCount = Math.ceil(len / (len <= 20 ? 5 : (len - 20) / 4 + 4));
            if (lineCount < 4) lineCount = 4;
        }

        // 调用均分文本函数
        text = splitTextEvenly(text, lineCount);
    }
    return text;
}
//均分文本函数
function splitTextEvenly(text, lineCount) {
    var len = text.length;
    var perLine = Math.floor(len / lineCount);
    var remainder = len % lineCount;
    var newText = "";
    var offset = 0;
    for (var i = 0; i < lineCount; i++) {
    var extra = i >= (lineCount - remainder) ? 1 : 0;
    newText += text.substr(offset, perLine + extra);
    offset += perLine + extra;
    if (i < lineCount - 1) {
        newText += "\r";
    }
    }
    return newText;
}
/**
 * 分析句子并对长于平均长度的句子进行再次断句，同时合并连续短句
 * @param {string} text - 输入文本
 * @returns {string} - 处理后的文本
 */
function analyzeAndSplitLongSentences(text) {
    var sentences = text.split("\r");
    var totalLength = 0;
    var sentenceCount = sentences.length;

    // 计算平均句子长度
    for (var i = 0; i < sentences.length; i++) {
        totalLength += sentences[i].length;
    }
    var avgLength = Math.floor(totalLength / sentenceCount);
    var longThreshold = avgLength + 2; // 定义长句阈值
    var shortThreshold = avgLength + -2; // 定义短句合并阈值

    var newText = "";
    var tempShortSentence = ""; // 用于暂存短句

    // 处理每个句子
    for (var i = 0; i < sentences.length; i++) {
        var sentence = sentences[i];

        if (sentence.length > longThreshold) {
            // 如果有暂存的短句，先合并到结果中
            if (tempShortSentence) {
                newText += tempShortSentence + "\r";
                tempShortSentence = "";
            }
            // 对长句调用均分函数
            sentence = splitTextEvenly(sentence, 3);
            newText += sentence + "\r";
        } else if (sentence.length < shortThreshold) {
            // 暂存短句
            tempShortSentence += sentence;
        } else {
            // 如果有暂存的短句，先合并到结果中
            if (tempShortSentence) {
                newText += tempShortSentence + "\r";
                tempShortSentence = "";
            }
            // 添加当前句子
            newText += sentence + "\r";
        }
    }

    // 如果最后还有暂存的短句，合并到结果中
    if (tempShortSentence) {
        newText += tempShortSentence + "\r";
    }

    return newText;
}