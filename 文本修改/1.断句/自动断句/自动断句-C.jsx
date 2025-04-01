/* 自动断句脚本 - E3兼容优化版 v1.1 */
// 兼容性垫片（ES3环境）
if (!Array.prototype.indexOf) {
    Array.prototype.indexOf = function(item) {
        for (var i = 0; i < this.length; i++) {
            if (this[i] === item) return i;
        }
        return -1;
    };
}
// 配置常量
var CONFIG = {
    PUNCTUATIONS: ["……", "。", "！", "？","?", "，", "；", "—", "?!","!!","！！","!!!","!?","？！"],
    COMMON_CHARS: ["的", "了", "吗", "吧", "呢"],
    FORBIDDEN_STARTS: ["，", "。", "！", "？", "》", "」", "）", "』", "、","：","……", "—"],
    EXCEPTIONS: {
        "的": ["的话", "的确", "的说", "的是"],
        "了": ["了不起", "了然", "了得", "了解"],
        "吗": ["吗啡"],
        "吧": ["吧台", "吧主"],
        "呢": ["呢喃", "呢子"]
    }
};

try {
    if (app.selection.length === 0) {
        alert("请至少选中一个文本框！");
    } else {
        // 遍历所有选中对象
        for (var i = 0; i < app.selection.length; i++) {
            var mySelection = app.selection[i];
            if (mySelection instanceof TextFrame) {
                var myText = mySelection.contents;
                
            // 处理流程
            myText = removeLineBreaks(myText);
            myText = smartPunctuationBreak(myText);
            myText = handleCommonCharacters(myText);
            myText = formatLineStart(myText);
            
            // myText = analyzeAndSplitLongSentences(myText);
            myText = balanceLineLength(myText);

                // 更新文本框内容
                mySelection.contents = myText;
            }
        }
    }
} catch (e) {
    alert("错误: " + e);
}

// 核心函数改写
function smartPunctuationBreak(text) {
    var marks = CONFIG.PUNCTUATIONS;
    for (var i = 0; i < marks.length; i++) {
        var mark = marks[i];
        var escaped = mark.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        var regex = new RegExp("(" + escaped + "+)(?![$\\r])", "g");
        text = text.replace(regex, "$1\r");
    }
    return text.replace(/\r+/g, "\r");
}

function handleCommonCharacters(text) {
    var tempPrefix = "\u1694";
    var tempStorage = [];
    var exceptionKeys = [];
    
    // ES3兼容的对象键获取
    for (var key in CONFIG.EXCEPTIONS) {
        if (CONFIG.EXCEPTIONS.hasOwnProperty(key)) {
            exceptionKeys.push(key);
        }
    }

    // 存储例外词
    for (var k = 0; k < exceptionKeys.length; k++) {
        var currentKey = exceptionKeys[k];
        var words = CONFIG.EXCEPTIONS[currentKey];
        for (var w = 0; w < words.length; w++) {
            var word = words[w];
            var tempKey = tempPrefix + tempStorage.length;
            tempStorage[tempStorage.length] = word;
            text = text.replace(new RegExp(word, "g"), tempKey);
        }
    }

    // 处理常见句末字
    var commonChars = CONFIG.COMMON_CHARS;
    var punctuationArr = [];
    for (var p = 0; p < CONFIG.PUNCTUATIONS.length; p++) {
        punctuationArr.push(CONFIG.PUNCTUATIONS[p].replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
    }
    var punctuationStr = punctuationArr.join("");

    for (var c = 0; c < commonChars.length; c++) {
        var currentChar = commonChars[c];
        // 正则表达式匹配不断句的部分：1. 句末标点在句末 2. 换行符或文本末尾
        var regexPattern = "([^" + punctuationStr + "\\r])" + 
                        "(?=\\r|$)";
        var regex = new RegExp(regexPattern, "g");
        text = text.replace(regex, "$1$2\r");
    }

    // 恢复例外词
    for (var t = 0; t < tempStorage.length; t++) {
        var tempKey = tempPrefix + t;
        var restoreRegex = new RegExp(tempKey, "g");
        text = text.replace(restoreRegex, tempStorage[t]);
    }

    return text;
}

function formatLineStart(text) {
    var lines = text.split("\r");
    for (var i = 0; i < lines.length; i++) {
        while (CONFIG.FORBIDDEN_STARTS.indexOf(lines[i].charAt(0)) > -1) {
            lines[i] = lines[i].substr(1);
        }
    }
    return lines.join("\r");
}

function balanceLineLength(text) {
    var lines = text.split("\r");
    var totalLength = 0;
    for (var i = 0; i < lines.length; i++) {
        totalLength += lines[i].length;
    }
    var avgLength = totalLength / lines.length;
    
    var newLines = [];
    for (var j = 0; j < lines.length; j++) {
        var line = lines[j];
        if (line.length > avgLength * 1.5) {
            newLines.push(splitAtNaturalBreak(line) || splitSmartly(line));
        } else {
            newLines.push(line);
        }
    }
    return newLines.join("\r");
}
/**
 * 智能分割文本行，优先在自然断句处或常见字符处分隔，否则均匀分割
 * 
 * @param {string} line - 需要分割的原始文本行
 * @returns {string} 处理后的文本行，包含换行符\r的分隔
 */
function splitSmartly(line) {
    // 优先在自然断句符号（。！？）后分割，排除行尾的情况
    var naturalBreak = line.match(/[。！？：；](?!$)/);
    if (naturalBreak) {
        return line.replace(naturalBreak[0], naturalBreak[0] + "\r");
    }
    
    // 次优选择在常见字符后分割
    var charBreak = line.match(new RegExp("[" + CONFIG.COMMON_CHARS.join("") + "](?!$)"));
    if (charBreak) {
        return line.replace(charBreak[0], charBreak[0] + "\r");
    }
    
    // 默认情况：将文本均匀分割为最多8个段落
    return splitTextEvenly(line, Math.ceil(line.length / 8));
}

// 移除换行符函数
function removeLineBreaks(text) {
    return text.replace(/(\r|\n)+/g, "");
}
function splitAtNaturalBreak(text) {
    // 优先在标点后换行（保留原逻辑）
    var punctMatch = text.match(/[。！？](?!$)/);
    if (punctMatch) return text.replace(punctMatch[0], punctMatch[0] + "\r");

    // 不断句的词组
    var SEMANTIC_GROUPS = [
        "然后", "但是", "所以", "因为", "如果", "即使", "虽然",
        "一个", "一只", "一些", "这种", "那种","我们", "你们",
        "他们", "她们", "它们", "这个", "那个", "这些", "那些",
        "这里", "那里", "上面", "下面", "前面", "后面",
        "里面", "外面", "旁边", "可是", "中间", "周围", "里面",
        "外面", "旁边", "对面", "中间", "周围", "一起", "同时",
        "之后", "之前", "当时", "现在", "将来", "过去", "未来",
        "一直", "总是", "有时候", "偶尔", "从来", "不再", "再也",
        "已经", "还要", "还会", "还可以", "还要", "还想",
        "还需要", "还应该", "还必须", "还能够", "还可以",
        "还可以", "还想", "还需要", "还应该",
        "汽车", "火车", "飞机", "轮船", "自行车", "摩托车",
        "公交车", "地铁", "出租车", "电动车", "滑板车", "步行",
        "不再", "不想", "不需要", "不应该", "不必须", "不能够",
        "不可以", "不想要", "传说","在下",
        "斋藤一", "浪人","志志雄真实", "志志雄", "真选组","阿薰","大久保利通",
        "绯村","剑心","新选组", "剑客", "剑士", "剑道", "剑术","刽子手","拔刀斋",
        "神谷","道场", "剑心", "斋藤","弥彦","左之助","士族",
        "京都", "江户", "明治", "维新", "幕末", "武士","东京",
        "忍者", "刺客", "刀剑", "武器", "战斗", "战争","战争",
        "神奈川", "横滨", "横须贺", "箱根", "镰仓", "富士山",
        "长崎", "佐贺", "熊本", "福冈", "广岛", "冲绳",
        "四国", "九州", "北海道", "东北", "关东", "关西",
        "中部", "中国", "四国", "九州", "近畿", "东海",
        "西日本", "东日本", "南日本", "北日本", "东南亚",
        "东亚", "东北亚", "东南亚", "南亚", "西亚", "中东",
    ];

    // 最大正向匹配算法
    for (var len = 4; len >= 2; len--) { // 优先匹配长词组
        var maxPos = Math.min(text.length - 1, 15);
        for (var i = maxPos; i >= len; i--) { // 从前15字开始匹配
            var segment = text.substr(i - len, len);
            if (contains(SEMANTIC_GROUPS, segment)) {
                return text.slice(0, i) + "\r" + text.slice(i);
            }
        }
    }
    
    // 无匹配时返回null交给后续处理
    return null;

    // ES3兼容的数组包含判断
    function contains(arr, item) {
        for (var x = 0; x < arr.length; x++) {
            if (arr[x] === item) return true;
        }
        return false;
    }
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