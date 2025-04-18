/**
 * 合并jieba分词结果为漫画断句段落
 * @param {string} text - jieba分词结果，词组用\n分隔
 * @returns {string} 合并后的段落
 */
function mergeJiebaLines(text) {
    // 停顿标点和句末助词
    var PAUSE_PUNCT = ["……", "。", "！", "？", "?", "，", "；", "、", "—", "?!", "!!", "！！", "!!!", "!?", "？！"];
    var END_PARTICLE = ["的", "了", "吗", "吧", "呢"];
    // 绝对不能分开的词（可外挂词典）
    var ABSOLUTE_UNBREAKABLE = ["……", "？！", "!!", "！！", "!!!", "?!", "!?", "——"];
    //不可分割的词
    var UNBREAKABLE_WORDS = ["这双","这个","这把","这件","这条","这块","这张","这部","这台","这辆","这本","这支","这颗","这套","这幅","这首","这篇",
    "会不会", "能不能", "要不要", "好不好", "对不对", "是不是", "有没有", "行不行", "该不该", "肯不肯", "愿不愿", "敢不敢", "愧不愧", "怪不得", "看样子", "要不然", "算起来", "说不定", "无所谓", "差不多", "没关系", "不管怎样", "不管怎么说", "不管怎么样", "不管怎么说", "不管怎样", "不管怎么说","还需要", "还应该", "还必须", "还能够", "还可以","拔刀斋",
    "和平年代","是人是鬼","二八佳人", "黄金时代", "青春年少", "柴米油盐", "针头线脑", "猫哭老鼠", "虚情假意", "张王李赵", "张三李四", "大街小巷", "左邻右舍", "鸡毛蒜皮", "说三道四", "男女老少", "生老病死", "妖魔鬼怪", "魑魅魍魉", "东南西北", "春夏秋冬", "油盐酱醋", "琴棋书画", "笔墨纸砚", "风花雪月", "诗词歌赋", "山川河流", "风土人情", "衣食住行", "车水马龙", "烟火气息", "人来人往", "天南地北", "天上人间", "天长地久", "天翻地覆", "天真无邪", "天高地厚", "天经地义", "天马行空", "天女散花", "天伦之乐", "天衣无缝", "天灾人祸", "天真烂漫"];
    // 若有外挂词典文件，可在此处加载并合并到ABSOLUTE_UNBREAKABLE
    for (var u = 0; u < UNBREAKABLE_WORDS.length; u++) {
        if (ABSOLUTE_UNBREAKABLE.indexOf(UNBREAKABLE_WORDS[u]) < 0) {
            ABSOLUTE_UNBREAKABLE.push(UNBREAKABLE_WORDS[u]);
        }
    }
    // 1. 词组数组
    var words = text.replace(/\r/g, '').split('\n');
    // 新增：合并括号/引号包裹的词
    var LEFT_BRACKETS = ['（', '【', '「', '『', '《', '〈', '“', '"', '‘', "'"];
    var RIGHT_BRACKETS = ['）', '】', '」', '』', '》', '〉', '”', '"', '’', "'"];
    var mergedBracketWords = [];
    for (var i = 0; i < words.length; ) {
        // 检查左括号+词+右括号的模式
        if (
            i + 2 < words.length &&
            LEFT_BRACKETS.indexOf(words[i]) >= 0 &&
            RIGHT_BRACKETS.indexOf(words[i + 2]) >= 0
        ) {
            mergedBracketWords.push(words[i] + words[i + 1] + words[i + 2]);
            i += 3;
        } else {
            mergedBracketWords.push(words[i]);
            i++;
        }
    }
    words = mergedBracketWords;
    // 1.1 恢复被分割的绝对不能分割的词
    var i, j, found, phrase, len, idx;
    var maxPhraseLen = 0;
    for (i = 0; i < ABSOLUTE_UNBREAKABLE.length; i++) {
        if (ABSOLUTE_UNBREAKABLE[i].length > maxPhraseLen) {
            maxPhraseLen = ABSOLUTE_UNBREAKABLE[i].length;
        }
    }
    var mergedWords = [];
    for (i = 0; i < words.length;) {
        found = false;
        // 尝试最长的绝对不可分割词
        for (len = Math.min(maxPhraseLen, words.length - i); len > 1; len--) {
            phrase = "";
            for (j = 0; j < len; j++) {
                phrase += words[i + j];
            }
            if (ABSOLUTE_UNBREAKABLE.indexOf(phrase) >= 0) {
                mergedWords.push(phrase);
                i += len;
                found = true;
                break;
            }
        }
        if (!found) {
            mergedWords.push(words[i]);
            i++;
        }
    }
    words = mergedWords;
    // 2. 合并停顿标点到前一个词
    var merged = [];
    for (var i = 0; i < words.length; i++) {
        var w = words[i];
        // 合并标点到前一个词
        if (PAUSE_PUNCT.indexOf(w) >= 0 && merged.length > 0) {
            merged[merged.length - 1] += w;
        } else if (w === "" && merged.length > 0) {
            // 跳过空词
            continue;
        } else {
            merged.push(w);
        }
    }
    // 3. 合并句末助词
    var merged2 = [];
    for (var j = 0; j < merged.length; j++) {
        var w2 = merged[j];
        if (END_PARTICLE.indexOf(w2) >= 0 && merged2.length > 0) {
            merged2[merged2.length - 1] += w2;
        } else {
            merged2.push(w2);
        }
    }
    // 4. 统计总字数
    var totalLen = 0;
    for (var k = 0; k < merged2.length; k++) {
        totalLen += merged2[k].length;
    }
    // 5. 计算行数和理想行长
    var minLines = 1;
    var maxLineLen = 7; // 漫画气泡常用最大行长
    var lines = Math.max(minLines, Math.ceil(totalLen / maxLineLen));
    var idealLen = Math.round(totalLen / lines);
    // 6. 分行
    var result = [];
    var line = "";
    var lineLen = 0;
    for (var m = 0; m < merged2.length; m++) {
        var word = merged2[m];
        // 绝对不能分开的词，单独成段或与前后词合并时整体移动
        if (ABSOLUTE_UNBREAKABLE.indexOf(word) >= 0) {
            if (lineLen > 0) {
                result.push(line);
                line = "";
                lineLen = 0;
            }
            result.push(word);
            continue;
        }
        // 如果加上当前词超出理想长度2字，且当前行不空，则换行
        if (lineLen > 0 && (lineLen + word.length > idealLen + 1)) {
            result.push(line);
            line = "";
            lineLen = 0;
        }
        // 长词组不能拆分
        if (lineLen > 0 && word.length >= idealLen + 2) {
            result.push(line);
            result.push(word);
            line = "";
            lineLen = 0;
            continue;
        }
        // 合并到当前行
        line += word;
        lineLen += word.length;
        // 如果刚好到理想长度附近，且不是最后一个词，则换行
        if (lineLen >= idealLen && m < merged2.length - 1) {
            result.push(line);
            line = "";
            lineLen = 0;
        }
    }
    if (lineLen > 0) {
        result.push(line);
    }
    // 7. 处理标点后多余换行
    for (var n = 0; n < result.length - 1; n++) {
        var lastChar = result[n].charAt(result[n].length - 1);
        // 如果以标点结尾，下一行不是标点，则用\r，否则用\n
        if (PAUSE_PUNCT.join('').indexOf(lastChar) >= 0) {
            result[n] += "\r";
        } else {
            result[n] += "\n";
        }
    }
    // 8. 合并结果
    return result.join('').replace(/[\r\n]+$/,"");
}