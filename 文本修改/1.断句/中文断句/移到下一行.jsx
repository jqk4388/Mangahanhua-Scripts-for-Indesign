main();

function main() {
    // 检查是否有活动文档
    if (app.documents.length === 0) {
        alert("请先打开一个文档！");
        return;
    }

    // 获取当前文档
    var doc = app.activeDocument;

    try {
        // 开始事务
        app.doScript(moveTextToNextLine, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "移动到前一行");
    } catch (error) {
        alert("发生错误：" + error);
    }
}

// 移动文本到下一行的具体实现
function moveTextToNextLine() {
    var doc = app.activeDocument;
    var sel = app.selection;
    if (sel.length !== 1 || !(sel[0].hasOwnProperty("insertionPoints"))) {
        alert("请选中一个文本插入点！");
        return;
    }
    var insertionPoint = sel[0];
    // 在光标处插入换行符
    insertionPoint.contents = "\n";

    // 获取插入点在story中的索引
    var story = insertionPoint.parentStory;
    var ipIndex = insertionPoint.index;

    // 向后查找第一个换行符并删除
    var storyLen = story.characters.length;
    for (var i = ipIndex + 1; i < storyLen; i++) {
        var ch = story.characters[i].contents;
        if (ch === "\r" || ch === SpecialCharacters.FORCED_LINE_BREAK) {
            story.characters[i].remove();
            break;
        }
    }
}