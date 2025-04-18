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
        app.doScript(moveTextToPreviousLine, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "移动到前一行");
    } catch (error) {
        alert("发生错误：" + error);
    }
}

// 移动文本到前一行的具体实现
function moveTextToPreviousLine() {
    // 获取当前选区
    var selection = app.selection[0];
    
    // 检查是否有文本选择
    if (!selection || !(selection instanceof Text || selection instanceof InsertionPoint || selection instanceof TextStyleRange)) {
        alert("请先选择一段文本！");
        return;
    }

    // 获取选中的文本内容
    var selectedText = selection.contents;
    
    // 获取选中文本的起始位置
    var startIndex = selection.insertionPoints[0].index;
    
    // 检查选中文本前面是否有换行符
    var previousChar = selection.parentStory.characters[startIndex - 1].contents;
    
    // 如果前面有换行符
    if (previousChar === "\r" || previousChar === "\n") {
        // 删除前面的换行符
        selection.parentStory.characters[startIndex - 1].remove();
        
        // 在选中文本后添加换行符
        selection.insertionPoints[-1].contents = "\r";
        
        // 重新选中处理后的文本
        selection.select();
    } else {
        alert("选中文本前没有换行符！");
        return;
    }
}