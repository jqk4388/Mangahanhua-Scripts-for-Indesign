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
    
    // 检查是否有有效选择
    if (!selection || !(selection instanceof Text || selection instanceof InsertionPoint || selection instanceof TextStyleRange)) {
        // alert("无效的选择！");
        return;
    }

    // 获取选区起始位置
    var startIndex = selection.insertionPoints[0].index;
    
    // 检查前一个字符是否为换行符
    var previousChar = selection.parentStory.characters[startIndex - 1].contents;
    if (previousChar !== SpecialCharacters.FORCED_LINE_BREAK && previousChar !== "\r") {
        // alert("前面没有换行符！");
        return;
    }

    // 处理选中的文本或单个字符
    var textToMove;
    if (selection.contents.length > 0) {
        // 有选中文本的情况
        textToMove = selection.contents;
        selection.contents = ""; // 清除原位置的文本
    } else {
        // 光标位置的情况，获取后面的一个字符
        try {
            textToMove = selection.parentStory.characters[startIndex].contents;
            // 检查是否为特殊字符
            if (textToMove === SpecialCharacters.ELLIPSIS_CHARACTER){
                textToMove = "……" ;
                selection.parentStory.characters[startIndex].remove();
                selection.parentStory.characters[startIndex].remove();
            }else if (textToMove === SpecialCharacters.EM_DASH){
                textToMove = "—";
                selection.parentStory.characters[startIndex].remove();
            }
            
        } catch (e) {
            // alert("光标后没有字符可移动！");
            return;
        }
    }

    // 删除前面的换行符
    selection.parentStory.characters[startIndex - 1].remove();
    
    // 在原位置插入文本和换行符
    var insertionPoint = selection.parentStory.insertionPoints[startIndex - 1];
    insertionPoint.contents = textToMove + "\r";
    
    // 将光标移动到处理后的位置
    insertionPoint.select();
}