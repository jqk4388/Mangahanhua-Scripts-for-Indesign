// 检查是否有选中的文本框
var selection = app.selection;
if (selection.length == 0) {
    alert("请先选中一个文本框再运行脚本。");
    exit();
}

// 检查选中的是否为文本框
var textFrame = selection[0];
if (!(textFrame instanceof TextFrame)) {
    alert("请确保选中的是一个文本框。");
    exit();
}

// 检查文本框是否包含内容
if (textFrame.contents === "") {
    alert("选中的文本框为空，无法执行操作。");
    exit();
}

// 弹出输入框获取用户输入
var userText = prompt("请输入要添加到每行末尾的文字：", "");

// 检查用户是否取消输入或输入为空
if (userText === null) {
    // 用户点击了取消
    exit();
}
if (userText === "") {
    alert("输入内容不能为空。");
    exit();
}

try {
    // 获取文本框中的所有段落
    var paragraphs = textFrame.parentStory.paragraphs.everyItem().getElements();
    
    // 遍历每个段落，在末尾添加用户输入的文字
    for (var i = 0; i < paragraphs.length; i++) {
        // 确保不是空段落才添加内容
        if (paragraphs[i].contents.replace(/\r/g, "") !== "") {
            // 移除段落末尾的换行符，添加用户文本，再恢复换行符
            paragraphs[i].contents = paragraphs[i].contents.replace(/\r$/, "") + userText + "\r";
        }
    }
    
    alert("操作完成！已在每一行末尾添加指定文字。");
} catch (e) {
    alert("执行过程中发生错误：" + e.message);
}
