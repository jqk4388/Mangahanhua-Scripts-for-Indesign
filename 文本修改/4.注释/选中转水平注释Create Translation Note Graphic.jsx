// 确保当前选中内容是文本
if (app.selection.length > 0 && app.selection[0].constructor.name === "Text") {
    // 获取文档引用
    var myDocument = app.documents.item(0);

    // 获取当前选中文字所在的页面
    var selectedText = app.selection[0];
    var parentPage = selectedText.parentTextFrames[0].parentPage;

    if (parentPage !== null) {
        // 获取选中的文本内容
        var selectedTextContent = selectedText.contents;
        app.cut();
        // 在选中文字的同一页面创建新的文本框
        var myTextBox = parentPage.textFrames.add({
            geometricBounds: [2, 2, 5, 52] // 设置文本框的几何位置，可以根据需要调整
        });

        // 将选中的文本内容粘贴到新的文本框中
        myTextBox.contents = selectedTextContent;

        // 应用对象样式
        try {
            var myBoxStyle = myDocument.objectStyles.item("水平注释");
            myBoxStyle.name; // 如果样式不存在，会抛出错误
        } catch (myError) {
            // 创建新的对象样式
            myBoxStyle = myDocument.objectStyles.add({ name: "水平注释" });
        }
        myTextBox.applyObjectStyle(myBoxStyle);

        // 将文本框剪切到剪贴板
        myTextBox.select();
        // app.cut();

        alert("注释文本框已创建在当前页面右上角，请自行修改它的位置。");
    } else {
        alert("无法找到选中文字的所在页面。");
    }
} else {
    alert("请先选择一段文本。");
}
