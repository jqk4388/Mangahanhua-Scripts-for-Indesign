// 主函数，执行整个流程
function main() {
    var selectedFrames = getSelectedTextFrames();
    if (selectedFrames.length === 0) {
        alert("请先选择一个或多个文本框架。");
        return;
    }

    // 新增：获取当前页码
    var activePage = app.activeWindow.activePage;
    if (!activePage) {
        alert("无法获取当前页面");
        return;
    }

    // 修改：传递当前页面参数
    var polygonFrame = createPolygonFrame(selectedFrames[0], activePage);
    cutSelectedFrames(selectedFrames);
    pasteContentIntoFrame(polygonFrame);
}

// 新增：获取当前页码函数
function getCurrentPageNumber() {
    return app.activeWindow.activePage.documentOffset + 1;
}

// 获取当前用户选中的文本框架
function getSelectedTextFrames() {
    var selectedItems = app.selection;
    var textFrames = [];
    for (var i = 0; i < selectedItems.length; i++) {
        if (selectedItems[i] instanceof TextFrame) {
            textFrames.push(selectedItems[i]);
        }
    }
    return textFrames;
}

// 剪切选中的文本框架，若有多个则进行编组
function cutSelectedFrames(textFrames) {
    if (textFrames.length > 1) {
        var group = app.activeDocument.groups.add(textFrames);
        app.select(group); // 选中编组
        app.cut();
    } else {
        app.select(textFrames[0]); // 选中单个文本框架
        app.cut();
    }
}

// 新建一个多边形框架
function createPolygonFrame(referenceFrame, activePage) {
    var doc = app.activeDocument;
    var bounds = referenceFrame.geometricBounds;
    
    // 确保在当前页面创建
    var polygon = activePage.polygons.add({
        geometricBounds: bounds,
        strokeWeight: 0,
        fillColor: doc.swatches.item("None"),
        strokeColor: doc.swatches.item("None")
    });
    
    polygon.convertShape(1129533519,4,0);
    return polygon;
}

// 将剪切的内容粘贴到指定框架内
function pasteContentIntoFrame(targetFrame) {
    app.select(targetFrame)
    alert("已经创立新框架，请按Ctrl+Alt+V贴入内部！");
}

// 运行主函数
main();
