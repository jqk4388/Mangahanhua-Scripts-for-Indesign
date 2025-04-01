function stackWords(obj) {
    obj.contents = obj.contents.split(' ').join(' \n');
}

// 批量处理逻辑
if (app.selection.length > 0) {
    var selection = app.selection;
    for (var i = 0; i < selection.length; i++) {
        var selectedObj = selection[i];
        if (selectedObj.constructor.name === "TextFrame") {
            stackWords(selectedObj);
        }
    }
} else {
    alert("请至少选择一个文本框");
}