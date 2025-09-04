// 遍历所有打开的文档，将其保存为idml副本到原文件夹

if (app.documents.length === 0) {
    alert("没有打开的文档。");
} else {
    // 获取当前时间并格式化
    var now = new Date();
    var timeStr = now.getFullYear().toString() +
                 ("0" + (now.getMonth() + 1)).slice(-2) +
                 ("0" + now.getDate()).slice(-2) + "_" +
                 ("0" + now.getHours()).slice(-2) +
                 ("0" + now.getMinutes()).slice(-2);

    var docs = app.documents;
    for (var i = 0; i < docs.length; i++) {
        var doc = docs[i];
        // 获取文件名（去掉扩展名）
        var baseName = doc.name.replace(/\.[^\.]+$/, "");
        // 获取保存路径：已保存的用原文件夹，未保存的用桌面
        var folder;
        if (doc.saved) {
            folder = doc.fullName.parent;
        } else {
            folder = Folder.desktop;
        }
        var idmlFile = new File(folder + "/" + baseName + "_" + timeStr + ".idml");
        doc.exportFile(ExportFormat.INDESIGN_MARKUP, idmlFile);
    }
    alert("所有文档已导出为IDML副本。");
}
