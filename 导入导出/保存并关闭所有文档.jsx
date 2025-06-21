var docs = app.documents;
var i, doc, file, desktopPath, filename, ext, count = 1;

// 获取桌面路径
if ($.os.indexOf("Windows") !== -1) {
    desktopPath = Folder.desktop.fsName + "\\";
} else {
    desktopPath = Folder.desktop.fsName + "/";
}

for (i = docs.length - 1; i >= 0; i--) {
    doc = docs[i];
    if (doc.saved) {
        // 已保存过的文档
        doc.save();
        doc.close();
    } else {
        // 新建未保存的文档，直接用baseName
        ext = ".indd";
        var baseName = doc.name.replace(/\.[^\.]+$/, "");
        filename = baseName + ext;
        file = new File(desktopPath + filename);
        doc.save(file);
        doc.close();
    }
}
alert("所有文档已保存并关闭。");
