// 导出主页覆盖选项报告.jsx

if (app.documents.length === 0) {
    alert("请先打开一个文档。");
} else {
    var doc = app.activeDocument;
    var report = [];
    var i, j, k;

    for (i = 0; i < doc.pages.length; i++) {
        var page = doc.pages[i];
        var pageName = page.name;
        var master = page.appliedMaster;
        if (!master) continue;

        // 统计母版元素
        var masterItems = master.pageItems;
        var coveredList = [];
        for (j = 0; j < masterItems.length; j++) {
            var mItem = masterItems[j];
            var found = false;
            // 检查页面上是否有覆盖该母版元素
            for (k = 0; k < page.pageItems.length; k++) {
                var pItem = page.pageItems[k];
                if (pItem.parentPage == page && pItem.overriddenMasterPageItem && pItem.overriddenMasterPageItem == mItem) {
                    found = true;
                    break;
                }
            }
            // 检查是否被覆盖或被删除
            if (found) {
                coveredList.push("覆盖: " + mItem.constructor.name + " (母版元素名: " + mItem.name + ")");
            } else {
                // 检查母版元素是否被删除（即页面上没有覆盖且母版元素未显示）
                // InDesign没有直接API判断母版元素被删除，只能通过覆盖与否判断
                // 这里仅统计被覆盖的情况
            }
        }
        // 检查页面的masterPageItems（母版元素被删除的情况）
        var masterPageItems = page.masterPageItems;
        for (j = 0; j < masterPageItems.length; j++) {
            var mpi = masterPageItems[j];
            if (!mpi.isValid) {
                coveredList.push("删除: " + mpi.constructor.name + " (母版元素名: " + mpi.name + ")");
            }
        }
        if (coveredList.length > 0) {
            report.push("页面: " + pageName);
            for (j = 0; j < coveredList.length; j++) {
                report.push("    " + coveredList[j]);
            }
        }
    }

    if (report.length === 0) {
        alert("未检测到覆盖或删除的主页元素。");
    } else {
        var desktopPath = Folder.desktop.fsName;
        var fileName = "主页覆盖选项报告.txt";
        var filePath = desktopPath + "/" + fileName;
        var file = new File(filePath);
        file.encoding = "UTF-8";
        file.open("w");
        file.write(report.join("\r\n"));
        file.close();
        alert("主页覆盖选项报告已导出到桌面：" + filePath);
    }
}
