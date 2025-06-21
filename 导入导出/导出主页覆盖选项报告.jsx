// 导出主页覆盖选项报告.jsx

if (app.documents.length === 0) {
    alert("请先打开一个文档。");
} else {
    var doc = app.activeDocument;
    var report = [];
    var i, j, k;
    // 1. 收集所有母版及其元素
    var masterMap = {};
    for (i = 0; i < doc.masterSpreads.length; i++) {
        var master = doc.masterSpreads[i];
        var masterName = master.name;
        masterMap[masterName] = [];
        var items = master.pageItems;
        report.push("主页：" + masterName);
        for (j = 0; j < items.length; j++) {
            masterMap[masterName].push(items[j]);
            var x0=items[j]['geometricBounds']['2'] - items[j]['geometricBounds']['0'];
            var y0=items[j]['geometricBounds']['3'] - items[j]['geometricBounds']['1'];
            report.push("     元素 " + items[j].id + " 已收集，尺寸："+x0 +"，"+ y0);
        }
    }
    // 2. 遍历所有页面，检查主页元素的覆盖、修改、删除
    for (i = 0; i < doc.pages.length; i++) {
        var page = doc.pages[i];
        var pageName = page.name;
        var master = page.appliedMaster;
        if (!master) continue;
        var masterName = master.name;
        var coveredList = [];
        var masterItems = masterMap[masterName];
        if (!masterItems) continue;
        for (j = 0; j < masterItems.length; j++) {
            var mItem = masterItems[j];
            var found = false;
            var changed = false;
            var deleted = true;
            // 检查页面上是否有覆盖该母版元素
            for (k = 0; k < page.pageItems.length; k++) {
                var pItem = page.pageItems[k];
                if (pItem.overriddenMasterPageItem && pItem.overriddenMasterPageItem == mItem) {
                    found = true;
                    // 判断是否有修改（如几何属性、内容等，可根据需要扩展）
                    if (pItem.geometricBounds + '' !== mItem.geometricBounds + '') {
                        changed = true;
                    }
                    deleted = false;
                    break;
                }
            }
            // 检查页面的masterPageItems（母版元素被删除的情况）
            for (k = 0; k < page.masterPageItems.length; k++) {
                var mpi = page.masterPageItems[k];
                if (mpi == mItem && !mpi.isValid) {
                    deleted = true;
                } else if (mpi == mItem) {
                    deleted = false;
                }
            }
            if (found && changed) {
                coveredList.push("修改: " + mItem.id );
            } else if (found) {
                coveredList.push("覆盖: " + mItem.id );
            } else if (deleted) {
                coveredList.push("删除: " + mItem.id );
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
        alert("未检测到覆盖、修改或删除的主页元素。");
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
