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
        for (j = 0; j < items.length; j++) {
            masterMap[masterName].push(items[j]);
        }
    }
    // 2. 创建元素状态映射
    var elementMap = {};
    for (i = 0; i < doc.pages.length; i++) {
        var page = doc.pages[i];
        var pageNumber = i + 1; // 页面编号从1开始
        var master = page.appliedMaster;
        if (!master) continue;
        var masterName = master.name;
        var masterItems = masterMap[masterName];
        if (!masterItems) continue;
        for (j = 0; j < masterItems.length; j++) {
            var mItem = masterItems[j];
            var elementId = mItem.id;
            if (!elementMap[elementId]) {
                elementMap[elementId] = { status: "", pages: [] };
            }
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
            var status = "";
            if (found && changed) {
                status = "修改";
            } else if (found) {
                status = "覆盖";
            } else if (deleted) {
                status = "删除";
            }
            if (status) {
                elementMap[elementId].status = status;
                elementMap[elementId].pages.push(pageNumber);
            }
        }
    }
    // 3. 生成报告
    for (var elementId in elementMap) {
        var info = elementMap[elementId];
        if (info.pages.length > 0) {
            report.push("元素 " + elementId + " (" + info.status + ") 所在页面：" + info.pages.join(","));
        }
    }
    if (report.length === 0) {
        alert("未检测到覆盖、修改或删除的主页元素。");
    } else {
        var desktopPath = Folder.desktop.fsName;
        var fileName = "主页覆盖选项报告_" + getDateString() + ".txt";
        var filePath = desktopPath + "/" + fileName;
        var file = new File(filePath);
        file.encoding = "UTF-8";
        file.open("w");
        file.write(report.join("\r\n"));
        file.close();
        alert("主页覆盖选项报告已导出到桌面：" + filePath);
    }
}

// 获取格式化的日期字符串
function getDateString() {
    var now = new Date();
    return now.getFullYear() + 
           padZero(now.getMonth() + 1) + 
           padZero(now.getDate()) + "_" + 
           padZero(now.getHours()) + 
           padZero(now.getMinutes()) + 
           padZero(now.getSeconds());
}
// 数字补零函数
function padZero(num) {
    return (num < 10) ? "0" + num : num;
}    