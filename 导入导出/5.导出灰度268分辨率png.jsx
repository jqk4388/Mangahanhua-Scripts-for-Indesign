// 获取当前活动文档
var doc = app.activeDocument;

// 获取文档名称并去掉扩展名、日语假名，并将带圈数字转换为阿拉伯数字
var docName = doc.name
    .replace(/\.[^\.]+$/, '') // 去掉扩展名
    .replace(/[\u3040-\u30ff]+/g, '') // 去掉日语假名
    .replace(/[\u2460-\u2473]/g, function(m) {
        // 带圈数字Unicode范围：\u2460(①)-\u2473(⑳)
        // ①是1，⑳是20
        return (m.charCodeAt(0) - 0x245F).toString();
    });

// 设置默认导出文件夹
var defaultFolder = new Folder("M:/汉化/PS_PNG");
if (!defaultFolder.exists) {
    defaultFolder = Folder.desktop;
}
var folder = Folder.selectDialog("选择导出文件夹", defaultFolder);
if (folder == null) {
    alert("未选择导出文件夹，操作取消。");
    exit();
}

// 新增：选择导出范围
var exportCurrentPageOnly = false;
if (confirm("是否只导出当前页面？\n\n选择“是”导出当前页面，选择“否”导出全部页面。")) {
    exportCurrentPageOnly = true;
}

// 新增：是否修改分辨率
var exportResolution = 268;
if (confirm("是否要修改导出分辨率？\n\n选择“是”输入新分辨率，选择“否”使用默认268。")) {
    while (true) {
        var inputRes = prompt("请输入导出分辨率（正整数，留空为268）：", "268");
        if (inputRes === null) {
            exportResolution = 268;
            break;
        }
        inputRes = inputRes.replace(/^\s+|\s+$/g, "");
        if (inputRes === "") {
            exportResolution = 268;
            break;
        }
        var numRes = parseInt(inputRes, 10);
        if (!isNaN(numRes) && numRes > 0) {
            exportResolution = numRes;
            break;
        } else {
            alert("请输入正整数！");
        }
    }
}

// 创建进度条窗口
var progressWin = new Window("palette", "导出进度", [150, 150, 600, 270]);
progressWin.progressBar = progressWin.add("progressbar", [20, 20, 430, 40], 0, doc.pages.length);
progressWin.progressText = progressWin.add("statictext", [20, 50, 430, 70], "正在导出页面...");
var cancelButton = progressWin.add("button", [20, 90, 120, 110], "取消", {name: "cancel"});

var userCancelled = false;
cancelButton.onClick = function() {
    userCancelled = true;
    progressWin.close();
};

progressWin.show();

// 获取开始时间
var startTime = new Date().getTime();
var maxDuration = 50 * 60 * 1000; // 50分钟

// 检查并刷新页面链接函数
function checkAndUpdateLinksForPage(page) {
    var pageLinks = [];
    // 收集该页面上的所有链接
    for (var i = 0; i < doc.links.length; i++) {
        var link = doc.links[i];
        try {
            // 判断链接是否属于该页面
            if (link.parent.parentPage && link.parent.parentPage.id == page.id) {
                pageLinks.push(link);
            }
        } catch (e) {
            // 某些链接可能没有parentPage属性，忽略
        }
    }
    // 检查缺失
    var missingLinks = [];
    var outOfDateLinks = [];
    for (var j = 0; j < pageLinks.length; j++) {
        var l = pageLinks[j];
        if (l.status == LinkStatus.LINK_MISSING) {
            missingLinks.push(l.name);
        } else if (l.status == LinkStatus.LINK_OUT_OF_DATE) {
            outOfDateLinks.push(l);
        }
    }
    if (missingLinks.length > 0) {
        alert("页面 " + page.name + " 存在缺失链接，已跳过导出。\n缺失文件:\n" + missingLinks.join("\n"));
        return false;
    }
    // 刷新未更新的链接
    for (var k = 0; k < outOfDateLinks.length; k++) {
        try {
            outOfDateLinks[k].update();
        } catch (e) {
            alert("刷新链接失败: " + outOfDateLinks[k].name);
        }
    }
    return true;
}

if (exportCurrentPageOnly) {
    // 只导出当前页面
    var currentPage = app.activeWindow.activePage;
        if (checkAndUpdateLinksForPage(currentPage)) {
            var filePath = new File(folder.fsName + "/" + docName + "_" + currentPage.appliedSection.name + ("00" + (currentPage.name)).slice(-3) + ".png");
            // 设置导出选项
            app.pngExportPreferences.pngQuality = PNGQualityEnum.MAXIMUM;
            app.pngExportPreferences.exportResolution = exportResolution;
            app.pngExportPreferences.antiAlias = true;
            app.pngExportPreferences.pngColorSpace = PNGColorSpaceEnum.GRAY;
            app.pngExportPreferences.useDocumentBleeds = false;
            app.pngExportPreferences.simulateOverprint = false;
            app.pngExportPreferences.exportingSpread = false;
            app.pngExportPreferences.pngExportRange = PNGExportRangeEnum.EXPORT_RANGE;
            app.pngExportPreferences.pageString = currentPage.appliedSection.name + currentPage.name;

            doc.exportFile(ExportFormat.PNG_FORMAT, filePath, false);
            progressWin.progressBar.value = 1;
            progressWin.progressText.text = "已导出当前页面";
        }
} else {
    // 导出全部页面
    for (var i = 0; i < doc.pages.length; i++) {
        if (userCancelled) {
            alert("操作已取消。");
            exit();
        }

        var currentTime = new Date().getTime();
        if (currentTime - startTime > maxDuration) {
            alert("导出时间超过50分钟，操作已停止。");
            break;
        }

        var page = doc.pages[i];
        if (!checkAndUpdateLinksForPage(page)) {
            progressWin.progressBar.value = i + 1;
            progressWin.progressText.text = "跳过页面 " + (i + 1) + " / " + doc.pages.length;
            continue;
        }
        var filePath = new File(folder.fsName + "/" + docName + "_" + page.appliedSection.name + ("00" + (page.name)).slice(-3) + ".png");
        
        // 设置导出选项
        app.pngExportPreferences.pngQuality = PNGQualityEnum.MAXIMUM;
        app.pngExportPreferences.exportResolution = exportResolution;
        app.pngExportPreferences.antiAlias = true;
        app.pngExportPreferences.pngColorSpace = PNGColorSpaceEnum.GRAY;
        app.pngExportPreferences.useDocumentBleeds = false;
        app.pngExportPreferences.simulateOverprint = false;
        app.pngExportPreferences.exportingSpread = false;
        app.pngExportPreferences.pngExportRange = PNGExportRangeEnum.EXPORT_RANGE;
        app.pngExportPreferences.pageString = page.appliedSection.name + page.name;

        doc.exportFile(ExportFormat.PNG_FORMAT, filePath, false);
        
        // 更新进度条
        progressWin.progressBar.value = i + 1;
        progressWin.progressText.text = "正在导出页面 " + (i + 1) + " / " + doc.pages.length;
    }
}

// 关闭进度条窗口
progressWin.close();

alert("导出完成！");