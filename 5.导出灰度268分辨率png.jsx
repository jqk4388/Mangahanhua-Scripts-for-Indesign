// 获取当前活动文档
var doc = app.activeDocument;

// 获取文档名称并去掉扩展名
var docName = doc.name.replace(/\.[^\.]+$/, '');

// 设置默认导出文件夹
var defaultFolder = new Folder("M:/汉化/PS_PNG");
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
var maxDuration = 5 * 60 * 1000; // 5分钟

if (exportCurrentPageOnly) {
    // 只导出当前页面
    var currentPage = app.activeWindow.activePage;
    var filePath = new File(folder.fsName + "/" + docName + "_" + ("00" + (currentPage.documentOffset + 1)).slice(-3) + ".png");
    // 设置导出选项
    app.pngExportPreferences.pngQuality = PNGQualityEnum.MAXIMUM;
    app.pngExportPreferences.exportResolution = exportResolution;
    app.pngExportPreferences.antiAlias = true;
    app.pngExportPreferences.pngColorSpace = PNGColorSpaceEnum.GRAY;
    app.pngExportPreferences.useDocumentBleeds = false;
    app.pngExportPreferences.simulateOverprint = false;
    app.pngExportPreferences.exportingSpread = false;
    app.pngExportPreferences.pngExportRange = PNGExportRangeEnum.EXPORT_RANGE;
    app.pngExportPreferences.pageString = currentPage.name;

    doc.exportFile(ExportFormat.PNG_FORMAT, filePath, false);
    progressWin.progressBar.value = 1;
    progressWin.progressText.text = "已导出当前页面";
} else {
    // 导出全部页面
    for (var i = 0; i < doc.pages.length; i++) {
        if (userCancelled) {
            alert("操作已取消。");
            exit();
        }

        var currentTime = new Date().getTime();
        if (currentTime - startTime > maxDuration) {
            alert("导出时间超过5分钟，操作已停止。");
            break;
        }

        var page = doc.pages[i];
        var filePath = new File(folder.fsName + "/" + docName + "_" + ("00" + (i + 1)).slice(-3) + ".png");
        
        // 设置导出选项
        app.pngExportPreferences.pngQuality = PNGQualityEnum.MAXIMUM;
        app.pngExportPreferences.exportResolution = exportResolution;
        app.pngExportPreferences.antiAlias = true;
        app.pngExportPreferences.pngColorSpace = PNGColorSpaceEnum.GRAY;
        app.pngExportPreferences.useDocumentBleeds = false;
        app.pngExportPreferences.simulateOverprint = false;
        app.pngExportPreferences.exportingSpread = false;
        app.pngExportPreferences.pngExportRange = PNGExportRangeEnum.EXPORT_RANGE;
        app.pngExportPreferences.pageString = page.name;

        doc.exportFile(ExportFormat.PNG_FORMAT, filePath, false);
        
        // 更新进度条
        progressWin.progressBar.value = i + 1;
        progressWin.progressText.text = "正在导出页面 " + (i + 1) + " / " + doc.pages.length;
    }
}

// 关闭进度条窗口
progressWin.close();

alert("导出完成！");