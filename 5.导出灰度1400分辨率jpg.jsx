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

// 设置导出选项
app.jpegExportPreferences.jpegQuality = JPEGOptionsQuality.MAXIMUM; // 设置最大质量
app.jpegExportPreferences.exportResolution = 1400; // 设置分辨率为1400
app.jpegExportPreferences.jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING; // 使用基线编码
app.jpegExportPreferences.antiAlias = true; // 使用抗锯齿
app.jpegExportPreferences.jpegColorSpace = JpegColorSpaceEnum.GRAY; // 使用枚举设置颜色空间
app.jpegExportPreferences.useDocumentBleeds = false; //不使用文档出血
app.jpegExportPreferences.simulateOverprint = false; // 不模拟叠印
app.jpegExportPreferences.exportingSpread = false;　// 导出范围为页面
app.jpegExportPreferences.jpegExportRange　= ExportRangeOrAllPages.EXPORT_RANGE;    // 导出范围为单页面 

// 获取开始时间
var startTime = new Date().getTime();
var maxDuration = 5 * 60 * 1000; // 5分钟

// 导出每一页为单独的JPEG文件
for (var i = 0; i < doc.pages.length; i++) {
    if (userCancelled) {
        alert("操作已取消。");
        exit();
    }

    var currentTime = new Date().getTime();
    if (currentTime - startTime > maxDuration) {
        alert("导出时间超过2分钟，操作已停止。");
        break;
    }

    var page = doc.pages[i];
    var filePath = new File(folder.fsName + "/" + docName + "_" + ("00" + (i + 1)).slice(-3) + ".jpg");
    app.jpegExportPreferences.pageString = page.name;
    doc.exportFile(ExportFormat.JPG, filePath, false);
    
    // 更新进度条
    progressWin.progressBar.value = i + 1;
    progressWin.progressText.text = "正在导出页面 " + (i + 1) + " / " + doc.pages.length;
}

// 关闭进度条窗口
progressWin.close();

alert("导出完成！");