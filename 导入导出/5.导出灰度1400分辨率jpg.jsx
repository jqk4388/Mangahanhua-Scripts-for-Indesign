// 获取当前活动文档
var doc = app.activeDocument;

// 获取文档名称并去掉扩展名和日语假名
var docName = doc.name.replace(/\.[^\.]+$/, '').replace(/[\u3040-\u30ff]+/g, '');

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

// 新增：导出选项窗口
var exportMode = "all"; // all, current, range
var pageRangeInput = "";
var w = new Window("dialog", "导出选项");
var rbCurrent = w.add("radiobutton", undefined, "导出当前页面");
var rbAll = w.add("radiobutton", undefined, "导出全部页面");
var rbRange = w.add("radiobutton", undefined, "导出指定页码范围");
rbCurrent.value = true;
var rangeGroup = w.add("group");
rangeGroup.add("statictext", undefined, "页码范围（如1-2或者001-002）：");
var rangeInput = rangeGroup.add("edittext", undefined, "");
rangeInput.characters = 10;
rangeInput.enabled = false;

rbCurrent.onClick = function() {
    rangeInput.enabled = false;
};
rbAll.onClick = function() {
    rangeInput.enabled = false;
};
rbRange.onClick = function() {
    rangeInput.enabled = true;
};

w.add("button", undefined, "确定", {name: "ok"});
if (w.show() != 1) exit();
if (rbCurrent.value) exportMode = "current";
else if (rbAll.value) exportMode = "all";
else if (rbRange.value) {
    exportMode = "range";
    pageRangeInput = rangeInput.text;
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
var maxDuration = 15 * 60 * 1000; // 15分钟

// 导出每一页为单独的JPEG文件
if (exportMode == "current") {
    // 只导出当前页面
    var currentPage = app.activeWindow.activePage;
    var filePath = new File(folder.fsName + "/" + docName + "_" + currentPage.appliedSection.name + ("00" + (currentPage.name)).slice(-3) + ".jpg");
    app.jpegExportPreferences.pageString = currentPage.appliedSection.name + currentPage.name;
    doc.exportFile(ExportFormat.JPG, filePath, false);
    progressWin.progressBar.value = 1;
    progressWin.progressText.text = "已导出当前页面";
} else if (exportMode == "all") {
    // 导出全部页面
    for (var i = 0; i < doc.pages.length; i++) {
        if (userCancelled) {
            alert("操作已取消。");
            exit();
        }
        var currentTime = new Date().getTime();
        if (currentTime - startTime > maxDuration) {
            alert("导出时间超过15分钟，操作已停止。");
            break;
        }
        var page = doc.pages[i];
        var filePath = new File(folder.fsName + "/" + docName + "_" + page.appliedSection.name + ("00" + (page.name)).slice(-3) + ".jpg");
        app.jpegExportPreferences.pageString = page.appliedSection.name + page.name;
        doc.exportFile(ExportFormat.JPG, filePath, false);
        progressWin.progressBar.value = i + 1;
        progressWin.progressText.text = "正在导出页面 " + (i + 1) + " / " + doc.pages.length;
    }
} else if (exportMode == "range") {
    // 导出指定页码范围
    var rangeParts = pageRangeInput.split("-");
    var startCode = rangeParts[0];
    var endCode = rangeParts[1];
    var exportPages = [];
    for (var i = 0; i < doc.pages.length; i++) {
        var pageName = doc.pages[i].name;
        var pageNumber = getPageNumber(pageName);
        if (pageNumber >= Number(startCode) && pageNumber <= Number(endCode)) {
            exportPages.push(doc.pages[i]);
        }
    }
    if (exportPages.length == 0) {
        alert("指定范围内没有找到页面。");
        exit();
    }
    for (var j = 0; j < exportPages.length; j++) {
        if (userCancelled) {
            alert("操作已取消。");
            exit();
        }
        var currentTime = new Date().getTime();
        if (currentTime - startTime > maxDuration) {
            alert("导出时间超过15分钟，操作已停止。");
            break;
        }
        var page = exportPages[j];
        var filePath = new File(folder.fsName + "/" + docName + "_" + page.appliedSection.name + ("00" + (page.name)).slice(-3) + ".jpg");
        app.jpegExportPreferences.pageString = page.appliedSection.name + page.name;
        doc.exportFile(ExportFormat.JPG, filePath, false);
        progressWin.progressBar.value = j + 1;
        progressWin.progressText.text = "正在导出页面 " + page.name;
    }
}

// 关闭进度条窗口
progressWin.close();

alert("导出完成！");


function getPageNumber(pageName) {
    // 先尝试阿拉伯数字
    if (!isNaN(Number(pageName))) return Number(pageName);
    // 尝试汉字数字
    var cn = chineseNumToInt(pageName);
    if (!isNaN(cn)) return cn;
    // 可扩展罗马数字等
    return null; // 无法识别
}

// 汉字数字转阿拉伯数字（仅支持一到十，可扩展）
function chineseNumToInt(str) {
    var map = {
        "零":0, "一":1, "二":2, "三":3, "四":4, "五":5, "六":6, "七":7, "八":8, "九":9, "十":10
    };
    if (str.length === 1) {
        return map[str] !== undefined ? map[str] : NaN;
    } else if (str.length === 2 && str[0] === "十") {
        // 十一、十二...
        return 10 + map[str[1]];
    } else if (str.length === 2 && str[1] === "十") {
        // 二十、三十...
        return map[str[0]] * 10;
    } else if (str.length === 3 && str[1] === "十") {
        // 二十一、三十二...
        return map[str[0]] * 10 + map[str[2]];
    }
    return NaN;
}
