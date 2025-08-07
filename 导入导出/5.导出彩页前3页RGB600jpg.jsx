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

// 新增：导出选项窗口
var exportMode = "first"; // first, range
var pageCount = 3;
var pageRangeInput = "";
var w = new Window("dialog", "导出彩页选项");
var rbFirst = w.add("radiobutton", undefined, "导出前X页");
var rbRange = w.add("radiobutton", undefined, "导出页面范围");
rbFirst.value = true;
var groupFirst = w.add("group");
groupFirst.add("statictext", undefined, "X=");
var inputFirst = groupFirst.add("edittext", undefined, "3");
inputFirst.characters = 4;
inputFirst.enabled = true;
var groupRange = w.add("group");
groupRange.add("statictext", undefined, "页面范围【非页码】（如1-3）：");
var inputRange = groupRange.add("edittext", undefined, "");
inputRange.characters = 8;
inputRange.enabled = false;
rbFirst.onClick = function() {
    inputFirst.enabled = true;
    inputRange.enabled = false;
};
rbRange.onClick = function() {
    inputFirst.enabled = false;
    inputRange.enabled = true;
};
w.add("button", undefined, "确定", {name: "ok"});
if (w.show() != 1) exit();
if (rbFirst.value) {
    exportMode = "first";
    var num = parseInt(inputFirst.text, 10);
    if (isNaN(num) || num < 1 || num > doc.pages.length) {
        alert("请输入1到" + doc.pages.length + "之间的正整数！");
        exit();
    }
    pageCount = num;
} else {
    exportMode = "range";
    pageRangeInput = inputRange.text;
    if (!/^\d+-\d+$/.test(pageRangeInput)) {
        alert("请输入正确的范围格式，如1-3");
        exit();
    }
    var parts = pageRangeInput.split("-");
    var startIdx = parseInt(parts[0], 10) - 1;
    var endIdx = parseInt(parts[1], 10) - 1;
    if (isNaN(startIdx) || isNaN(endIdx) || startIdx < 0 || endIdx >= doc.pages.length || startIdx > endIdx) {
        alert("请输入有效的页面范围！");
        exit();
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

// 设置导出选项
app.jpegExportPreferences.jpegQuality = JPEGOptionsQuality.MAXIMUM; // 设置最大质量
app.jpegExportPreferences.exportResolution = 600; // 设置分辨率为600
app.jpegExportPreferences.jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING; // 使用基线编码
app.jpegExportPreferences.antiAlias = true; // 使用抗锯齿
app.jpegExportPreferences.jpegColorSpace = JpegColorSpaceEnum.RGB; // 使用RGB
app.jpegExportPreferences.useDocumentBleeds = false; //不使用文档出血
app.jpegExportPreferences.simulateOverprint = true; // 模拟叠印
app.jpegExportPreferences.exportingSpread = false;　// 导出范围为单页面非跨页
app.jpegExportPreferences.jpegExportRange　= ExportRangeOrAllPages.EXPORT_RANGE;    // 导出范围页面 
app.jpegExportPreferences.embedColorProfile = true; // 嵌入颜色配置文件

// 获取开始时间
var startTime = new Date().getTime();
var maxDuration = 5 * 60 * 1000; // 5分钟

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
// 导出彩页
if (exportMode == "first") {
    for (var i = 0; i < pageCount; i++) {
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
        if (!checkAndUpdateLinksForPage(page)) {
            progressWin.progressBar.value = i + 1;
            progressWin.progressText.text = "跳过页面 " + (i + 1) + " / " + doc.pages.length;
            continue;
        }
        var filePath = new File(folder.fsName + "/" + docName + "_" + page.appliedSection.name + ("00" + (page.name)).slice(-3) + ".jpg");
        app.jpegExportPreferences.pageString = page.appliedSection.name + page.name;
        doc.exportFile(ExportFormat.JPG, filePath, false);
        
        // 更新进度条
        progressWin.progressBar.value = i + 1;
        progressWin.progressText.text = "正在导出页面 " + (i + 1) + " / " + pageCount;
    }
} else if (exportMode == "range") {
    var parts = pageRangeInput.split("-");
    var startIdx = parseInt(parts[0], 10) - 1;
    var endIdx = parseInt(parts[1], 10) - 1;
    for (var i = startIdx; i <= endIdx; i++) {
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
        if (!checkAndUpdateLinksForPage(page)) {
            progressWin.progressBar.value = i + 1;
            progressWin.progressText.text = "跳过页面 " + (i + 1) + " / " + doc.pages.length;
            continue;
        }
        var filePath = new File(folder.fsName + "/" + docName + "_" + page.appliedSection.name + ("00" + (page.name)).slice(-3) + ".jpg");
        app.jpegExportPreferences.pageString = page.appliedSection.name + page.name;
        doc.exportFile(ExportFormat.JPG, filePath, false);
        
        // 更新进度条
        progressWin.progressBar.value = i + 1;
        progressWin.progressText.text = "正在导出页面 " + (i + 1) + " / " + pageCount;
    }
}

// 关闭进度条窗口
progressWin.close();

alert("导出完成！");