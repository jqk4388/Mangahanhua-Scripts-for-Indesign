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

// 导出每一页为单独的JPEG文件
if (exportMode == "current") {
    // 只导出当前页面
    var currentPage = app.activeWindow.activePage;
    if (checkAndUpdateLinksForPage(currentPage)) {
        var filePath = new File(folder.fsName + "/" + docName + "_" + currentPage.appliedSection.name + ("00" + (currentPage.name)).slice(-3) + ".jpg");
        app.jpegExportPreferences.pageString = currentPage.appliedSection.name + currentPage.name;
        doc.exportFile(ExportFormat.JPG, filePath, false);
        progressWin.progressBar.value = 1;
        progressWin.progressText.text = "已导出当前页面";
    }
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
        if (!checkAndUpdateLinksForPage(page)) {
            progressWin.progressBar.value = i + 1;
            progressWin.progressText.text = "跳过页面 " + (i + 1) + " / " + doc.pages.length;
            continue;
        }
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
        if (!checkAndUpdateLinksForPage(page)) {
            progressWin.progressBar.value = j + 1;
            progressWin.progressText.text = "跳过页面 " + page.name;
            continue;
        }
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
