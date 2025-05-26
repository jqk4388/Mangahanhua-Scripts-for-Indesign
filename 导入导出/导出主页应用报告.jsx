#target indesign

// 确保有文档打开
if (app.documents.length === 0) {
    alert("请先打开一个文档！");
} else {
    generateMasterPageReport();
}

function generateMasterPageReport() {
    var doc = app.activeDocument;
    var report = "文档：" + doc.name + "\n";
    report += "页面数：" + doc.pages.length + "\n\n";
    report += "页面主页使用情况报告：\n";
    report += "----------------------------------------\n";
    
    // 遍历所有页面
    for (var i = 0; i < doc.pages.length; i++) {
        var page = doc.pages[i];
        var masterName = page.appliedMaster.name;
        report += "页面 " + (i + 1) + " 使用的主页：" + masterName + "\n";
    }
    
    // 保存报告到桌面
    var desktop = Folder.desktop;
    var reportFile = new File(desktop + "/主页使用情况报告_" + getDateString() + ".txt");
    
    try {
        reportFile.open("w");
        reportFile.write(report);
        reportFile.close();
        alert("报告已保存到桌面：\n" + reportFile.fsName);
    } catch (e) {
        alert("保存文件时出错：" + e);
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