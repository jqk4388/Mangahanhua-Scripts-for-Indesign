// 获取当前激活的文档
var activeDoc = app.activeDocument;

// 打开文件选择对话框选择 INDD 文件
var file = File.openDialog("选择一个 INDD 文件", "*.indd");

// 检查用户是否选择了文件
if (file != null) {
    activeDoc.importStyles(ImportFormat.textStylesFormat, file, GlobalClashResolutionStrategy.loadAllWithOverwrite);
    activeDoc.importStyles(ImportFormat.OBJECT_STYLES_FORMAT, file, GlobalClashResolutionStrategy.loadAllWithOverwrite);
    alert("字符样式、段落样式、对象样式已成功导入！");
} else {
    alert("未选择任何文件。");
}
