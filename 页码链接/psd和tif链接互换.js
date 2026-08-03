// ----------- 重新链接图像 ----------- // 
// 将当前活动页面上的基础图像重新链接到保存在同一目录中的转换后图像。
// 目前仅对页面上的第一个链接图像有效。
//
// 更新日期：2026年8月3日
//
// 示例： 
//  在 INDD 中链接的文件是 PSD。在 Photoshop 中，您将图像转换并合并为新的 .tif
//  您希望将 INDD 中的图像重新链接为新的 .tif
// 
// 注意： 
//  新图像需要与原始图像位于同一文件夹中。 

// 根据需要修改以下文件类型！
var fileTypes = [
    { oldType: '.psd', newType: '.tif' },
    { oldType: '.tif', newType: '.psd' },
]

function main() {
    if (!isError()) {
        var srcImage = app.activeDocument.layoutWindows[0].activePage.allGraphics[0];
        relink(srcImage);
    }
}

function isError() {
    if (app.activeDocument.layoutWindows[0].activePage.allGraphics.length < 1) {
        alert('请确保活动页面上有图像');
        return true;
    }
    return false;
}

function relink(srcImage) {
    var isFound = false;
    for (var i = 0; i < fileTypes.length && !isFound; i++) {
        var ref = fileTypes[i];
        var oldPath = srcImage.itemLink.filePath;
        var escapedExt = ref.oldType.replace(/\./g, '\\.');
        var extRegex = new RegExp(escapedExt + '$', 'i');
        var newPath = oldPath.replace(extRegex, ref.newType);
        var newImage = new File(newPath);

        if (oldPath.toLowerCase() !== newPath.toLowerCase() && newImage.exists) {
            srcImage.itemLink.relink(newImage);
            isFound = true;
        }
    }
}

main()