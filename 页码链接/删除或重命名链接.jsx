// 删除链接或重命名链接图像脚本
// 该脚本提供一个用户界面来选择操作范围和类型，并对文档中的链接进行相应处理

#targetengine "session"

// 主函数
function main() {
    try {
        // 检查是否有打开的文档
        if (app.documents.length === 0) {
            alert("请先打开一个文档");
            return;
        }
        
        var myDocument = app.activeDocument;
        
        // 创建对话框
        var dialog = new Window("dialog", "链接处理工具");
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        
        // 范围选择面板
        var rangePanel = dialog.add("panel", undefined, "选择操作范围");
        rangePanel.orientation = "row";
        rangePanel.alignChildren = "left";
        var currentPageRadio = rangePanel.add("radiobutton", undefined, "当前页面");
        var entireDocumentRadio = rangePanel.add("radiobutton", undefined, "整个文档");
        entireDocumentRadio.value = true; // 默认选择整个文档
        
        // 操作类型面板
        var operationPanel = dialog.add("panel", undefined, "选择操作类型");
        operationPanel.orientation = "row";
        operationPanel.alignChildren = "left";
        var renameRadio = operationPanel.add("radiobutton", undefined, "重命名");
        var deleteRadio = operationPanel.add("radiobutton", undefined, "删除");
        renameRadio.value = true; // 默认选择重命名
        
        // 重命名参数面板（默认隐藏）
        var renamePanel = dialog.add("panel", undefined, "重命名参数");
        renamePanel.orientation = "column";
        renamePanel.alignChildren = "fill";
        renamePanel.visible = true;
        
        var prefixGroup = renamePanel.add("group");
        prefixGroup.orientation = "row";
        prefixGroup.add("statictext", undefined, "前缀:");
        var prefixInput = prefixGroup.add("edittext", undefined, "");
        prefixInput.characters = 20;
        
        var suffixGroup = renamePanel.add("group");
        suffixGroup.orientation = "row";
        suffixGroup.add("statictext", undefined, "后缀:");
        var suffixInput = suffixGroup.add("edittext", undefined, "");
        suffixInput.characters = 20;
        
        // 添加页码选项
        var includePageNumberCheckbox = renamePanel.add("checkbox", undefined, "在文件名最前面添加页码");
        includePageNumberCheckbox.value = false; // 默认不勾选
        
        // 删除参数面板（默认隐藏）
        var deletePanel = dialog.add("panel", undefined, "删除参数");
        deletePanel.orientation = "column";
        deletePanel.alignChildren = "left";
        deletePanel.visible = false;
        var keepFrameCheckbox = deletePanel.add("checkbox", undefined, "保留框架");
        keepFrameCheckbox.value = true; // 默认勾选
        
        // 按钮组
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "center";
        var okButton = buttonGroup.add("button", undefined, "确定");
        var cancelButton = buttonGroup.add("button", undefined, "取消");
        
        // 监听操作类型变化，显示相应面板
        renameRadio.onClick = function() {
            renamePanel.visible = true;
            deletePanel.visible = false;
        };
        
        deleteRadio.onClick = function() {
            renamePanel.visible = false;
            deletePanel.visible = true;
        };
        // 确定操作范围
        var targetPages = [];
        if (currentPageRadio.value) {
            // 当前页面
            targetPages.push(app.activeWindow.activePage);
        } else {
            // 整个文档
            targetPages = myDocument.pages;
        }        
        // 确定按钮事件处理
        okButton.onClick = function() {
            dialog.close(1);
            // 执行相应操作
        if (renameRadio.value) {
            // 执行重命名操作
            renameLinks(myDocument, targetPages, prefixInput.text, suffixInput.text, includePageNumberCheckbox.value);
        } else {
            // 执行删除操作
            deleteLinks(myDocument, targetPages, keepFrameCheckbox.value);
        }
        };
        
        // 取消按钮事件处理
        cancelButton.onClick = function() {
            dialog.close(0);
        };
        
        // 显示对话框
        var result = dialog.show();
        
        // 如果用户点击取消，则退出
        if (result === 0) {
            return;
        }

    } catch (e) {
        alert("发生错误: " + e.message);
    }
}

// 重命名链接函数
function renameLinks(document, pages, prefix, suffix, includePageNumber) {
    try {
        var processedCount = 0;
        var links = document.links;
        
        // 创建进度条窗口
        var progressWindow = new Window("palette", "处理进度");
        progressWindow.orientation = "column";
        progressWindow.alignChildren = "fill";
        var progressBar = progressWindow.add("progressbar", undefined, 0, links.length);
        var progressText = progressWindow.add("statictext", undefined, "正在处理...");
        progressWindow.show();
        
        // 遍历所有链接
        for (var i = 0; i < links.length; i++) {
            var link = links[i];
            
            // 更新进度条
            progressBar.value = i + 1;
            progressText.text = "正在处理 " + (i + 1) + " / " + links.length;
            progressWindow.update();
            
            // 检查链接是否在目标页面中
            if (isLinkOnTargetPages(link, pages)) {
                try {
                    // 获取链接文件
                    var file = new File(link.filePath);
                    if (!file.exists) {
                        continue;
                    }
                    
                    // 获取原始文件名（不含扩展名）
                    var originalName = getFileNameWithoutExtension(link.name);
                    var extension = link.name.substring(link.name.lastIndexOf(".") + 1);
                    
                    // 构造新名称
                    var newName = prefix + originalName + suffix;
                    
                    // 获取链接的父对象所在的页面
                    var parentImage = link.parent;
                    var parentFrame = parentImage.parent;
                    var page = parentFrame.parentPage;
                    var pageName = page.name;
                    
                    // 添加页面名称到新文件名中（根据选项决定是否包含页码）
                    var nameChanged = "";
                    if (includePageNumber) {
                        nameChanged = padZero(pageName) + "-" + newName;
                    } else {
                        nameChanged = newName;
                    }
                    
                    // 检查文件是否已存在，如果存在则添加版本号
                    var nameNew = nameChanged;
                    var fileVersion = 0;
                    var fileNew = new File(file.path + "/" + nameNew + "." + extension);
                    while (fileNew.exists) {
                        fileVersion++;
                        nameNew = nameChanged + "~" + fileVersion;
                        fileNew = new File(file.path + "/" + nameNew + "." + extension);
                    }
                    
                    // 重命名文件
                    if (file.rename(nameNew + "." + extension)) {
                        // 重链接所有同名链接
                        for (var j = 0; j < document.links.length; j++) {
                            if (document.links[j].name == link.name) {
                                var relinkFile = new File(file.path + "/" + nameNew + "." + extension);
                                document.links[j].relink(relinkFile);
                                document.links[j].update();
                                processedCount++;
                            }
                        }
                    }
                } catch (renameError) {
                    // 忽略单个链接的错误，继续处理其他链接
                }
            }
        }
        
        // 关闭进度条
        progressWindow.close();
        
        // 显示结果
        alert("重命名完成，共处理了 " + processedCount + " 个链接。");
        
    } catch (e) {
        alert("重命名过程中发生错误: " + e.message);
    }
}

// 删除链接函数
function deleteLinks(document, pages, keepFrame) {
    try {
        var processedCount = 0;
        var links = document.links;
        
        // 创建进度条窗口
        var progressWindow = new Window("palette", "处理进度");
        progressWindow.orientation = "column";
        progressWindow.alignChildren = "fill";
        var progressBar = progressWindow.add("progressbar", undefined, 0, links.length);
        var progressText = progressWindow.add("statictext", undefined, "正在处理...");
        progressWindow.show();
        
        // 收集需要处理的链接
        var linksToProcess = [];
        for (var i = 0; i < links.length; i++) {
            var link = links[i];
            if (isLinkOnTargetPages(link, pages)) {
                linksToProcess.push(link);
            }
        }
        
        // 反向遍历以避免索引问题
        for (var j = linksToProcess.length - 1; j >= 0; j--) {
            var linkToProcess = linksToProcess[j];
            
            // 更新进度条
            progressBar.value = linksToProcess.length - j;
            progressText.text = "正在处理 " + (linksToProcess.length - j) + " / " + linksToProcess.length;
            progressWindow.update();
            
            try {
                // 获取链接的父对象（Image对象）
                var parentImage = linkToProcess.parent;
                
                // 获取Image对象的父对象（Rectangle框架）
                var parentFrame = parentImage.parent;
                
                if (keepFrame) {
                    // 仅删除链接内容，保留框架
                    parentImage.remove();
                } else {
                    // 删除整个框架
                    parentFrame.remove();
                }
                
                processedCount++;
            } catch (deleteError) {
                // 忽略单个链接的错误，继续处理其他链接
            }
        }
        
        // 关闭进度条
        progressWindow.close();
        
        // 显示结果

        if (processedCount === 0) {
            alert("未找到符合条件的链接。可能是因为链接被锁定无法删除。");
        }else {
            alert("删除完成，共删除了 " + processedCount + " 个链接。");
        }
        
    } catch (e) {
        alert("删除过程中发生错误: " + e.message);
    }
}

// 检查链接是否在目标页面上
function isLinkOnTargetPages(link, pages) {
    try {
        // 获取链接的父对象(Image)
        var parentImage = link.parent;
        
        // 获取Image的父对象(Rectangle框架)
        var parentFrame = parentImage.parent;
        
        // 获取框架所在的页面
        var page = parentFrame.parentPage;
        
        if (page != null) {
            // 检查此页面是否在目标页面列表中
            for (var j = 0; j < pages.length; j++) {
                if (pages[j] == page) {
                    return true;
                }
            }
        }
        
        return false;
    } catch (e) {
        return false;
    }
}

// 获取不带扩展名的文件名
function getFileNameWithoutExtension(fileName) {
    try {
        var lastDotIndex = -1;
        
        // 手动查找最后一个点的位置
        for (var i = fileName.length - 1; i >= 0; i--) {
            if (fileName[i] == ".") {
                lastDotIndex = i;
                break;
            }
        }
        
        // 如果找到点，则返回点之前的部分，否则返回完整文件名
        if (lastDotIndex > 0) {
            return fileName.substring(0, lastDotIndex);
        } else {
            return fileName;
        }
    } catch (e) {
        return fileName;
    }
}

// 补零函数
function padZero(v) {
    if (v < 10) {
        return ("0" + v).slice(-2);
    }
    return v;
}

// 执行主函数
main();