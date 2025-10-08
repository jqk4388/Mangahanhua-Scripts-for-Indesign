#include "../Library/KTUlib.jsx"
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
        var selectedLinksRadio = rangePanel.add("radiobutton", undefined, "当前选中的链接");
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
        // 确定按钮事件处理：只关闭对话框，实际执行在 dialog.show() 返回后进行
        okButton.onClick = function() {
            dialog.close(1);
        };
        
        // 取消按钮事件处理
        cancelButton.onClick = function() {
            dialog.close(0);
        };
        
        // 显示对话框
        var result = dialog.show();
        
        // 如果用户点击取消，则退出；如果确认 (result === 1)，在对话框消失后执行具体操作
        if (result === 0) {
            return;
        }

        // 构建选项对象并在对话框关闭后执行操作
        var options = {
            range: (selectedLinksRadio.value ? 'selectedLinks' : (currentPageRadio.value ? 'currentPage' : 'entireDocument')),
            operation: (renameRadio.value ? 'rename' : 'delete'),
            prefix: prefixInput.text,
            suffix: suffixInput.text,
            includePageNumber: includePageNumberCheckbox.value,
            keepFrame: keepFrameCheckbox.value
        };

        // 先预览要执行的操作，让用户确认
        var preview = previewActions(myDocument, options);

        if (!preview || preview.count === 0) {
            alert('未找到要处理的链接。操作已取消。');
            return;
        }

        // 显示预览对话框
        var previewDlg = new Window('dialog', '操作预览');
        previewDlg.orientation = 'column';
        previewDlg.alignChildren = 'fill';

        var info = preview.operationDesc + '\n\n共计: ' + preview.count + ' 个链接将被处理。\n\n样例（最多显示 10 项）:';
        previewDlg.add('statictext', undefined, info);

        var sampleText = preview.sample.join('\n');
        var previewBox = previewDlg.add('edittext', undefined, sampleText, {multiline: true});
        previewBox.minimumSize = [400, 200];
        previewBox.readonly = true;

        var btnGroup = previewDlg.add('group');
        btnGroup.alignment = 'center';
        var proceedBtn = btnGroup.add('button', undefined, '继续执行');
        var cancelBtn = btnGroup.add('button', undefined, '取消');

        proceedBtn.onClick = function() { previewDlg.close(1); };
        cancelBtn.onClick = function() { previewDlg.close(0); };

        var previewResult = previewDlg.show();
        if (previewResult === 1) {
            // 把 preview 中计算到的 matches 传给执行函数，避免二次扫描
            options.matches = preview.matches;
            performActions(myDocument, options);
        } else {
            // 用户取消
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

// 重命名选中的链接函数
function renameSelectedLinks(document, prefix, suffix, includePageNumber) {
    try {
        var selectedLinks = getSelectedLinksFromClipboard(document);
        var processedCount = 0;
        
        // 创建进度条窗口
        var progressWindow = new Window("palette", "处理进度");
        progressWindow.orientation = "column";
        progressWindow.alignChildren = "fill";
        var progressBar = progressWindow.add("progressbar", undefined, 0, selectedLinks.length);
        var progressText = progressWindow.add("statictext", undefined, "正在处理...");
        progressWindow.show();
        
        // 遍历选中的链接
        for (var i = 0; i < selectedLinks.length; i++) {
            var link = selectedLinks[i];
            
            // 更新进度条
            progressBar.value = i + 1;
            progressText.text = "正在处理 " + (i + 1) + " / " + selectedLinks.length;
            progressWindow.update();
            
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

// 删除选中的链接函数
function deleteSelectedLinks(document, keepFrame) {
    try { 
        var selectedLinks = [];
        selectedLinks = getSelectedLinksFromClipboard(document)
        // 筛选出选中的链接
        for (var i = 0; i < app.selection.length; i++) {
            var item = app.selection[i];
            if (item.hasOwnProperty("images") && item.images.length > 0) {
                var image = item.images[0];
                if (image.hasOwnProperty("itemLink")) {
                    selectedLinks.push(image.itemLink);
                }
            }
        }
        
        if (selectedLinks.length === 0) {
            alert("未找到选中的链接图像");
            return;
        }
        
        var processedCount = 0;
        
        // 创建进度条窗口
        var progressWindow = new Window("palette", "处理进度");
        progressWindow.orientation = "column";
        progressWindow.alignChildren = "fill";
        var progressBar = progressWindow.add("progressbar", undefined, 0, selectedLinks.length);
        var progressText = progressWindow.add("statictext", undefined, "正在处理...");
        progressWindow.show();
        
        // 反向遍历以避免索引问题
        for (var j = selectedLinks.length - 1; j >= 0; j--) {
            var linkToProcess = selectedLinks[j];
            
            // 更新进度条
            progressBar.value = selectedLinks.length - j;
            progressText.text = "正在处理 " + (selectedLinks.length - j) + " / " + selectedLinks.length;
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
        } else {
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

// 在对话框关闭后执行的操作函数
function performActions(document, options) {
    try {
        // 计算 targetPages（仅在需要时）
        var targetPages = [];
        if (options.range === 'currentPage') {
            try {
                targetPages.push(app.activeWindow.activePage);
            } catch (e) {
                targetPages = document.pages;
            }
        } else if (options.range === 'entireDocument') {
            targetPages = document.pages;
        } else {
            targetPages = [];
        }
        // 如果 preview 传入了 matches，则优先使用 matches 来执行，避免二次遍历
        if (options.matches && options.matches.length >= 0) {
            if (options.operation === 'rename') {
                renameLinksFromList(document, options.matches, options.prefix, options.suffix, options.includePageNumber);
            } else if (options.operation === 'delete') {
                KTUDoScriptAsUndoable(function() { deleteLinksFromList(document, options.matches, options.keepFrame); }, (options.range === 'selectedLinks' ? "删除选中的链接" : "删除页面链接"));
            }
            return;
        }

        if (options.operation === 'rename') {
            if (options.range === 'selectedLinks') {
                renameSelectedLinks(document, options.prefix, options.suffix, options.includePageNumber);
            } else {
                renameLinks(document, targetPages, options.prefix, options.suffix, options.includePageNumber);
            }
        } else if (options.operation === 'delete') {
            if (options.range === 'selectedLinks') {
                KTUDoScriptAsUndoable(function() { deleteSelectedLinks(document, options.keepFrame); }, "删除选中的页面链接");
            } else {
                KTUDoScriptAsUndoable(function() { deleteLinks(document, targetPages, options.keepFrame); }, "删除页面链接");
            }
        }
    } catch (e) {
        alert("执行操作时发生错误: " + e.message);
    }
}

// 预览要执行的操作，返回 {count, sample[], operationDesc}
function previewActions(document, options) {
    try {
        var links = document.links;
        var matches = [];

        if (options.range === 'selectedLinks') {
            // 收集来自剪贴板或当前选区的选中链接
            matches = getSelectedLinksFromClipboard(document) || [];
        }

        // 准备样例文本
        var sample = [];
        for (var m = 0; m < Math.min(10, matches.length); m++) {
            try {
                var ln = matches[m].name;
                var pageName = '';
                try { pageName = matches[m].parent.parent.parentPage.name; } catch (e) { pageName = ''; }
                sample.push((m+1) + '. ' + ln + (pageName ? ('  (page: ' + pageName + ')') : ''));
            } catch (e) {
                sample.push((m+1) + '. (无法读取名称)');
            }
        }

        var opDesc = (options.operation === 'rename') ? ('重命名' + (options.prefix || '' ) + '...') : '删除链接';

        return { count: matches.length, sample: sample, operationDesc: opDesc, matches: matches };
    } catch (e) {
        alert("预览操作时发生错误: " + e.message);
        return { count: 0, sample: [], operationDesc: '', matches: [] };
    }
}

// 根据 preview.matches 直接重命名
function renameLinksFromList(document, selectedLinks, prefix, suffix, includePageNumber) {
    try {
        var processedCount = 0;

        var progressWindow = new Window("palette", "处理进度");
        progressWindow.orientation = "column";
        progressWindow.alignChildren = "fill";
        var progressBar = progressWindow.add("progressbar", undefined, 0, selectedLinks.length);
        var progressText = progressWindow.add("statictext", undefined, "正在处理...");
        progressWindow.show();

        for (var i = 0; i < selectedLinks.length; i++) {
            var link = selectedLinks[i];
            progressBar.value = i + 1;
            progressText.text = "正在处理 " + (i + 1) + " / " + selectedLinks.length;
            progressWindow.update();

            try {
                var file = new File(link.filePath);
                if (!file.exists) continue;

                var originalName = getFileNameWithoutExtension(link.name);
                var extension = link.name.substring(link.name.lastIndexOf(".") + 1);
                var newName = prefix + originalName + suffix;

                var parentImage = link.parent;
                var parentFrame = parentImage.parent;
                var page = parentFrame.parentPage;
                var pageName = page ? page.name : '';

                var nameChanged = includePageNumber && pageName ? padZero(pageName) + "-" + newName : newName;

                var nameNew = nameChanged;
                var fileVersion = 0;
                var fileNew = new File(file.path + "/" + nameNew + "." + extension);
                while (fileNew.exists) {
                    fileVersion++;
                    nameNew = nameChanged + "~" + fileVersion;
                    fileNew = new File(file.path + "/" + nameNew + "." + extension);
                }

                if (file.rename(nameNew + "." + extension)) {
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
                // 忽略
            }
        }

        progressWindow.close();
        alert("重命名完成，共处理了 " + processedCount + " 个链接。");
    } catch (e) {
        alert("重命名过程中发生错误: " + e.message);
    }
}

// 根据 preview.matches 直接删除
function deleteLinksFromList(document, selectedLinks, keepFrame) {
    try {
        var processedCount = 0;

        var progressWindow = new Window("palette", "处理进度");
        progressWindow.orientation = "column";
        progressWindow.alignChildren = "fill";
        var progressBar = progressWindow.add("progressbar", undefined, 0, selectedLinks.length);
        var progressText = progressWindow.add("statictext", undefined, "正在处理...");
        progressWindow.show();

        for (var j = selectedLinks.length - 1; j >= 0; j--) {
            var linkToProcess = selectedLinks[j];
            progressBar.value = selectedLinks.length - j;
            progressText.text = "正在处理 " + (selectedLinks.length - j) + " / " + selectedLinks.length;
            progressWindow.update();

            try {
                var parentImage = linkToProcess.parent;
                var parentFrame = parentImage.parent;
                if (keepFrame) {
                    parentImage.remove();
                } else {
                    parentFrame.remove();
                }
                processedCount++;
            } catch (deleteError) {
                // 忽略
            }
        }

        progressWindow.close();
        if (processedCount === 0) {
            alert("未找到符合条件的链接。可能是因为链接被锁定无法删除。");
        } else {
            alert("删除完成，共删除了 " + processedCount + " 个链接。");
        }
    } catch (e) {
        alert("删除过程中发生错误: " + e.message);
    }
}

// 执行主函数
main();