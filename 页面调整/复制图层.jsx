#include "../Library/KTUlib.jsx"

// 跨文档图层复制脚本
// 功能：遍历源文档（A）指定图层的元素，按页面名称匹配目标文档（B）的页面，并在目标文档中原位粘贴
// 若目标文档无对应页面则新建，图层同名时自动添加数字后缀

// 检测已打开文档
function getOpenDocuments() {
    try {
        var docs = [];
        for (var i = 0; i < app.documents.length; i++) {
            docs.push(app.documents[i]);
        }
        return docs; // 返回文档对象列表
    } catch (e) {
        throw new Error("文档检测失败: " + e.message);
    }
}

// 从文档中获取所有图层名称
function getLayerNames(doc) {
    try {
        var layerNames = [];
        for (var i = 0; i < doc.layers.length; i++) {
            layerNames.push(doc.layers[i].name);
        }
        return layerNames;
    } catch (e) {
        throw new Error("获取图层名称失败: " + e.message);
    }
}

// 构建用户界面（UI）
function showCopyDialog(sourceDocs) {
    if (sourceDocs.length < 2) {
        alert("需要至少打开两个文档才能进行复制操作！");
        return null;
    }

    var dialog = new Window("dialog", "复制图层");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    // 上半部分：选择源文档和目标文档
    var docGroup = dialog.add("panel", undefined, "文档选择");
    docGroup.orientation = "row";
    docGroup.alignChildren = "top";

    var sourceGroup = docGroup.add("group");
    sourceGroup.orientation = "column";
    sourceGroup.add("statictext", undefined, "源文档:");
    var sourceDropdown = sourceGroup.add("dropdownlist", undefined, []);
    for (var i = 0; i < sourceDocs.length; i++) {
        sourceDropdown.add("item", sourceDocs[i].name);
    }
    if (sourceDropdown.items.length > 0) {
        sourceDropdown.items[0].selected = true;
    }

    var targetGroup = docGroup.add("group");
    targetGroup.orientation = "column";
    targetGroup.add("statictext", undefined, "目标文档:");
    var targetDropdown = targetGroup.add("dropdownlist", undefined, []);
    for (var i = 0; i < sourceDocs.length; i++) {
        targetDropdown.add("item", sourceDocs[i].name);
    }
    if (sourceDropdown.items.length > 1) {
        targetDropdown.items[1].selected = true;
    }

    // 中间部分：选择图层
    var layerGroup = dialog.add("panel", undefined, "图层选择");
    layerGroup.orientation = "column";
    layerGroup.alignChildren = "left";
    
    var layerStatic = layerGroup.add("statictext", undefined, "选择图层:");
    var layerDropdown = layerGroup.add("dropdownlist", undefined, []);

    // 更新图层列表的函数
    function updateLayerList(docIndex) {
        if (docIndex >= 0 && docIndex < sourceDocs.length) {
            var selectedDoc = sourceDocs[docIndex];
            var layerNames = getLayerNames(selectedDoc);
            layerDropdown.removeAll();
            for (var i = 0; i < layerNames.length; i++) {
                layerDropdown.add("item", layerNames[i]);
            }
            if (layerDropdown.items.length > 0) {
                layerDropdown.items[0].selected = true;
            }
        }
    }

    // 当选择源文档时更新图层列表
    sourceDropdown.onChange = function() {
        updateLayerList(sourceDropdown.selection.index);
    };

    // 下半部分：选择元素类型和页码范围
    var optionsGroup = dialog.add("panel", undefined, "复制选项");
    optionsGroup.orientation = "column";
    optionsGroup.alignChildren = "left";

    // 元素类型复选框
    var elementTypesGroup = optionsGroup.add("group");
    elementTypesGroup.add("statictext", undefined, "元素类型:");
    var textFramesCheckbox = elementTypesGroup.add("checkbox", undefined, "文本框");
    textFramesCheckbox.value = true; // 默认勾选
    var graphicFramesCheckbox = elementTypesGroup.add("checkbox", undefined, "图像框");
    graphicFramesCheckbox.value = true; // 默认勾选
    var pathsCheckbox = elementTypesGroup.add("checkbox", undefined, "路径");
    pathsCheckbox.value = true; // 默认勾选

    // 页码范围输入
    var rangeGroup = optionsGroup.add("group");
    rangeGroup.add("statictext", undefined, "页码范围 (如: 1-3 或 all):");
    var pageRangeEdit = rangeGroup.add("edittext", undefined, "all");
    pageRangeEdit.characters = 20;

    // 按钮组
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignChildren = "center";
    var okButton = buttonGroup.add("button", undefined, "确定");
    var cancelButton = buttonGroup.add("button", undefined, "取消");
    
    // 保存对控件的引用到对话框对象，以便主函数访问
    dialog.sourceDropdown = sourceDropdown;
    dialog.targetDropdown = targetDropdown;
    dialog.layerDropdown = layerDropdown;
    dialog.textFramesCheckbox = textFramesCheckbox;
    dialog.graphicFramesCheckbox = graphicFramesCheckbox;
    dialog.pathsCheckbox = pathsCheckbox;
    dialog.pageRangeEdit = pageRangeEdit;

    // 默认选择第一个文档作为源，第二个作为目标
    if (sourceDocs.length >= 2) {
        sourceDropdown.items[0].selected = true;
        targetDropdown.items[1].selected = true;
        updateLayerList(0);
    }

    // 确定按钮事件
    okButton.onClick = function() {
        dialog.close(1); // 返回 1 表示确定
    };

    // 取消按钮事件
    cancelButton.onClick = function() {
        dialog.close(0); // 返回 0 表示取消
    };

    return dialog;
}

// 检查页面是否在指定范围内
function isPageInRange(pageName, pageRangeStr) {
    // 解析页码范围，例如 "1-3" 或 "all"
    if (pageRangeStr === "all" || pageRangeStr === "") {
        return true;
    }

    // 提取页面名称中的数字部分
    var pageNumbers = pageName.match(/\d+/g);
    if (!pageNumbers || pageNumbers.length === 0) {
        return false; // 如果页面名称不包含数字，则认为不在范围内
    }

    var pageNum = parseInt(pageNumbers[0]); // 取第一个数字作为页码

    // 解析范围字符串
    var ranges = pageRangeStr.split(",");
    for (var i = 0; i < ranges.length; i++) {
        var range = ranges[i].replace(/\s/g, ""); // 去除空格
        if (range.indexOf("-") !== -1) {
            // 处理范围，如 "1-3"
            var parts = range.split("-");
            var start = parseInt(parts[0]);
            var end = parseInt(parts[1]);
            if (pageNum >= start && pageNum <= end) {
                return true;
            }
        } else {
            // 处理单个数字，如 "1"
            if (parseInt(range) === pageNum) {
                return true;
            }
        }
    }

    return false;
}
// 确保页面数足够
function ensureEnoughPages(doc,requiredPageCount) {
    var currentPageCount = doc.pages.length;
    
    if (currentPageCount < requiredPageCount) {
        var pagesToAdd = requiredPageCount - currentPageCount;
        for (var i = 0; i < pagesToAdd; i++) {
            doc.pages.add();
        }
    }
}
// 按页面名称匹配/新建页面
function createPageByName(doc, pageName) {
    // 首先尝试查找同名页面
    for (var j = 0; j < doc.pages.length; j++) {
        var currentpageName = doc.pages[j].name;
        if (currentpageName === pageName) {
            return doc.pages[j];
        }
    }
    var requiredPageCount=parseInt(pageName)+parseInt(currentpageName);
    // 无匹配时新建页面
    newPage = ensureEnoughPages(doc, requiredPageCount);
    return newPage;
}

// 复制元素到目标页面
function copyElementsToPage(sourcePage, targetPage, sourceLayer, includeTextFrames, includeGraphicFrames, includePaths) {
    try {
        // 使用与图层克隆器类似的逻辑，遍历源图层中的页面元素
        var elementCount = sourceLayer.pageItems.length;
        var linkMap = []; // 用于存储文本框链接关系
        var tfMap = {}; // 旧元素ID -> 新元素映射

        // 复制元素
        for (var i = elementCount - 1; i >= 0; i--) {
            var sourceItem = sourceLayer.pageItems.item(i);
            
            // 检查元素是否在当前页面上
            if (sourceItem.parentPage !== sourcePage) {
                continue; // 跳过不在当前页面上的元素
            }
            
            // 检查元素是否被隐藏或锁定
            if (sourceItem.visible === false || sourceItem.locked === true) {
                continue; // 跳过隐藏或锁定的元素
            }
            
            // 根据用户选择过滤元素类型
            var shouldCopy = false;
            if (sourceItem.hasOwnProperty("contents") && includeTextFrames) {
                // 文本框
                shouldCopy = true;
            } else if (sourceItem.hasOwnProperty("graphics") && includeGraphicFrames) {
                // 图像框
                shouldCopy = true;
            } else if (sourceItem.hasOwnProperty("paths") && includePaths) {
                // 路径
                shouldCopy = true;
            }
            
            if (shouldCopy) {
                // 使用duplicate方法直接复制到目标页面
                var targetItem = sourceItem.duplicate(targetPage.parent); // 复制到目标页面的父级（文档）
                                
                // 存储ID映射关系
                tfMap[sourceItem.id] = targetItem;

                // 检查文本框链接关系
                if (typeof sourceItem.previousTextFrame != 'undefined' &&
                    sourceItem.previousTextFrame && sourceItem.textFrameIndex > 0) {
                    linkMap.push({
                        item: sourceItem,
                        prev: sourceItem.previousTextFrame,
                    });
                }
            }
        }

        // 修复文本框链接关系
        repairTextFramesLinks(linkMap, tfMap);
    } catch (e) {
        throw new Error("复制元素失败: " + e.message);
    }
}

// 复制图层核心逻辑
function copyLayerContents(sourceDoc, targetDoc, layerName, pageRangeStr, includeTextFrames, includeGraphicFrames, includePaths) {
    try {
        // 获取源图层
        var sourceLayer = null;
        for (var i = 0; i < sourceDoc.layers.length; i++) {
            if (sourceDoc.layers[i].name === layerName) {
                sourceLayer = sourceDoc.layers[i];
                break;
            }
        }
        
        if (!sourceLayer) {
            throw new Error("在源文档中找不到名为 '" + layerName + "' 的图层");
        }
        
        // 遍历源文档页面
        for (var i = 0; i < sourceDoc.pages.length; i++) {
            var page = sourceDoc.pages[i];
            
            // 若用户指定页码范围且当前页不在范围内，跳过
            if (!isPageInRange(page.name, pageRangeStr)) continue;
            
            // 在目标文档创建同名页面（若不存在）
            var targetPage = createPageByName(targetDoc, page.name);
            
            // 复制图层元素（跳过隐藏/锁定元素）
            copyElementsToPage(page, targetPage, sourceLayer, includeTextFrames, includeGraphicFrames, includePaths);
        }
        
        // 重置选择
        app.select(null);
        
        alert("图层复制完成！");
    } catch (e) {
        alert("复制中断: " + e.message);
    }
}

// 修复文本框链接关系的函数
function repairTextFramesLinks(linkMap, copyMap) {
    for (var c = 0, length = linkMap.length; c < length; c++) {
        var element = linkMap[c];
        var oldItem = element.item;
        var parentItem = copyMap[oldItem.id];
        var newPrevItem;

        if (element.prev) {
            newPrevItem = copyMap[element.prev.id];

            if (parentItem && newPrevItem) {
                parentItem.previousTextFrame = newPrevItem;
            }
        }
    }
}

// 主函数
function main() {
    try {
        // 获取已打开的文档
        var openDocs = getOpenDocuments();
        
        if (openDocs.length < 2) {
            alert("需要至少打开两个文档才能进行复制操作！");
            return;
        }
        
        // 显示对话框
        var dialog = showCopyDialog(openDocs);
        if (!dialog) {
            return; // 文档数量不足
        }
        
        if (dialog.show() === 1) { // 用户点击了确定
            // 获取用户选择
            var sourceDoc = openDocs[dialog.sourceDropdown.selection.index];
            var targetDoc = openDocs[dialog.targetDropdown.selection.index];
            var layerName = dialog.layerDropdown.selection.text;
            var includeTextFrames = dialog.textFramesCheckbox.value;
            var includeGraphicFrames = dialog.graphicFramesCheckbox.value;
            var includePaths = dialog.pathsCheckbox.value;
            var pageRangeStr = dialog.pageRangeEdit.text;
            
            // 执行复制操作
            copyLayerContents(sourceDoc, targetDoc, layerName, pageRangeStr, 
                             includeTextFrames, includeGraphicFrames, includePaths);
        }
        
    } catch (e) {
        alert("脚本执行出错: " + e.message);
        alert("可能是页数不够、源文档没有该图层、源文档没有该页码范围或目标文档没有同名页面");
    }
}

// 运行主函数
KTUDoScriptAsUndoable(function() {main()}, "复制图层内容");
