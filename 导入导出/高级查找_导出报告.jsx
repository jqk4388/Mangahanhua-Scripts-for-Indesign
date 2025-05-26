// 导出查找报告.jsx

// 检查是否有打开的文档
if (app.documents.length === 0) {
    alert("请先打开一个文档。");
} else {
    var doc = app.activeDocument;
    var layers = doc.layers;
    var layerNames = [];
    var i;
    for (i = 0; i < layers.length; i++) {
        layerNames.push(layers[i].name);
    }

    // UI对话框（palette类型，浮动）
    var dlg = new Window("dialog", "查找与导出工具");
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    // 图层选择
    dlg.add("statictext", undefined, "选择图层：");
    var layerDropdown = dlg.add("dropdownlist", undefined, layerNames);
    layerDropdown.selection = 0;

    // 查找内容
    dlg.add("statictext", undefined, "查找内容：");
    var findTypeDropdown = dlg.add("dropdownlist", undefined, ["对象", "文本"]);
    findTypeDropdown.selection = 0;

    // GREP输入（仅文本时显示）
    var grepGroup = dlg.add("group");
    grepGroup.orientation = "row";
    grepGroup.visible = false;
    grepGroup.add("statictext", undefined, "GREP表达式：");
    var grepInput = grepGroup.add("edittext", undefined, "");
    grepInput.characters = 20;

    // 对象类型
    dlg.add("statictext", undefined, "对象类型：");
    var objectTypeDropdown = dlg.add("dropdownlist", undefined, ["图框", "文本框", "未指定的框架"]);
    objectTypeDropdown.selection = 0;

    // 查找范围
    dlg.add("statictext", undefined, "查找范围：");
    var scopeDropdown = dlg.add("dropdownlist", undefined, ["当前页面", "所有页面"]);
    scopeDropdown.selection = 0;

    // 查找方向
    dlg.add("statictext", undefined, "查找方向：");
    var directionDropdown = dlg.add("dropdownlist", undefined, ["向前", "向后"]);
    directionDropdown.selection = 0;

    // 按钮组
    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    var findBtn = btnGroup.add("button", undefined, "查找下一个");
    var exportBtn = btnGroup.add("button", undefined, "导出页面列表");
    var closeBtn = btnGroup.add("button", undefined, "关闭");

    // 变量保存查找结果
    var foundItems = [];
    var currentIndex = -1;

    // 查找函数
    function doSearch() {
        foundItems = [];
        currentIndex = -1;
        var layerName = layerDropdown.selection.text;
        var findType = findTypeDropdown.selection.text;
        var objectType = objectTypeDropdown.selection.text;
        var scope = scopeDropdown.selection.text;
        var grepStr = grepInput.text;

        var pages = [];
        var j, p;
        if (scope === "当前页面") {
            pages.push(app.activeWindow.activePage);
        } else {
            for (i = 0; i < doc.pages.length; i++) {
                pages.push(doc.pages[i]);
            }
        }

        for (p = 0; p < pages.length; p++) {
            var page = pages[p];
            var pageItems = page.allPageItems;
            for (j = 0; j < pageItems.length; j++) {
                var item = pageItems[j];
                if (item.itemLayer && item.itemLayer.name === layerName) {
                    if (findType === "对象") {
                        if (
                            (objectType === "图框" && item.constructor.name === "Rectangle" && item.graphics.length > 0) ||
                            (objectType === "文本框" && item.constructor.name === "TextFrame") ||
                            (objectType === "未指定的框架" && item.constructor.name === "Rectangle" && item.graphics.length === 0 && (!("parentStory" in item) || !item.parentStory))
                        ) {
                            foundItems.push(item);
                        }
                    } else if (findType === "文本") {
                        if (item.constructor.name === "TextFrame" && item.contents !== "") {
                            // GREP查找
                            var stories = [item.parentStory];
                            var k;
                            for (k = 0; k < stories.length; k++) {
                                var story = stories[k];
                                // 备份原查找设置
                                var origGrepFind = app.findGrepPreferences.properties;
                                app.findGrepPreferences = NothingEnum.nothing;
                                app.findGrepPreferences.findWhat = grepStr;
                                var results = story.findGrep();
                                for (var m = 0; m < results.length; m++) {
                                    // 只收集属于当前TextFrame的结果
                                    if (results[m].parentTextFrames.length > 0) {
                                        for (var n = 0; n < results[m].parentTextFrames.length; n++) {
                                            if (results[m].parentTextFrames[n] == item) {
                                                foundItems.push(results[m]);
                                            }
                                        }
                                    }
                                }
                                app.findGrepPreferences = NothingEnum.nothing;
                                app.findGrepPreferences.properties = origGrepFind;
                            }
                        }
                    }
                }
            }
        }
    }

    // 查找下一个
    findBtn.onClick = function() {
        if (foundItems.length === 0) {
            doSearch();
            if (foundItems.length === 0) {
                alert("未找到匹配项。");
                return;
            }
        }
        var dir = directionDropdown.selection.text;
        if (dir === "向前") {
            currentIndex--;
            if (currentIndex < 0) currentIndex = foundItems.length - 1;
        } else {
            currentIndex++;
            if (currentIndex >= foundItems.length) currentIndex = 0;
        }
        var item = foundItems[currentIndex];
        // 跳转并选中
        if (item.hasOwnProperty("parentPage") && item.parentPage) {
            app.activeWindow.activePage = item.parentPage;
        } else if (item.parentTextFrames && item.parentTextFrames.length > 0 && item.parentTextFrames[0].parentPage) {
            app.activeWindow.activePage = item.parentTextFrames[0].parentPage;
        }
        if (item.hasOwnProperty("select")) {
            item.select();
        } else if (item.hasOwnProperty("characters")) {
            // GREP结果为文本对象
            app.select(item);
        }
        app.activeWindow.zoom(ZoomOptions.FIT_PAGE);
        // alert("第 " + (currentIndex + 1) + " 个，共 " + foundItems.length + " 个。");
    };

    // 导出页面列表
    exportBtn.onClick = function() {
        if (foundItems.length === 0) {
            doSearch();
            if (foundItems.length === 0) {
                alert("未找到匹配项，无法导出。");
                return;
            }
        }
        var pageNumbers = {};
        var i;
        for (i = 0; i < foundItems.length; i++) {
            var pg = null;
            if (foundItems[i].hasOwnProperty("parentPage") && foundItems[i].parentPage) {
                pg = foundItems[i].parentPage;
            } else if (foundItems[i].parentTextFrames && foundItems[i].parentTextFrames.length > 0 && foundItems[i].parentTextFrames[0].parentPage) {
                pg = foundItems[i].parentTextFrames[0].parentPage;
            }
            if (pg) pageNumbers[pg.name] = true;
        }
        var pageList = [];
        for (var key in pageNumbers) {
            pageList.push(key);
        }
        pageList.sort();

        var desktopPath = Folder.desktop.fsName;
        var fileName = "查找页面列表.txt";
        var filePath = desktopPath + "/" + fileName;
        var file = new File(filePath);
        file.encoding = "UTF-8";
        file.open("w");
        file.write(pageList.join("\r\n"));
        file.close();
        alert("页面列表已导出到桌面：" + filePath);
    };

    closeBtn.onClick = function() {
        dlg.close();
    };

    // 切换查找内容时，控制对象类型和GREP输入可用性
    findTypeDropdown.onChange = function() {
        var isObj = (findTypeDropdown.selection.text === "对象");
        objectTypeDropdown.enabled = isObj;
        grepGroup.visible = !isObj;
        // 选项变更时清空查找并重新查找
        foundItems = [];
        currentIndex = -1;
        doSearch();
    };
    objectTypeDropdown.enabled = true;
    grepGroup.visible = false;

    // 其它下拉列表变更时也清空查找并重新查找
    layerDropdown.onChange = function() {
        foundItems = [];
        currentIndex = -1;
        doSearch();
    };
    objectTypeDropdown.onChange = function() {
        foundItems = [];
        currentIndex = -1;
        doSearch();
    };
    scopeDropdown.onChange = function() {
        foundItems = [];
        currentIndex = -1;
        doSearch();
    };
    directionDropdown.onChange = function() {
        foundItems = [];
        currentIndex = -1;
        doSearch();
    };

    // GREP输入变更时也清空查找并重新查找
    grepInput.onChanging = function() {
        foundItems = [];
        currentIndex = -1;
        doSearch();
    };

    dlg.center();
    dlg.show();

    // 添加一个全局的事件监听器，防止对话框意外关闭
    app.addEventListener(EventType.CLOSE, function (event) {
        if (event.target === dlg) {
            event.preventDefault();
        }
    });
}