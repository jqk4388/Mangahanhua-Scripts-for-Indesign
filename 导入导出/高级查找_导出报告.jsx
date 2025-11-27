// 导出查找报告.jsx
#targetengine "session"
var dlg;

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
    dlg = new Window("palette", "查找与导出工具");
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

    // 查找格式分组（仅文本时显示）
    var formatGroup = dlg.add("panel", undefined, "查找格式");
    formatGroup.orientation = "column";
    formatGroup.alignChildren = "left";
    formatGroup.visible = false;

    var enableFormatCheckbox = formatGroup.add("checkbox", undefined, "启用查找格式");
    enableFormatCheckbox.value = false;

    // 获取字符样式和段落样式（递归，含分组）
    var charStyleNames = ["(无)"];
    var paraStyleNames = ["(无)"];
    var charStyleObjs = [null];
    var paraStyleObjs = [null];

    function countCharStyles(styles, groupName) {
        for (var i = 0; i < styles.length; i++) {
            if (styles[i].constructor.name === "CharacterStyleGroup") {
                var gName = styles[i].name;
                countCharStyles(styles[i].characterStyles, gName);
                countCharStyles(styles[i].characterStyleGroups, gName);
            } else if (styles[i].constructor.name === "CharacterStyle") {
                var name = styles[i].name;
                var CStyle = styles[i];
                var group = groupName && groupName !== "" ? groupName : "";
                if (group !== "") {
                    name = name + "(" + group + ")";
                }
                charStyleNames.push(name);
                charStyleObjs.push(CStyle);
            }
        }
    }
    function countParaStyles(styles, groupName) {
        for (var i = 0; i < styles.length; i++) {
            if (styles[i].constructor.name === "ParagraphStyleGroup") {
                var gName = styles[i].name;
                countParaStyles(styles[i].paragraphStyles, gName);
                countParaStyles(styles[i].paragraphStyleGroups, gName);
            } else if (styles[i].constructor.name === "ParagraphStyle") {
                if (styles[i].name !== "[基本段落]" && styles[i].name !== "[无]") {
                    var name = styles[i].name;
                    var PStyle = styles[i];
                    var group = groupName && groupName !== "" ? groupName : "";
                    if (group !== "") {
                        name = name + "(" + group + ")";
                    }
                    paraStyleNames.push(name);
                    paraStyleObjs.push(PStyle);
                }
            }
        }
    }
    // 统计组外样式和组内样式
    countCharStyles(doc.characterStyles, "");
    countCharStyles(doc.characterStyleGroups, "");
    countParaStyles(doc.paragraphStyles, "");
    countParaStyles(doc.paragraphStyleGroups, "");

    var charStyleGroup = formatGroup.add("group");
    charStyleGroup.add("statictext", undefined, "字符样式：");
    var charStyleDropdown = charStyleGroup.add("dropdownlist", undefined, charStyleNames);
    charStyleDropdown.selection = 0;
    charStyleDropdown.enabled = false;

    var paraStyleGroup = formatGroup.add("group");
    paraStyleGroup.add("statictext", undefined, "段落样式：");
    var paraStyleDropdown = paraStyleGroup.add("dropdownlist", undefined, paraStyleNames);
    paraStyleDropdown.selection = 0;
    paraStyleDropdown.enabled = false;

    enableFormatCheckbox.onClick = function() {
        var enabled = enableFormatCheckbox.value;
        charStyleDropdown.enabled = enabled;
        paraStyleDropdown.enabled = enabled;
        foundItems = [];
        currentIndex = -1;
    };
    charStyleDropdown.onChange = function() {
        foundItems = [];
        currentIndex = -1;
    };
    paraStyleDropdown.onChange = function() {
        foundItems = [];
        currentIndex = -1;
    };

    // 对象类型
    dlg.add("statictext", undefined, "对象类型：");
    var objectTypeDropdown = dlg.add("dropdownlist", undefined, [
        "图框", "文本框", "未指定的框架", "矩形", "直线", "多边形", "椭圆", "组", "表格", "按钮", "复合路径", "图像", "QR码", "文本路径", "注释"
    ]);
    objectTypeDropdown.selection = 0;

    // 查找范围
    dlg.add("statictext", undefined, "查找范围：");
    var scopeDropdown = dlg.add("dropdownlist", undefined, ["当前页面", "所有页面"]);
    scopeDropdown.selection = 1;

    // 查找方向
    dlg.add("statictext", undefined, "查找方向：");
    var directionDropdown = dlg.add("dropdownlist", undefined, ["向前", "向后"]);
    directionDropdown.selection = 0;

    // 按钮组
    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    var findBtn = btnGroup.add("button", undefined, "查找下一个");
    var exportBtn = btnGroup.add("button", undefined, "导出页面列表");
    var highlightBtn = btnGroup.add("button", undefined, "突出显示");
    var closeBtn = btnGroup.add("button", undefined, "关闭");

    // 变量保存查找结果
    var foundItems = [];
    var currentIndex = -1;
    var highlightApplied = false;
    
    // 查找函数
    function doSearch() {
        foundItems = [];
        currentIndex = -1;
        var layerName = layerDropdown.selection.text;
        var findType = findTypeDropdown.selection.text;
        var objectType = objectTypeDropdown.selection.text;
        var scope = scopeDropdown.selection.text;
        var grepStr = grepInput.text;

        // 查找格式相关
        var enableFormat = enableFormatCheckbox.value;
        var charStyleObj = charStyleDropdown.selection ? charStyleObjs[charStyleDropdown.selection.index] : null;
        var paraStyleObj = paraStyleDropdown.selection ? paraStyleObjs[paraStyleDropdown.selection.index] : null;

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
                            (objectType === "未指定的框架" && item.constructor.name === "Rectangle" && item.graphics.length === 0 && (!("parentStory" in item) || !item.parentStory)) ||
                            (objectType === "矩形" && item.constructor.name === "Rectangle") ||
                            (objectType === "直线" && item.constructor.name === "GraphicLine") ||
                            (objectType === "多边形" && item.constructor.name === "Polygon") ||
                            (objectType === "椭圆" && item.constructor.name === "Oval") ||
                            (objectType === "组" && item.constructor.name === "Group") ||
                            (objectType === "表格" && item.constructor.name === "Table") ||
                            (objectType === "按钮" && item.constructor.name === "Button") ||
                            (objectType === "复合路径" && item.constructor.name === "CompoundPath") ||
                            (objectType === "图像" && item.constructor.name === "Image") ||
                            (objectType === "QR码" && item.constructor.name === "Rectangle" && item.graphics.length > 0 && item.graphics[0].constructor.name === "QRCode") ||
                            (objectType === "文本路径" && item.constructor.name === "TextPath") ||
                            (objectType === "注释" && item.constructor.name === "Note")
                        ) {
                            foundItems.push(item);
                        }
                    } else if (findType === "文本") {
                        if (item.constructor.name === "TextFrame" && item.contents !== "") {
                            var stories = [item.parentStory];
                            var k;
                            for (k = 0; k < stories.length; k++) {
                                var story = stories[k];
                                // 备份原查找设置
                                var origGrepFind = app.findGrepPreferences.properties;
                                app.findGrepPreferences = NothingEnum.nothing;
                                app.findGrepPreferences.findWhat = grepStr;

                                // 应用查找格式（用对象）
                                if (enableFormat) {
                                    if (charStyleObj) {
                                        app.findGrepPreferences.appliedCharacterStyle = charStyleObj;
                                    }
                                    if (paraStyleObj) {
                                        app.findGrepPreferences.appliedParagraphStyle = paraStyleObj;
                                    }
                                }

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
        file.write("当前文件名：" + doc.name + "\r\n");
        file.write("查找条件：\r\n");
        file.write("图层：" + layerDropdown.selection.text + "\r\n");
        file.write("查找内容：" + findTypeDropdown.selection.text + "\r\n");
        if (findTypeDropdown.selection.text === "对象") {
            file.write("对象类型：" + objectTypeDropdown.selection.text + "\r\n");
        } else {
            file.write("GREP表达式：" + grepInput.text + "\r\n");
            if (enableFormatCheckbox.value) {
                file.write("查找格式：\r\n");
                file.write("  字符样式：" + charStyleDropdown.selection.text + "\r\n");
                file.write("  段落样式：" + paraStyleDropdown.selection.text + "\r\n");
            }
        }
        file.write("查找范围：" + scopeDropdown.selection.text + "\r\n");
        file.write("查找方向：" + directionDropdown.selection.text + "\r\n\r\n");
        file.write("单页导出用：\r\n");
        file.write(pageList.join(",")+"\r\n\r\n");
        file.write("列表：\r\n");
        file.write(pageList.join("\r\n"));
        file.close();
        alert("页面列表已导出到桌面：" + filePath);
    };

    // 突出显示按钮：使用条件文本（Condition）来标记/取消标记找到的文本
    highlightBtn.onClick = function() {
        if (foundItems.length === 0) {
            doSearch();
            if (foundItems.length === 0) {
                alert("未找到匹配项，无法高亮。");
                return;
            }
        }
        var csName = "高级查找";
        var condName = "高亮查找";
        var cs;
        var cond;
        // 获取或创建条件集
        try {
            cs = doc.conditionSets.item(csName);
            cs.name; // 触发异常以检测是否存在
        } catch (e) {
            cs = doc.conditionSets.add();
            cs.name = csName;
            doc.conditionalTextPreferences.showConditionIndicators = ConditionIndicatorMode.SHOW_INDICATORS;
            doc.conditionalTextPreferences.activeConditionSet = cs;
        }
        // 获取或创建条件
        try {
            cond = doc.conditions.item(condName);
            cond.name;
        } catch (e) {
            cond = doc.conditions.add();
            cond.name = condName;
            cond.indicatorMethod = ConditionIndicatorMethod.USE_HIGHLIGHT
            cond.indicatorColor = UIColors.GOLD;
            cond.visible = true;
        }

        if (!highlightApplied) {
            // 应用条件到所有找到的项
            for (var i = 0; i < foundItems.length; i++) {
                var it = foundItems[i];
                try {
                    if (it.appliedConditions) {
                        it.applyConditions([cond], true);
                    }
                } catch (e) {
                    // 忽略单项应用错误
                }
            }
            highlightApplied = true;
            highlightBtn.text = "取消高亮";
        } else {
            // 移除条件
            for (var i = 0; i < foundItems.length; i++) {
                var it = foundItems[i];
                try {
                    if (it.appliedConditions) {
                        it.applyConditions([], true);
                    }
                } catch (e) {
                    // 忽略单项移除错误
                }
            }
            highlightApplied = false;
            highlightBtn.text = "突出显示";
        }
    };

    closeBtn.onClick = function() {
        dlg.close();
    };

    // 切换查找内容时，控制对象类型、GREP输入和查找格式可用性
    findTypeDropdown.onChange = function() {
        var isObj = (findTypeDropdown.selection.text === "对象");
        objectTypeDropdown.enabled = isObj;
        grepGroup.visible = !isObj;
        formatGroup.visible = !isObj;
        // 选项变更时清空查找并重新查找
        foundItems = [];
        currentIndex = -1;
        doSearch();
    };
    objectTypeDropdown.enabled = true;
    grepGroup.visible = false;
    formatGroup.visible = false;

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
    // app.addEventListener(EventType.CLOSE, function (event) {
    //     if (event.target === dlg) {
    //         event.preventDefault();
    //     }
    // });
}