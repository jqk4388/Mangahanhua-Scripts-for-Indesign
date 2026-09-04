/*
 * InDesign 漫畫輔助：磁力吸附操作腳本 v14 (極速優化版)
 * 效能修正：
 * 1. 【搜尋範圍】從「全文檔」縮小至「當前頁面」。
 * 2. 【解決卡頓】即使文檔有數百頁，執行速度也如同單頁般瞬間完成。
 * 3. 依然支援多選。
 */

main();

function main() {
    var doc = app.activeDocument;
    var targetLayerName = "[輔助定位] 完整矩形框";
    var ghostLayer = doc.layers.itemByName(targetLayerName);

    if (!ghostLayer.isValid) {
        alert("找不到輔助圖層！請先執行生成辅助框腳本 (v12)。");
        return;
    }

    // 1. 獲取所有選取物件
    try {
        // 显示UI界面
        var userOptions = showUI();
        if (!userOptions) {
            alert("操作已取消。");
            return;
        }

        // 根据用户选择的范围执行断句
        var textFrames = getTextFrames(userOptions.range);

        if (textFrames.length === 0) {
            alert("未找到任何文本框。");
            return;
        }


    // 2. 收集所有「文字框 ↔ 輔助框」的相交配對，並記錄兩者中心點的距離
    var pairs = [];

    for (var s = 0; s < textFrames.length; s++) {
        var selObj = textFrames[s];

        // 只處理未鎖定的物件 (避免選到背景)
        if (selObj.locked) continue;

        // 【關鍵優化】：獲取選取物件所在的「容器」(通常是 Page 或 Spread)
        var parentPage = selObj.parent;

        // 安全檢查：如果選取物件不在頁面上 (例如在剪貼板)，則跳過
        if (!parentPage || !parentPage.pageItems) continue;

        // 【關鍵優化】：只獲取「當前頁面」上的所有物件
        var localItems = parentPage.pageItems;

        // 在當前頁面的物件中，找出所有「輔助框」且「相交」的配對
        for (var i = 0; i < localItems.length; i++) {
            var item = localItems[i];

            // 快速過濾 1: 必須是輔助圖層的物件
            if (item.itemLayer.name !== targetLayerName) continue;

            // 快速過濾 2: 判定相交，相交則記錄為候選配對
            if (isIntersecting(selObj, item)) {
                pairs.push({
                    tf: selObj,
                    ghost: item,
                    dist: centerDistance(selObj, item)
                });
            }
        }
    }

    // 3. 依中心距離「由近到遠」排序，再逐組貪婪配對：
    //    每個輔助框只吸附距離最近的一文字框，其餘較遠的文字框保持不動；
    //    每個文字框也只會被移動一次（對齊到離它最近的輔助框），避免重複移動。
    pairs.sort(function (a, b) {
        return a.dist - b.dist;
    });

    var usedGhosts = {};
    var usedTfs = {};
    var alignCount = 0;

    for (var p = 0; p < pairs.length; p++) {
        var pair = pairs[p];
        var ghostKey = pair.ghost.id;
        var tfKey = pair.tf.id;

        // 此輔助框已吸附過更近的文字框，或此文字框已對齊過，則跳過
        if (usedGhosts[ghostKey] || usedTfs[tfKey]) continue;

        usedGhosts[ghostKey] = true;
        usedTfs[tfKey] = true;
        alignCenter(pair.tf, pair.ghost);
        alignCount++;
    }
    
        // 隐藏ghostLayer
        ghostLayer.visible = false;
    } catch (e) {
        alert("发生错误: " + e.message);
    }
}

// --- 數學運算區 (無變更) ---

function isIntersecting(obj1, obj2) {
    var b1 = obj1.geometricBounds; // [y1, x1, y2, x2]
    var b2 = obj2.geometricBounds;

    if (b1[3] < b2[1] || b1[1] > b2[3] || b1[2] < b2[0] || b1[0] > b2[2]) {
        return false;
    }
    return true;
}

function alignCenter(obj1, obj2) {
    var b1 = obj1.geometricBounds;
    var b2 = obj2.geometricBounds;

    var w1 = b1[3] - b1[1];
    var h1 = b1[2] - b1[0];
    var w2 = b2[3] - b2[1];
    var h2 = b2[2] - b2[0];

    var center2_X = b2[1] + (w2 / 2);
    var center2_Y = b2[0] + (h2 / 2);

    var new_X1 = center2_X - (w1 / 2);
    var new_Y1 = center2_Y - (h1 / 2);
    var new_X2 = new_X1 + w1;
    var new_Y2 = new_Y1 + h1;

    obj1.geometricBounds = [new_Y1, new_X1, new_Y2, new_X2];
}

// 計算兩個物件中心點之間的距離（歐式距離）
function centerDistance(obj1, obj2) {
    var c1 = getCenter(obj1);
    var c2 = getCenter(obj2);

    var dx = c1.x - c2.x;
    var dy = c1.y - c2.y;

    return Math.sqrt(dx * dx + dy * dy);
}

// 取得物件的中心點座標
function getCenter(obj) {
    var b = obj.geometricBounds; // [y1, x1, y2, x2]

    return {
        x: b[1] + (b[3] - b[1]) / 2,
        y: b[0] + (b[2] - b[0]) / 2
    };
}

// 显示UI界面
function showUI() {
    var dialog = new Window("dialog", "大语言模型断句");

    // 断句范围选项
    dialog.add("statictext", undefined, "断句范围：");
    var rangeGroup = dialog.add("group");
    var currentSelectionRadio = rangeGroup.add("radiobutton", undefined, "当前选中的文本框");
    var currentPageRadio = rangeGroup.add("radiobutton", undefined, "当页所有文本框");
    var entireDocumentRadio = rangeGroup.add("radiobutton", undefined, "文档中所有文本框");
    currentSelectionRadio.value = true; // 默认选中第一个选项

    // 确定和取消按钮
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    var okButton = buttonGroup.add("button", undefined, "确定", { name: "ok" });
    var cancelButton = buttonGroup.add("button", undefined, "取消", { name: "cancel" });

    // 处理用户交互
    if (dialog.show() === 1) {
        var range;
        if (currentSelectionRadio.value) {
            range = "currentSelection";
        } else if (currentPageRadio.value) {
            range = "currentPage";
        } else {
            range = "entireDocument";
        }

        return {
            range: range,

        };
    } else {
        return null;
    }
}

// 获取指定范围的文本框（只在当前激活图层中查找）
function getTextFrames(range) {
    var textFrames = [];
    var doc = app.activeDocument;
    
    // 检查并激活正确的图层
    var activeLayer = doc.activeLayer;
    var layerName = activeLayer.name;
    
    // 如果当前激活图层不是 "Text" 或 "文字"，则查找并激活
    if (layerName !== "Text" && layerName !== "文字") {
        var foundLayer = null;
        
        // 查找名为 "Text" 的图层
        var textLayer = doc.layers.itemByName("Text");
        if (textLayer.isValid) {
            foundLayer = textLayer;
        }
        
        // 如果没找到，查找名为 "文字" 的图层
        if (!foundLayer) {
            var textLayerCN = doc.layers.itemByName("文字");
            if (textLayerCN.isValid) {
                foundLayer = textLayerCN;
            }
        }
        
        // 激活找到的图层
        if (foundLayer) {
            doc.activeLayer = foundLayer;
            activeLayer = foundLayer;
        }
    }

    if (range === "currentSelection") {
        for (var i = 0; i < app.selection.length; i++) {
            if (app.selection[i].constructor.name === "TextFrame") {
                // 只添加当前激活图层的文本框
                if (app.selection[i].itemLayer && app.selection[i].itemLayer.id === activeLayer.id) {
                    textFrames.push(app.selection[i]);
                }
            }
        }
    } else if (range === "currentPage") {
        var currentPage = app.activeWindow.activePage;
        var allFramesOnPage = currentPage.textFrames.everyItem().getElements();
        // 只取当前激活图层的文本框
        for (var j = 0; j < allFramesOnPage.length; j++) {
            if (allFramesOnPage[j].itemLayer && allFramesOnPage[j].itemLayer.id === activeLayer.id) {
                textFrames.push(allFramesOnPage[j]);
            }
        }
    } else if (range === "entireDocument") {
        var allTextFrames = doc.textFrames.everyItem().getElements();

        for (var k = 0; k < allTextFrames.length; k++) {
            var tf = allTextFrames[k];
            try {
                // 必须在当前激活图层
                if (!tf.itemLayer || tf.itemLayer.id !== activeLayer.id) continue;
                // 如果 textFrame 所在的 parentPage 存在，并且该页面不是主页
                if (tf.parentPage != null && tf.parentPage.parent.constructor.name != "MasterSpread") {
                    textFrames.push(tf);
                }
            } catch (e) {
                // 忽略错误继续
            }
        }
    }

    return textFrames;
}