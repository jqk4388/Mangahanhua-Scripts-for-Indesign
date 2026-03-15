/*
 * InDesign 漫畫輔助：v12
 * * 更新內容：
 * 1. 【鎖定】生成後自動鎖定，防止移動。
 * 2. 【層級】置於最頂層，確保不被圖片遮擋。
 * 3. 【簡化】移除虛線樣式代碼，使用實線。
 */

main();

function main() {
    if (app.documents.length === 0) {
        alert("請先打開文檔！");
        return;
    }
    var doc = app.activeDocument;
    var GHOST_LABEL = "GhostFrame";

    // 1. 清除舊結果
    removeGhostItems(doc, GHOST_LABEL);
    var oldLayerNames = ["[輔助定位] 幽靈靶心", "[輔助定位] 幽靈矩形", "[輔助定位] 完整矩形框", "GhostFrames"];
    for (var L = 0; L < oldLayerNames.length; L++) {
        var oldL = doc.layers.itemByName(oldLayerNames[L]);
        if (oldL.isValid) {
            oldL.locked = false;
            oldL.remove();
        }
    }

    // 2. 準備綠色
    var colorName = "GhostGreen";
    var ghostColor = doc.colors.itemByName(colorName);
    if (!ghostColor.isValid) {
        ghostColor = doc.colors.add({
            name: colorName,
            model: ColorModel.PROCESS,
            space: ColorSpace.RGB,
            colorValue: [0, 255, 0]
        });
    }

    // 3. 建立圖層
    var layerName = "[輔助定位] 完整矩形框";
    var ghostLayer = doc.layers.add({name: layerName});
    ghostLayer.move(LocationOptions.AT_BEGINNING); // 最頂層
    ghostLayer.printable = false;
    ghostLayer.layerColor = UIColors.GREEN; 
    ghostLayer.locked = false; // 暫時解鎖以寫入

    // 4. 讀取與生成
    var jsonFile;
    var skipDialogs = app.extractLabel("MasterRunner.skipDialogs");
    
    if (skipDialogs === "true") {
        // 從主腳本獲取 JSON 路徑（輔助框用 JSON）
        var jsonPath = app.extractLabel("MasterRunner.jsonGhostFilePath");
        if (!jsonPath) {
            alert("錯誤：無法從主腳本獲取輔助框 JSON 路徑！");
            return;
        }
        jsonFile = new File(jsonPath);
        if (!jsonFile.exists) {
            alert("錯誤：輔助框 JSON 檔案不存在：" + jsonPath);
            return;
        }
    } else {
        jsonFile = File.openDialog("請選擇 JSON", "*.json");
        if (!jsonFile) return;
    }
    
    var jsonData = readJSON(jsonFile);
    if (!jsonData) return;
    
    // 檢查必要結構
    if (!jsonData['image_info']) {
        alert("錯誤：JSON 檔案缺少 image_info！\n\n支援格式：\n{ \"image_info\": { \"01.png\": { \"width\": 1644, \"height\": 2400 }, ... }, \"pages\": {...} }");
        return;
    }
    if (!jsonData['pages']) {
        alert("錯誤：JSON 檔案缺少 pages！");
        return;
    }

    var pages = doc.pages;
    var totalRects = 0;
    
    // 建立進度條
    var progressWin = new Window("palette", "處理中...");
    var progressBar = progressWin.add("progressbar", undefined, 0, pages.length);
    progressBar.preferredSize.width = 300;
    progressWin.show();
    
    for (var i = 0; i < pages.length; i++) {
        progressBar.value = i + 1;
        var page = pages[i];
        var pageImageName = getPageImageName(page);
        if (!pageImageName) continue;

        var matchedKey = findKeyInJson(jsonData.pages, pageImageName.replace(/\.[^\.]+$/, ""));
        
        if (matchedKey) {
            // 從 image_info 取得該頁的寬高
            var imgInfo = jsonData['image_info'][matchedKey];
            if (!imgInfo) {
                // 嘗試用完整檔名
                imgInfo = jsonData['image_info'][pageImageName];
            }
            if (!imgInfo || !imgInfo.width || !imgInfo.height) {
                continue; // 跳過沒有尺寸資訊的頁面
            }
            
            var textBlocks = jsonData.pages[matchedKey];
            if (textBlocks && textBlocks.length > 0) {
                totalRects += createRects(page, textBlocks, ghostLayer, imgInfo.width, imgInfo.height, ghostColor, GHOST_LABEL);
            }
        }
    }
    
    progressWin.close();

    // 5. 最後鎖定
    ghostLayer.locked = true;

    // 只在獨立執行時顯示 alert
    var skipDialogs = app.extractLabel("MasterRunner.skipDialogs");
    if (skipDialogs !== "true") {
        alert("v11 完成！\n已生成 " + totalRects + " 個綠色實線框。\n圖層已鎖定。\n\n【如何實現中心對中心吸附？】\n請確認：\n1. [檢視] > [靠齊文件格點] 必須關閉。\n2. [偏好設定] > [智能輔助線] > [對齊物件中心] 必須勾選。\n\n拖曳時看到十字交叉線即代表中心對齊。");
    }
}

// ---------------- 核心功能區 ----------------

function createRects(page, blocks, layer, origW, origH, colorObj, label) {
    var bounds = page.bounds; 
    var pageH_mm = bounds[2] - bounds[0];
    var pageW_mm = bounds[3] - bounds[1];
    var scaleX = pageW_mm / origW;
    var scaleY = pageH_mm / origH;
    
    var created = 0;
    for (var i = 0; i < blocks.length; i++) {
        var b = blocks[i];
        if (!b.xyxy) continue;
        
        var pix = b.xyxy; 
        var y1 = pix[1] * scaleY;
        var x1 = pix[0] * scaleX;
        var y2 = pix[3] * scaleY;
        var x2 = pix[2] * scaleX;

        var tf = page.textFrames.add(layer);
        tf.label = label;
        tf.geometricBounds = [y1, x1, y2, x2];
        
        tf.fillColor = "None";
        tf.strokeColor = colorObj; 
        tf.strokeWeight = 1.5;
        tf.textWrapPreferences.textWrapMode = TextWrapModes.NONE;
        created++;
    }
    return created;
}

function removeGhostItems(doc, label) {
    var frames = doc.textFrames;
    for (var i = frames.length - 1; i >= 0; i--) {
        var tf = frames[i];
        if (tf.label === label) {
            try {
                if (tf.itemLayer && tf.itemLayer.locked) tf.itemLayer.locked = false;
                tf.remove();
            } catch (e) {}
        }
    }
}

function readJSON(file) {
    file.open("r");
    var content = file.read();
    file.close();
    try { return eval("(" + content + ")"); } catch (e) { return null; }
}

function getPageImageName(page) {
    var graphics = page.allGraphics;
    for (var j = 0; j < graphics.length; j++) {
        var g = graphics[j];
        if (g.parent.geometricBounds[1] >= -0.1) { 
            if (g.itemLink) return g.itemLink.name;
        }
    }
    return null;
}

function findKeyInJson(pagesObj, baseName) {
    for (var key in pagesObj) {
        if (pagesObj.hasOwnProperty(key)) {
            var keyBase = key.replace(/\.[^\.]+$/, "");
            if (keyBase === baseName) return key;
        }
    }
    return null;
}