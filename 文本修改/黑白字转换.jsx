/*
    脚本功能：将选中的文本框内的黑色文字转换为白色，
    白色文字转换为黑色，其它颜色保持不变。
    
    注意：
    1. 脚本依赖文档中存在名为 "Black" 和 "White" 的色板，
         否则会弹出提示并退出。
    2. 为避免转换混乱，先将黑色文字先转换为临时色板，
         然后将白色文字转换为黑色，再将临时颜色转换为白色。
*/

if (app.documents.length === 0) {
        alert("请先打开一个文档");
        exit();
}

var doc = app.activeDocument;
if (app.selection.length === 0) {
        alert("请先选择一个文本框");
        exit();
}

var blackSwatch, whiteSwatch;
try {
        blackSwatch = doc.swatches.itemByName("Black");
        // 尝试读取属性以确保该色板存在
        var dummy = blackSwatch.name;
} catch (e) {
        // alert("文档中没有名为 'Black' 的色板");
        exit();
}
try {
        whiteSwatch = doc.swatches.itemByName("Paper");
        var dummy = whiteSwatch.name;
} catch (e) {
        // alert("文档中没有名为 'White' 的色板");
        exit();
}

// 创建或获取临时色板，用于中转黑色→白色的转换
var tempSwatchName = "TempSwapColor";
var tempSwatch;
if (doc.swatches.itemByName(tempSwatchName).isValid) {
        tempSwatch = doc.swatches.itemByName(tempSwatchName);
} else {
        try {
                // 新建一个临时色板，颜色值不会影响后续转换
                tempSwatch = doc.colors.add({ name: tempSwatchName, model: ColorModel.process, colorValue: [0, 0, 0, 0] });
        } catch (e) {
                alert("无法创建临时色板");
                exit();
        }
}

for (var i = 0; i < app.selection.length; i++) {
        var selObj = app.selection[i];
        if (selObj.constructor.name === "TextFrame") {
                // 1. 将所有黑色文本替换为临时色板
                app.findTextPreferences = NothingEnum.nothing;
                app.changeTextPreferences = NothingEnum.nothing;
                app.findTextPreferences.fillColor = blackSwatch;
                var foundTexts = selObj.findText();
                for (var j = 0; j < foundTexts.length; j++) {
                        foundTexts[j].fillColor = tempSwatch;
                }

                // 2. 将所有白色文本替换为黑色
                app.findTextPreferences = NothingEnum.nothing;
                app.changeTextPreferences = NothingEnum.nothing;
                app.findTextPreferences.fillColor = whiteSwatch;
                foundTexts = selObj.findText();
                for (var j = 0; j < foundTexts.length; j++) {
                        foundTexts[j].fillColor = blackSwatch;
                }

                // 3. 将临时色板（原来是黑色的）替换为白色
                app.findTextPreferences = NothingEnum.nothing;
                app.changeTextPreferences = NothingEnum.nothing;
                app.findTextPreferences.fillColor = tempSwatch;
                foundTexts = selObj.findText();
                for (var j = 0; j < foundTexts.length; j++) {
                        foundTexts[j].fillColor = whiteSwatch;
                }
        } else {
                alert("所选对象不是文本框");
        }
}

// 清除查找和替换设置
app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;

// alert("黑白字转换完成");