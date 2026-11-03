// ----------- Refit Overset Frames ----------- // 
/* 
    Refits all of the overflowing frames on either the current page or all pages. 

    Usage Instructions: 
    - Select one or many text frames that are overset
    - Run this script

    This is the same thing as doing `Object > Fitting > Fit Frame to Content`, but on a bunch of text frames at once. 
    UPDATE: Now the script will expand the frame by 1 pixel until it's no longer overset
    新增: 勾选「对所有文本框执行【框架适合内容】」后，页面上的所有文本框都会执行适合内容；
          不勾选时只处理溢流文本框。

    Installation Instructions: https://github.com/saraoswald/Manga-Scripts/#how-to-use-scripts-in-indesign

    Nov 21 2020, Sara Linsley
*/

/* ------ Start of Script ------ */

function main() {
    if (app.documents.length === 0) {
        alert("请先打开一个文档。");
        return;
    }

    var options = showOptionsDialog();
    if (!options) return; // 用户取消

    var doc = app.activeDocument;
    var usersUnits = app.scriptPreferences.measurementUnit; // so we can revert 'em back later
    app.scriptPreferences.measurementUnit = MeasurementUnits.PIXELS;

    try {
        if (options.allPages) { // if "All Pages" is selected
            fitAllPages(doc, options.fitAll);
        } else { // if "This Page" is selected
            var activeWin = app.activeWindow;
            if (!activeWin || !activeWin.activePage) {
                alert("无法获取当前页面，请重试。");
                return;
            }
            fitFramesOnPage(activeWin.activePage, options.fitAll);
        }
    } catch (err) {
        alert(err);
    } finally {
        app.scriptPreferences.measurementUnit = usersUnits;
    }
}

// 选项窗口：页面范围 + 是否对所有文本框执行【框架适合内容】
function showOptionsDialog() {
    var dialogWindow = new Window("dialog", "Refit Overset Frames");
    dialogWindow.orientation = "column";
    dialogWindow.alignChildren = "left";

    dialogWindow.add("statictext", undefined, "Fit all the overset frames on:");

    var pageGroup = dialogWindow.add("group");
    pageGroup.orientation = "column";
    pageGroup.alignChildren = "left";
    var thisPageRadio = pageGroup.add("radiobutton", undefined, "This Page");
    thisPageRadio.value = true;
    var allPagesRadio = pageGroup.add("radiobutton", undefined, "All Pages");

    var fitAllCheckbox = dialogWindow.add("checkbox", undefined, "对所有文本框执行【框架适合内容】");
    fitAllCheckbox.value = false;

    var buttonGroup = dialogWindow.add("group");
    buttonGroup.alignment = "center";
    var okButton = buttonGroup.add("button", undefined, "确定");
    var cancelButton = buttonGroup.add("button", undefined, "取消");

    var options = null;
    okButton.onClick = function () {
        options = {
            allPages: allPagesRadio.value,
            fitAll: fitAllCheckbox.value
        };
        dialogWindow.close();
    };
    cancelButton.onClick = function () {
        dialogWindow.close();
    };

    dialogWindow.show();
    return options;
}

// 遍历所有页面，并显示进度条
function fitAllPages(doc, fitAll) {
    var progressWindow = new Window("palette", "Refitting Overset Frames...");
    progressWindow.orientation = "column";
    progressWindow.alignChildren = "left";

    var progressText = progressWindow.add("statictext", undefined, "0 / " + doc.pages.length);
    progressText.preferredSize.width = 250;
    var progressBar = progressWindow.add("progressbar", undefined, 0, doc.pages.length);
    progressBar.preferredSize.width = 250;

    try {
        progressWindow.show();
        for (var i = 0; i < doc.pages.length; i++) {
            fitFramesOnPage(doc.pages[i], fitAll);
            progressBar.value = i + 1;
            progressText.text = (i + 1) + " / " + doc.pages.length;
            progressWindow.update();
        }
    } finally {
        progressWindow.close();
    }
}

// 判断文本框是否应跳过：已锁定、对象隐藏或所在图层隐藏
function shouldSkipFrame(frame) {
    try {
        if (frame.locked) return true;
        if (frame.visible === false) return true;
        if (frame.itemLayer && frame.itemLayer.visible === false) return true;
    } catch (err) {
        return true; // 属性读取失败时也跳过，避免中断脚本
    }
    return false;
}

// on a given page, loops through all the text frames
// if fitAll is true, refit every text frame to its content
// otherwise only refit the frames that are overset
// if a frame is still overset after fitting, expand it manually
function fitFramesOnPage(page, fitAll) {
    var textFrames = page.textFrames;
    for (var i = 0; i < textFrames.length; i++) {
        var frame = textFrames[i];
        try {
            // 跳过已锁定或隐藏的文本框，避免报错
            if (!shouldSkipFrame(frame) && (frame.overflows || fitAll)) {
                // try using InDesign's default Fit Frame to Content feature
                frame.fit(FitOptions.FRAME_TO_CONTENT);
                // if it doesn't work, expand the frame manually
                if (frame.overflows) {
                    doFit(frame);
                }
            }
        } catch (err) {
            // 单个文本框处理失败时跳过，继续处理其他文本框
        }
    }
}

// recursive function that nudges the frame bounds out by 1
// until the frame is no longer overset
function doFit(frame, attempt) {
    attempt = attempt || 0;
    // 达到次数上限或文本框已锁定/隐藏时停止，避免无限递归报错
    if (!frame.overflows || attempt >= 200 || frame.locked || frame.visible === false) return;
    expandFrame(frame, 5);
    doFit(frame, attempt + 1);
}

// expands a given frame by a given unit
function expandFrame(frame, by) {
    frame.geometricBounds = transformCoords(frame.geometricBounds, [by * -1, by * -1, by, by]);
}

// applies transformations in the format [y1, x1, y2, x2]
function transformCoords(src, trns) {
    return [
        src[0] + trns[0],
        src[1] + trns[1],
        src[2] + trns[2],
        src[3] + trns[3]
    ];
}

try {
    main();
} catch (err) {
    alert(err);
}
