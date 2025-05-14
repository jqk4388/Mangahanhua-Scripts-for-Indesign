/* 
放置 Art.js

更新日期：2021年5月3日，Sara Linsley

----------------

安装说明：[https://github.com/saraoswald/Manga-Scripts#how-to-use-scripts-in-indesign](https://github.com/saraoswald/Manga-Scripts#how-to-use-scripts-in-indesign)

使用说明：
- 创建一个具有正确尺寸规格的新文档。
- 运行 Place Art.js。
- 从选择菜单中选择您的图片文件。
- 选择是否允许此脚本从文件名中检测页码，
  或根据提供的起始页按顺序放置它们。
- 选择可选更改：
  - 在放置所有图片时使用“缩放百分比(%)”进行缩放。
    - 这必须是一个整数才能工作（例如 119）。
  - 如果装订设置为“从左到右”，选择是否反向放置图片，
    （例如 001.tif 将被放置在文档的最后一页上）。

重要注意事项：
- 这对计算机来说是一个非常耗资源的操作。您可能需要分批处理。
- 此脚本不会在锁定层上放置图片，因此可以通过锁定图层来保护其他工作。
- 此脚本最初是为放置漫画设计的，与普通英语书籍相比是“反向”的。
  如果希望始终用于非漫画项目，请修改下方的“thisIsManga”变量。
*/

var thisIsManga = true; // 如果不需要反向放置图片，请将此改为 "false"（不加引号）


/* ------ Progress Bar Utility Functions ------ */

// define popup at the top level to please the runtime
var progressBarWindow = new Window("palette", "Placing Art...");
progressBarWindow.minimumSize = { width: 250, height: 50 }; // 尽管文档中怎么说，这实际上不能在初始化时设置！
var progressBar = progressBarWindow.add('progressbar', undefined, 'Progress');

function progressBarStatusText(curr, maxValue, message) {
    return curr.toString() + "/" + maxValue.toString() + ' ' + (message || '');
}

function startProgressBar(maxValue, message) {
    // 显示一个小进度条
    progressBar.maxvalue = maxValue;
    progressBar.minimumSize = { width: 200, height: 10 };
    // 静态文本的长度在初始化后无法编辑，因此尝试模拟字符串将会变多长
    // （progressBar.minWidth 完全不起作用）
    progressBarWindow.add('statictext', undefined, progressBarStatusText(maxValue, maxValue, message));
    progressBarWindow.show();
}

function updateProgressBar(curr, message) {
    // update progress bar values + text
    progressBar.value = curr;
    progressBarWindow.children[1].text = progressBarStatusText(curr, progressBar.maxvalue, message);
    progressBarWindow.update();
}

function destroyProgressBar() {
    progressBarWindow.close();
}

/* ------ Start of Script ------ */

var doc = app.activeDocument;
var bookSize = doc.pages.count();
var isLtR = doc.documentPreferences.pageBinding == PageBindingOptions.LEFT_TO_RIGHT;

var anchorPoints = {
    'Center': AnchorPoint.CENTER_ANCHOR,
    'Top Center': AnchorPoint.TOP_CENTER_ANCHOR,
    'Top Left': AnchorPoint.TOP_LEFT_ANCHOR,
    'Top Right': AnchorPoint.TOP_RIGHT_ANCHOR,
    'Bottom Center': AnchorPoint.BOTTOM_CENTER_ANCHOR,
    'Bottom Left': AnchorPoint.BOTTOM_LEFT_ANCHOR,
    'Bottom Right': AnchorPoint.BOTTOM_RIGHT_ANCHOR,
    'Left Center': AnchorPoint.LEFT_CENTER_ANCHOR,
    'Right Center': AnchorPoint.RIGHT_CENTER_ANCHOR,
}

var anchorPointLabels = [
    'Center',
    'Top Center',
    'Top Left',
    'Top Right',
    'Bottom Center',
    'Bottom Left',
    'Bottom Right',
    'Left Center',
    'Right Center',
]; // ExtendScript doesn't have Object.keys :')

var anchorPointsValues = [
    AnchorPoint.CENTER_ANCHOR,
    AnchorPoint.TOP_CENTER_ANCHOR,
    AnchorPoint.TOP_LEFT_ANCHOR,
    AnchorPoint.TOP_RIGHT_ANCHOR,
    AnchorPoint.BOTTOM_CENTER_ANCHOR,
    AnchorPoint.BOTTOM_LEFT_ANCHOR,
    AnchorPoint.BOTTOM_RIGHT_ANCHOR,
    AnchorPoint.LEFT_CENTER_ANCHOR,
    AnchorPoint.RIGHT_CENTER_ANCHOR,
];

function getAppAnchorPoint() {
    var res = 0;
    for (var i = 0; i < anchorPointsValues.length; i++) {
        if (anchorPointsValues[i] === app.activeWindow.transformReferencePoint)
            res = i;
    }
    return res;
};

var targetLayer = doc.layers.itemByName('Art').isValid ?
    doc.layers.itemByName('Art') :
    doc.layers.add({ name: 'Art' });

function getFilter() {
    if (File.fs == "Windows") {
        return "*.*";
    } else {
        return function() { return true }
    }
}

try {
    var artFiles = File.openDialog('选择图片文件来置入', getFilter(), true);
    if (artFiles !== null) {
        var artFilesToPageNums = [];
        for (var i = 0; i < artFiles.length; i++) {
            var file = artFiles[i];
            artFilesToPageNums.push([file, extractPageNum(file)])
        }
        startDialog(artFilesToPageNums);
    }
} catch (err) { alert(err) }; // the debugger that will never let u down :')

function startDialog(artFilesToPageNums) {
    w = new Window("dialog", "置入图片");
    w.alignChildren = "fill";

    // ---- First row ----

    w.add('group').add('statictext', [0, 0, 400, 50], "此脚本可以根据文件名检测图片应放置的页码，或根据给定的起始页按顺序放置图片。", { multiline: true });

    // ---- Second row ----

    var grp = w.add('group');
    grp.alignChildren = "left";
    // Radio buttons
    var radioGroup = grp.add('panel', undefined, "Place Files Based On:");
    radioGroup.alignChildren = "left";
    var useFileNamesRadio = radioGroup.add('radiobutton', undefined, "File Names");
    var useStartingPageRadio = radioGroup.add('radiobutton', undefined, "Starting Page:");
    // set default radio selection
    useFileNamesRadio.value = true;


    // Starting page text input
    var currentPage = (isLtR && !thisIsManga) || !isLtR ? app.activeWindow.activePage.name : '';
    var startingPageInput = radioGroup.add('edittext', undefined, currentPage);
    startingPageInput.alignment = "fill";

    // List of files and their page numbers
    var fileList = grp.add('listbox', undefined, "", {
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: ["Image Name", "Page Number"]
    });
    fileList.maximumSize = [500, 500];
    fileList.alignment = "fill";

    function fillFileList(array, pageNumFn) {
        fileList.removeAll();
        for (var i = 0; i < array.length; i++) {
            with(fileList.add("item", array[i][0].name)) {
                subItems[0].text = pageNumFn ? pageNumFn(i) : (array[i][1] || '??');
            }
        }
    }
    fillFileList(artFilesToPageNums);

    // ---- Optional Placement Settings ----

    var placementOptionsGroup = w.add('panel', undefined, "Placement Options");
    placementOptionsGroup.alignChildren = "fill";

    // Scale Percentage
    // TODO: some validation on this field?
    var spGrp = placementOptionsGroup.add('group');
    spGrp.add('statictext', undefined, "Scale Percentage (%):");
    var scaleFactorInput = spGrp.add('edittext', undefined, '100');
    scaleFactorInput.characters = 5;


    //  Anchor (Reference) Point
    var apGrp = placementOptionsGroup.add('group');
    apGrp.add('statictext', undefined, "Reference Point:");
    var anchorPointDd = apGrp.add('dropdownlist', undefined, anchorPointLabels);

    anchorPointDd.selection = getAppAnchorPoint();

    // 从左到右复选框行
    // 如果装订是从左到右
    // 检查用户希望以哪个方向放置图片
    var placeBackwardsInput = placementOptionsGroup.add('checkbox', undefined, "Place images in the \"backwards\" manga style");
    placeBackwardsInput.value = thisIsManga;
    if (!isLtR) placeBackwardsInput.hide();


    // ---- Final (Button) row ----

    // OK and Cancel buttons
    var buttonsGroup = w.add('group');
    var okButton = buttonsGroup.add('button', undefined, '置入', { name: 'ok' });
    buttonsGroup.add('button', undefined, '取消', { name: 'cancel' });

    // ---- Validation Functions ----
    function validateNum(inp) {
        return !!inp.text.match(/^-{0,1}\d*\.{0,1}\d+$/);
    }

    function validateStartingPage() {
        // 确保没有非数字值
        // 并且给定的数字小于或等于书籍的最后一页
        return validateNum(startingPageInput) &&
            parseInt(startingPageInput.text) <= parseInt(doc.pages.lastItem().name);
    }

    // ---- Event Handling ----

    useFileNamesRadio.onActivate = function() {
        fillFileList(artFilesToPageNums);
        okButton.enabled = true
    };
    useStartingPageRadio.onActivate = function() {
        var isStartingPageValid = validateStartingPage();
        var startingPage = parseInt(startingPageInput.text);
        if (isStartingPageValid) fillFileList(artFilesToPageNums, function(i) { return i + startingPage });
        okButton.enabled = validateStartingPage()
    };
    startingPageInput.onChanging = function() {
        var isStartingPageValid = validateStartingPage();
        var startingPage = parseInt(startingPageInput.text);
        if (isStartingPageValid) fillFileList(artFilesToPageNums, function(i) { return i + startingPage });
        okButton.enabled = useStartingPageRadio.value === true && validateStartingPage()
    };

    var myReturn = w.show();
    if (myReturn == true) {
        try {
            thisIsManga = placeBackwardsInput.value; // grab the checkbox value and assign it to the global variable
            var startingPage = useStartingPageRadio.value === true ? startingPageInput.text : null;
            placeArtOnAllPages(artFiles, startingPage, {
                scaleFactor: validateNum(scaleFactorInput) ? parseFloat(scaleFactorInput.text) : 100,
                anchorPoint: anchorPoints[anchorPointDd.selection]
            });
        } catch (err) { alert(err) }
    }
}

function placeArtOnAllPages(artFiles, startingPage, options) {
    var hasErrors = false,
        pagesCount = 0; // for debugging :')
    if (artFiles !== null && artFiles.length > 0) { // in case the user pressed cancel or something
        startProgressBar(artFiles.length, artFiles[0].name);
        var pageNum = startingPage; // null if the user wants the number determined from the file name
        for (var i = 0; !hasErrors && i < artFiles.length; i++) {
            // actually place the art
            hasErrors = placeArtOnPage(artFiles[i], pageNum, options);
            if (pageNum !== null) pageNum++;

            if (!hasErrors) pagesCount++; // keep track of how many images succeeded

            updateProgressBar(i + 1, artFiles[i].name);
        }
        destroyProgressBar();
    }
    return hasErrors, pagesCount;
}

function placeArtOnPage(artLink, pageNum, options) {
    if (!pageNum) pageNum = extractPageNum(artLink)
    var hasErrors = pageNum < 1;
    if (!hasErrors) {
        var image = new File(artLink);
        // determine INDD page number based on book binding
        var bookPageNum = thisIsManga && isLtR ? bookSize - pageNum + 1 : pageNum; // LtR books are "backwards", where "Page 1" the last page in the document
        var page = doc.pages.itemByName(bookPageNum.toString()); // which page in the INDD to place the art on

        hasErrors = prePlaceErrorHandling(page, bookPageNum, bookSize, image, artLink);

        if (!hasErrors) {
            var frameWithArt = placeArtOnPageHelper(page, targetLayer, image);
            if (frameWithArt) {
                // if successfully placed art, apply options
                // just assume there's only one placed grapic per page
                if (options.scaleFactor && frameWithArt.graphics && frameWithArt.graphics[0].isValid) {
                    scalePage(frameWithArt.graphics[0], options);
                }
            }
        }
    }
    return hasErrors;
}

function placeArtOnPageHelper(page, layer, image) {
    var frame = getFrame(page, layer);
    if (frame) frame.place(image);
    return frame
}

function prePlaceErrorHandling(page, pageNum, image, imageLink) {
    var debugInfo = "\n\n----Debug Info Below----\nPage number the script calculated from the link: " + pageNum.toString() + "\nNumber of pages in book: " + bookSize.toString() + "\nImage that messed up: " + imageLink;

    if (!page.isValid || pageNum === 0) {
        alert("Failed to match the number in the file name to one in the book.\nTry renaming your files to look like '123.TIF'." + debugInfo);
        return true;
    }
    if (image.error) {
        alert("Something wrong with the Image\nReview the image link below. Not sure when this would happen, but InDesign couldn't create a File out of the image's link" + debugInfo)
        return true;
    }

    return false; // yay no errors!!
}

function getFrame(page, layer) {
    var targetFrame = null;

    var findFrame = function(arr, ignoreLayer) {
        var returnVal = null;
        // iterate backwards because it's most likely on the bottom
        for (var i = arr.length; i > -1; i-- && returnVal === null) {
            // only grab an object if it's a Rectangle on a non-locked layer
            if (arr[i] && arr[i] instanceof Rectangle &&
                !arr[i].itemLayer.locked &&
                (ignoreLayer || arr[i].itemLayer == layer)) { // if using an element from a master, don't be picky about the layer
                returnVal = arr[i];
            }
        }
        return returnVal;
    }

    // check if there's a master, override the master items and use the frame there 
    if (page.appliedMaster !== null && page.appliedMaster.isValid) {
        var masterFrame = findFrame(page.masterPageItems, true); // 查找一个图框，但不挑剔它所在的图层
        targetFrame = masterFrame && masterFrame.isValid ?
            masterFrame.override(page) : null;
    };
        // 检查 Art 图层底部是否有矩形
    if (targetFrame === null) targetFrame = findFrame(page.allPageItems);
        // 如果没有图框，也没有主页图框，创建一个延伸到页面边缘的图框
    if (targetFrame === null) targetFrame = page.rectangles
        .add({
            geometricBounds: getPageBounds(page),
            itemLayer: targetLayer
        });

    return targetFrame;
}

function extractPageNum(path) {
    var regex = /\d{3,4}(?=\_?\d?[a-zA-Z]?\.[A-Za-z]{3,4})/;
    var regexResult = regex.exec(path);
    return path && regexResult !== null && regexResult.length > 0 ?
        parseInt(regexResult[0], 10) : 0;
}

// 接受边界输入，并以 [y1, x1, y2, x2] 的形式给出输出
// 添加文档的出血设置，同时考虑页面内侧空白区域
function getPageBounds(page) {
    var prfs = doc.documentPreferences,
        pb = page.bounds,
        bleedLeft = page.side == PageSideOptions.RIGHT_HAND ? 0 : prfs.documentBleedOutsideOrRightOffset,
        bleedRight = page.side == PageSideOptions.LEFT_HAND ? 0 : prfs.documentBleedOutsideOrRightOffset;
    var y1 = pb[0] - prfs.documentBleedTopOffset,
        x1 = pb[1] - bleedLeft,
        y2 = pb[2] + prfs.documentBleedBottomOffset,
        x2 = pb[3] + bleedRight;
    return [y1, x1, y2, x2];
}


function scalePage(graphic, options) {
    if (options.scaleFactor === 100) return; // 如果缩放比例未改变，则不浪费资源
    var scaleMatrix = app.transformationMatrices.add({ horizontalScaleFactor: options.scaleFactor / 100, verticalScaleFactor: options.scaleFactor / 100 });
    graphic.transform(CoordinateSpaces.INNER_COORDINATES, options.anchorPoint, scaleMatrix);
}