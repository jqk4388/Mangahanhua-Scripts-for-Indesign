#include "../Library/KTUlib.jsx"
var defaults = {
    includeAllElements: false,
    oddPages: {
        enabled: true,
        scale: false,
        scaleFactor: 100,
        shiftRight: 0,
        shiftDown: 0
    },
    evenPages: {
        enabled: true,
        scale: false,
        scaleFactor: 100,
        shiftRight: 0,
        shiftDown: 0
    },
    pageRange: 1, // 0 = Range, 1 = All
};

var doc = app.activeDocument;

function myDisplayDialog() {
    myDialog = app.dialogs.add({ name: "图像调整位置_精确数值" });
    var supportedPageRangeTypes = ["范围", "所有"];
    var pageRangeControl, pageRangeInput;
    var includeAllCheckbox;
    var oddPages = {};
    var evenPages = {};
    with(myDialog) {
        with(dialogColumns.add()) {
            includeAllCheckbox = checkboxControls.add({ staticLabel: "同时调整页面上所有元素（排除隐藏和锁定）", checkedState: defaults.includeAllElements });
            oddPages.enabledCheckbox = enablingGroups.add({ staticLabel: "奇数页", checkedState: defaults.oddPages.enabled });
            with(oddPages.enabledCheckbox) {
                with(dialogColumns.add()) {
                    oddPages.scaleCheckbox = checkboxControls.add({ checkedState: defaults.oddPages.scale, staticLabel: "调整图像大小:" });
                    staticTexts.add({ staticLabel: "左右偏移:" });
                    staticTexts.add({ staticLabel: "上下偏移:" });
                }
                with(dialogColumns.add()) {
                    oddPages.scaleFactor = percentEditboxes.add({ editValue: defaults.oddPages.scaleFactor });
                    oddPages.shiftRight = measurementEditboxes.add({ editValue: defaults.oddPages.shiftRight });
                    oddPages.shiftDown = measurementEditboxes.add({ editValue: defaults.oddPages.shiftDown });
                }
            }

            with(borderPanels.add()) {
                staticTexts.add({ staticLabel: "Pages:" });

                pageRangeControl = radiobuttonGroups.add();
                with(pageRangeControl) {
                    for (var i = 0; i < supportedPageRangeTypes.length; i++) {
                        radiobuttonControls.add({ staticLabel: supportedPageRangeTypes[i], checkedState: defaults.pageRange == i, minWidth: 65 });
                    }
                    pageRangeInput = textEditboxes.add({ editContents: app.activeWindow.activePage.name });
                }
            }
        }
        with(dialogColumns.add()) {
            evenPages.enabledCheckbox = enablingGroups.add({ staticLabel: "偶数页", checkedState: defaults.evenPages.enabled });
            with(evenPages.enabledCheckbox) {
                with(dialogColumns.add()) {
                    evenPages.scaleCheckbox = checkboxControls.add({ checkedState: defaults.evenPages.scale, staticLabel: "调整图像大小:" });
                    staticTexts.add({ staticLabel: "左右偏移:" });
                    staticTexts.add({ staticLabel: "上下偏移:" });
                }
                with(dialogColumns.add()) {
                    evenPages.scaleFactor = percentEditboxes.add({ editValue: defaults.evenPages.scaleFactor });
                    evenPages.shiftRight = measurementEditboxes.add({ editValue: defaults.evenPages.shiftRight });
                    evenPages.shiftDown = measurementEditboxes.add({ editValue: defaults.evenPages.shiftDown });
                }
            }
        }
    }
    var myReturn = myDialog.show();
    if (myReturn == true) {
        var needsReview = false;
        var includeAllElements = includeAllCheckbox.checkedState;
        var selectedPageRange = pageRangeControl.selectedButton;
        var pageRange = selectedPageRange === 0 ? getValidRange(pageRangeInput.editContents, oddPages.enabledCheckbox.checkedState, evenPages.enabledCheckbox.checkedState) : doc.pages;
        var regex = /\d{3,4}(?=\_?\d?\d?[a-zA-Z]?\.[A-Za-z]{3,4})/;

        if (selectedPageRange === 0 && pageRange.length === 0) {
            alert('请输入有效的页面范围（例如："12, 32-33"）');
            needsReview = true;
        }

        var transformPage = function(page, data) {
            // if (data.scaleCheckbox.checkedState) {
            //     scalePage(page, parseInt(data.scaleFactor.editValue));
            // };
            // shiftPage(page, parseInt(data.shiftRight.editValue), parseInt(data.shiftDown.editValue));
            // 去掉了取整函数
            if (data.scaleCheckbox.checkedState) {
                scalePage(page, data.scaleFactor.editValue, includeAllElements);
            };
            shiftPage(page, data.shiftRight.editValue, data.shiftDown.editValue, includeAllElements);
        }

        try {
            for (var i = 0; i < pageRange.length; i++) {
                var page = pageRange[i];
                transformPage(page, parseInt(page.name) % 2 > 0 ? oddPages : evenPages);
            }
        } catch (err) { alert(err) }

        if (!needsReview) {
            myDialog.destroy();
        }
    } else {
        myDialog.destroy();
    }
}

// returns sorted array of numbers based on given range string
// e.g. "11-13, 23" returns [11,12,13,23]
function getValidRange(inp, includeOdd, includeEven) {
    var getRange = function(start, stop) {
        var res = []
        for (var i = start; i <= stop; i++) res.push(i); // who needs polyfill when you have for loops amirite
        return res;
    }

    var getCleanRange = function(inp) {
        var nums = inp.split('-');
        var x = parseInt(nums[0]),
            y = parseInt(nums[1]) || parseInt(nums[0]);
        return (nums.length > 0 &&
            x && y &&
            x <= y
        ) ? getRange(x, y) : false;
    }

    var arr = inp.replace(/\s/g, '').split(',');
    var resArr = [];
    for (var i = 0; i < arr.length; i++) {
        var cleanRange = getCleanRange(arr[i]);
        for (var j = 0; cleanRange && j < cleanRange.length; j++) {
            var pageNum = cleanRange[j];
            var pageObj = !!pageNum ? doc.pages.itemByName(pageNum.toString()) : {};
            if (pageObj.isValid &&
                (includeOdd && pageNum % 2 > 0) || (includeEven && pageNum % 2 == 0)) { // exclude odd/even values based on checkbox values
                resArr.push(pageObj);
            }
        }
    }
    return resArr.sort();
}

// collects the items on a page that need to be scaled/shifted
// 默认只处理超出页边距的整页图像；勾选“同时调整页面上所有元素”后，
// 按“选择页面上全部对象”脚本的方式获取页面所有元素（排除隐藏和锁定的元素）
function getTransformTargets(page, includeAllElements) {
    if (!includeAllElements) {
        var graphics = page.allGraphics;
        var res = [];
        for (var i = 0; i < graphics.length; i++) {
            if (isFramePastMargin(page, graphics[i].parent)) {
                res.push({target: graphics[i], isGraphic: true});
            }
        }
        return res;
    }

    var targets = [];
    var pageItems = page.pageItems;
    for (var i = 0; i < pageItems.length; i++) {
        var item = pageItems[i];
        if (item.locked || item.visible === false) continue;
        if (item.itemLayer.locked || !item.itemLayer.visible) continue;
        // 组内嵌套元素随组一起处理，避免重复移动/缩放
        if (item.parent instanceof Group) continue;
        if (item instanceof Group) {
            // 编组作为整体进行移动和缩放
            targets.push({target: item, isGraphic: false});
        } else if (item.allGraphics.length > 0) {
            var itemGraphics = item.allGraphics;
            for (var j = 0; j < itemGraphics.length; j++) {
                targets.push({target: itemGraphics[j], isGraphic: true});
            }
        } else {
            targets.push({target: item, isGraphic: false});
        }
    }
    return targets;
}

// 图像：直接设置“图像相对于图框的比例”（如：图框119.318%、想让图像从30.1875%变到31.52%，填 31.52/119.318*100≈26.42）
// 非图像元素：按图像实际的放大倍数同比缩放。实际倍数 = 放大后有效倍率 / 放大前有效倍率
//          = (图框倍率 × 输入值/100) / (图框倍率 × 原比例/100) = 输入值 / 图像原来的相对比例（图框倍率在计算中相消），
//          而不是直接用输入值作为非图像元素的倍数
function scalePage(page, scaleFactor, includeAllElements) {
    if (!page.isValid || scaleFactor === NaN) return;

    var targets = getTransformTargets(page, includeAllElements);

    // 取页面上最大的图像（整页扫描图）作为倍数基准
    var factor = scaleFactor / 100; // 页面上没有图像时，退回按 输入值/100 处理
    var largestArea = -1;
    for (var i = 0; i < targets.length; i++) {
        if (!targets[i].isGraphic) continue;
        var gr = targets[i].target;
        var b = gr.parent.geometricBounds;
        var area = (b[2] - b[0]) * (b[3] - b[1]);
        var ratio = gr.absoluteHorizontalScale; // 图像当前相对于图框的比例
        if (area > largestArea && ratio > 0) {
            largestArea = area;
            factor = scaleFactor / ratio;
        }
    }

    for (var i = 0; i < targets.length; i++) {
        var t = targets[i].target;
        if (targets[i].isGraphic) {
            t.absoluteHorizontalScale = scaleFactor;
            t.absoluteVerticalScale = scaleFactor;
        } else {
            // 非图像元素：乘以图像实际的放大/缩小倍数
            t.absoluteHorizontalScale = t.absoluteHorizontalScale * factor;
            t.absoluteVerticalScale = t.absoluteVerticalScale * factor;
        }
    }
}

function shiftPage(page, right, down, includeAllElements) {
    if (right === 0 && down === 0) return;

    var targets = getTransformTargets(page, includeAllElements);
    for (var i = 0; i < targets.length; i++) {
        targets[i].target.move(undefined, [right, down]);
    }
}

// returns whether or not a given frame is larger than the page
function isFramePastMargin(page, frame) {
    var pageDims = getBounds(page.bounds),
        frameDims = getBounds(frame.geometricBounds);
    return frameDims.height >= pageDims.height &&
        frameDims.width >= pageDims.width;
}

// converts an array of coordinates into an object in the format {width, height}
// makes math easier to understand :')
function getBounds(bounds) {
    // bounds = x1, y1, x2, y2
    return {
        width: bounds[2] - bounds[0],
        height: bounds[3] - bounds[1]
    }
}


function main() {
    var usersUnits = app.scriptPreferences.measurementUnit; // so we can revert 'em back later
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

    if (doc.allGraphics.length === 0) {
        alert('不能在当前文档中找到任何链接的图像。');
    } else if (app.documents.length != 0) {
        KTUDoScriptAsUndoable(function() {myDisplayDialog()}, "图像调整位置");
    }

    app.scriptPreferences.measurementUnit = usersUnits;
};

try {
    main();
} catch (err) { alert(err) }