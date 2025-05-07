var theDoc = app.activeDocument;

var matchSelectionToBleed = function () {
    var frames = [];
    var theLayers = theDoc.layers; // 获取文档的所有图层
    if (app.selection.length === 0) {
        if (!confirm("没有选中图框，将对文档中所有图框进行操作，是否继续？")) {
            return;
        }
        // 获取文档中所有图框
        for (var i = 0; i < theDoc.allPageItems.length; i++) {
            var item = theDoc.allPageItems[i];
            if (item.constructor.name === "Rectangle" || item.constructor.name === "Polygon" || item.constructor.name === "Oval") {
                frames.push(item);
            }
        }
    } else {
        for (var i = 0; i < app.selection.length; i++) {
            frames.push(app.selection[i]);
        }
    }
    for (var j = 0; j < frames.length; j++) {
        if (frames[j].itemLayer.id == theLayers[theLayers.length - 1].id) { // 只处理底层（画层）上的对象
            KTUMatchFrameToInfoArea(frames[j]);
        }
    }
}

matchSelectionToBleed();

function KTUMatchFrameToInfoArea(theFrame) {
    // set the measurement units to Points, so our math lower down will work out
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
    // set the ruler to "spread", again so our math works out.
    var oldOrigin = theDoc.viewPreferences.rulerOrigin; // save old ruler origin
    theDoc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;

    // 假设辅助信息区的线在页面边界外扩一定距离
    var infoAreaOffset = 28.34; //10mm

    if (theDoc.documentPreferences.pageBinding == PageBindingOptions.leftToRight) { // LTR
        if (theFrame.parentPage.index % 2 == 0) { // 左页
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - infoAreaOffset,
                theFrame.parentPage.bounds[1] - infoAreaOffset,
                theFrame.parentPage.bounds[2] + infoAreaOffset,
                theFrame.parentPage.bounds[3] + 0
            ];
        } else { // 右页
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - infoAreaOffset,
                theFrame.parentPage.bounds[1] - 0,
                theFrame.parentPage.bounds[2] + infoAreaOffset,
                theFrame.parentPage.bounds[3] + infoAreaOffset
            ];
        }
    } else { // RTL
        if (theFrame.parentPage.index % 2 == 0) { // 右页
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - infoAreaOffset,
                theFrame.parentPage.bounds[1] + 0,
                theFrame.parentPage.bounds[2] + infoAreaOffset,
                theFrame.parentPage.bounds[3] + infoAreaOffset
            ];
        } else { // 左页
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - infoAreaOffset,
                theFrame.parentPage.bounds[1] - infoAreaOffset,
                theFrame.parentPage.bounds[2] + infoAreaOffset,
                theFrame.parentPage.bounds[3] + 0
            ];
        }
    }
    // 恢复原有设置
    theDoc.viewPreferences.rulerOrigin = oldOrigin;
    app.scriptPreferences.measurementUnit = AutoEnum.AUTO_VALUE;
    return theFrame;
}