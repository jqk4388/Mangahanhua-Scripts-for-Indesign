var docs = app.documents;
var i, doc;

// 获取用户输入
var width = parseFloat(prompt("请输入页面宽度（毫米）：", "130"));
if (isNaN(width) || width <= 0) width = 130;
var height = parseFloat(prompt("请输入页面高度（毫米）：", "185"));
if (isNaN(height) || height <= 0) height = 185;

for (i = 0; i < docs.length; i++) {
    doc = docs[i];
    // 设置文档单位
    doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
    doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
    doc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
    doc.viewPreferences.typographicMeasurementUnits = MeasurementUnits.POINTS;
    // 设置页面尺寸
    doc.documentPreferences.pageWidth = width;
    doc.documentPreferences.pageHeight = height;
}
alert("所有文档页面尺寸已设置为" + width + "mm x " + height + "mm，单位已调整。");