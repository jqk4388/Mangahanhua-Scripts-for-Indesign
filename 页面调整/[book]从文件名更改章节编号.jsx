var doc = app.activeDocument;
var docName = doc.name;
var chapterNum = docName.match(/\d+/);

if (chapterNum) {
    chapterNum = chapterNum[0];
    var num = parseInt(chapterNum, 10);
} else {
    alert("无法从文件名「" + docName + "」中获取章节编号，请检查文件名是否包含数字");
    exit();
}

app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;

// var masterD = doc.masterSpreads.itemByName("D-文字封面");
// var masterC = doc.masterSpreads.itemByName("C-版权页");

// if (!masterD.isValid) {
//     alert("找不到主页 D-Master");
//     exit();
// }

// if (!masterC.isValid) {
//     alert("找不到主页 C-Master");
//     exit();
// }

// var firstPage = doc.pages[0];

// for (var i = 0; i < 1; i++) {
//     var newPage = doc.pages.add(LocationOptions.BEFORE, firstPage);
//     newPage.appliedMaster = masterD;
// }
var cnp = doc.chapterNumberPreferences;
cnp.chapterNumber = num;
alert("章节编号已设置为 " + num);

// var lastPage = doc.pages[-1];
// var endPage = doc.pages.add(LocationOptions.AFTER, lastPage);
// endPage.appliedMaster = masterC;

