/**
 * 创建InDesign段落样式的脚本
 * 基于"9-默认"创建不同字号的段落样式，拼音字号为主字号的1/3
 */

// 创建基础段落样式"9-默认"
function createDefaultStyle() {
    try {
        var doc = app.activeDocument;
        var defaultStyle = doc.paragraphStyles.add({name: "9-默认"});
        defaultStyle.basedOn = doc.paragraphStyles.item("[基本段落]");
        return defaultStyle;
    } catch(e) {
        alert("创建默认段落样式失败：" + e);
        return null;
    }
}

// 创建指定字号的段落样式
function createSizeStyle(fontSize) {
    try {
        var doc = app.activeDocument;
        var rubySize = Math.round(fontSize / 3 * 10) / 10; // 计算拼音字号并保留一位小数
        var styleName = fontSize.toString();
        
        var newStyle = doc.paragraphStyles.add({name: styleName});
        newStyle.basedOn = doc.paragraphStyles.item("9-默认");
        newStyle.pointSize = fontSize;
        newStyle.rubyFontSize = rubySize;
        
        return newStyle;
    } catch(e) {
        alert("创建字号" + fontSize + "的段落样式失败：" + e);
        return null;
    }
}

// 生成字号序列
function generateSizes() {
    var sizes = [];
    var i;
    
    // 4-12点，步进0.5
    for (i = 4; i <= 12; i += 0.5) {
        sizes.push(i);
    }
    
    // 13-20点，步进1
    for (i = 13; i <= 20; i += 1) {
        sizes.push(i);
    }
    
    // 22-34点，步进2
    for (i = 22; i <= 34; i += 2) {
        sizes.push(i);
    }
    
    // 35-80点，步进5
    for (i = 35; i <= 80; i += 5) {
        sizes.push(i);
    }
    
    return sizes;
}

// 主函数
function main() {
    try {
        if (!app.documents.length) {
            alert("请先打开或创建一个文档！");
            return;
        }

        var defaultStyle = createDefaultStyle();
        if (!defaultStyle) return;

        var sizes = generateSizes();
        for (var i = 0; i < sizes.length; i++) {
            createSizeStyle(sizes[i]);
        }

        alert("段落样式创建完成！");
    } catch(e) {
        alert("脚本执行出错：" + e);
    }
}

// 执行主函数
main();
