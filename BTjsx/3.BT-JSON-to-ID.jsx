// InDesign JSON文本导入工具
// 该脚本用于处理JSON格式的翻译稿，将文本插入到InDesign文档中
// 支持从BallonsTranslator-JSON文件中读取图片列表和文本框数据，实现自动排版
// 作者：几千块
// 日期：20260201

// 声明全局变量
#include "../Library/KTUlib.jsx"
var version = "1.0";
var doc = app.activeDocument;
var selectedImages = []; // 保存选中的图片
var jsonData = {}; // 保存解析后的JSON数据
var imageList = []; // 保存图片文件列表
var pageOffset = 0; // 页码偏移，允许正/负偏移，默认0
var imageDirectory = ""; // 保存图片所在目录路径
var imageDimensions = {}; // 保存图片尺寸信息 {imageName: {width: w, height: h}}

// 主函数：显示用户界面
function showMainInterface() {
    var dialog = new Window("dialog", "JSON文本导入工具 " + version);
    dialog.preferredSize = [700, 450];
    
    // 顶部：文件选择部分
    var topGroup = dialog.add("group");
    topGroup.orientation = "row";
    topGroup.alignChildren = "center";
    topGroup.margins = [10, 10, 10, 5];
    topGroup.add("statictext", undefined, "选择JSON文件:");
    var filePathInput = topGroup.add("edittext", undefined, "");
    filePathInput.characters = 40;
    var browseButton = topGroup.add("button", undefined, "浏览");
    
    // 浏览按钮点击事件
    browseButton.onClick = function () {
        var jsonFile = File.openDialog("请选择一个JSON文件", "*.json");
        if (jsonFile) {
            filePathInput.text = jsonFile.fsName;
            // 选择文件后自动加载并刷新列表
            try {
                loadAndParseJSON(filePathInput.text);
                fillImageList();
            } catch (e) {
                alert("加载JSON文件失败: " + e.message);
            }
        }
    };
    
    // 中部：图片列表和配置选项
    var middleGroup = dialog.add("group");
    middleGroup.orientation = "row";
    middleGroup.alignChildren = "fill";
    middleGroup.margins = [10, 5, 10, 10];
    middleGroup.spacing = 10;
    
    // 左侧：图片列表
    var leftPanel = middleGroup.add("panel", undefined, "图片文件列表");
    leftPanel.alignChildren = "fill";
    leftPanel.preferredSize = [400, 350];
    
    var fileList = leftPanel.add("listbox", undefined, undefined, {
        numberOfColumns: 1,
        showHeaders: true,
        columnTitles: ["文件名"],
        multiselect: true
    });
    fileList.maximumSize.height = 300;
    fileList.preferredSize = [380, 300];
    fileList.alignment = "fill";
    
    // 右侧：配置选项
    var rightPanel = middleGroup.add("panel", undefined, "配置选项");
    rightPanel.alignChildren = "left";
    rightPanel.preferredSize = [250, 350];
    rightPanel.spacing = 10;
    
    // 页码偏移设置
    rightPanel.add("statictext", undefined, "页码偏移:");
    var offsetGroup = rightPanel.add("group");
    offsetGroup.orientation = "row";
    offsetGroup.alignChildren = "center";
    var offsetInput = offsetGroup.add("edittext", undefined, "0");
    offsetInput.characters = 4;
    offsetGroup.add("statictext", undefined, "(可为负数)");
    
    // 文本框选项
        rightPanel.add("statictext", undefined, "文本框选项:");
        var fitToContentCheckbox = rightPanel.add("checkbox", undefined, "框架大小适合文本");
        fitToContentCheckbox.value = true;
        
        var applyStylesCheckbox = rightPanel.add("checkbox", undefined, "根据json中属性创建样式");
        applyStylesCheckbox.value = true;
        
        var fitTextToFrameCheckbox = rightPanel.add("checkbox", undefined, "文本大小适合框架");
        fitTextToFrameCheckbox.value = false;
        
        var includeFontInfoCheckbox = rightPanel.add("checkbox", undefined, "导入带字体字号信息");
        includeFontInfoCheckbox.value = false;
        
        // 添加事件监听器，确保两个冲突选项不能同时选中
        applyStylesCheckbox.onClick = function() {
            if (this.value) {
                includeFontInfoCheckbox.value = false;
            }
        };
        
        includeFontInfoCheckbox.onClick = function() {
            if (this.value) {
                applyStylesCheckbox.value = false;
            }
        };
    
    // 基准字号设置
    rightPanel.add("statictext", undefined, "基准字号:");
    var baseFontSizeGroup = rightPanel.add("group");
    baseFontSizeGroup.orientation = "row";
    baseFontSizeGroup.alignChildren = "center";
    var baseFontSizeInput = baseFontSizeGroup.add("edittext", undefined, "26");
    baseFontSizeInput.characters = 4;
    baseFontSizeGroup.add("statictext", undefined, "(默认26点=9点)");
    
    // 图片置入选项
    rightPanel.add("statictext", undefined, "图片置入选项:");
    var placeImagesCheckbox = rightPanel.add("checkbox", undefined, "是否从json中置入图片");
    placeImagesCheckbox.value = false;
    
    // 底部：确定和取消按钮
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignChildren = "center";
    buttonGroup.margins = [10, 5, 10, 10];
    buttonGroup.alignment = "right";
    var cancelButton = buttonGroup.add("button", undefined, "取消");
    var confirmButton = buttonGroup.add("button", undefined, "确定");
    
    // 取消按钮点击事件
    cancelButton.onClick = function () {
        dialog.close(0);
    };
    
    // 确定按钮点击事件
    confirmButton.onClick = function () {
        // 如果未选择文件或列表为空，提示并返回
        if (!filePathInput.text || fileList.items.length === 0) {
            alert('请先选择一个 JSON 文件。');
            return;
        }
        
        // 获取用户选择的图片列表
        selectedImages = [];
        for (var i = 0; i < fileList.items.length; i++) {
            if (fileList.items[i].selected) {
                selectedImages.push(fileList.items[i].text);
            }
        }
        
        // 读取并解析页码偏移值
        pageOffset = parseInt(offsetInput.text, 10);
        if (isNaN(pageOffset)) {
            pageOffset = 0;
        }
        
        // 保存配置选项
        var baseFontSize = parseInt(baseFontSizeInput.text);
        if (isNaN(baseFontSize) || baseFontSize <= 0) {
            baseFontSize = 26; // 默认值
        }
        
        var config = {
            fitToContent: fitToContentCheckbox.value,
            applyStyles: applyStylesCheckbox.value,
            placeImages: placeImagesCheckbox.value,
            baseFontSize: baseFontSize,
            fitTextToFrame: fitTextToFrameCheckbox.value,
            includeFontInfo: includeFontInfoCheckbox.value
        };
        
        dialog.close(1);
        
        // 执行导入操作
        try {
            // 设置标尺原点和单位
            doc.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;
            doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
            doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
            doc.zeroPoint = [0, 0];
            try {
                var textLayer = doc.layers.itemByName("Text");
                if (textLayer.isValid) {
                    doc.activeLayer = textLayer;
                    textLayer.locked = false;
                    textLayer.visible = true;
            }
            } catch (e) {}
            KTUDoScriptAsUndoable(function() {processSelectedImages(selectedImages, config)}, "BT-JSON导入");
            alert("导入完成！");
        } catch (e) {
            alert("导入失败: " + e.message);
        }
    };
    
    // 填充图片列表函数
    function fillImageList() {
        fileList.removeAll();
        if (imageList.length === 0) {
            alert("JSON文件中未找到图片数据。");
            return;
        }
        
        for (var i = 0; i < imageList.length; i++) {
            var item = fileList.add("item", imageList[i]);
            item.selected = true; // 默认全选
        }
        
    }
    
    // 显示对话框
    dialog.center();
    var result = dialog.show();
}

// 读取并解析JSON文件
function loadAndParseJSON(filePath) {
    var file = new File(filePath);
    if (!file.exists) {
        throw new Error("文件不存在");
    }
    
    file.open("r", "UTF-8");
    var content = file.read();
    file.close();
    
    try {
        jsonData = eval('(' + content + ')');
        // 提取图片目录路径
        imageDirectory = jsonData.directory || "";
        // 提取图片文件列表
        imageList = [];
        if (jsonData.pages) {
            for (var imageName in jsonData.pages) {
                if (jsonData.pages.hasOwnProperty(imageName)) {
                    imageList.push(imageName);
                }
            }
        }
    } catch (e) {
        throw new Error("解析JSON文件失败: " + e.message);
    }
}

// 处理选中的图片
function processSelectedImages(images, config) {
    // 确保页面数足够
    ensureEnoughPages(images.length+ pageOffset);
    
    // 创建进度窗口
    var progressWindow = new Window("palette", "处理中...");
    progressWindow.preferredSize = [300, 80];
    var progressBar = progressWindow.add("progressbar", [10, 20, 290, 40], 0, images.length);
    var progressText = progressWindow.add("statictext", [10, 45, 290, 65], "准备中...");
    progressWindow.show();
    
    try {
        // 遍历每个选中的图片
        for (var i = 0; i < images.length; i++) {
            var imageName = images[i];
            var pageIndex = i + pageOffset;
            
            // 更新进度
            progressBar.value = i + 1;
            progressText.text = "处理第 " + (i + 1) + " 张图片: " + imageName;
            progressWindow.update();
            
            // 确保页面存在
            var page = doc.pages[pageIndex];
            if (!page) {
                page = doc.pages.add();
            }
            
            // 如果配置了置入图片，则先置入图片
            if (config && config.placeImages) {
                processImagePlacement(imageName, page);
            }
            
            // 处理图片对应的文本框
            processImageTextFrames(imageName, page, config);
        }
    } finally {
        // 关闭进度窗口
        progressWindow.close();
    }
}

// 处理图片对应的文本框
function processImageTextFrames(imageName, page, config) {
    if (!jsonData.pages || !jsonData.pages[imageName]) {
        return;
    }
    
    // 获取图片尺寸（从JSON最外层的image_info字段中提取）
    var imageDim;
    if (imageDimensions[imageName]) {
        imageDim = imageDimensions[imageName];
    } else {
        // 从最外层的image_info字段中提取
        var imageInfo = null;
        if (jsonData.image_info && jsonData.image_info[imageName]) {
            imageInfo = jsonData.image_info[imageName];
        }
        
        // 从image_info中提取尺寸
        if (imageInfo && imageInfo.width && imageInfo.height) {
            imageDim = { width: imageInfo.width, height: imageInfo.height };
        } else {
            // 默认尺寸
            imageDim = { width: 800, height: 1000 };
        }
        
        imageDimensions[imageName] = imageDim;
    }
    
    var textFrames = jsonData.pages[imageName];
    for (var i = 0; i < textFrames.length; i++) {
        var textFrameData = textFrames[i];
        try {
            createTextFrameFromData(textFrameData, page, config, imageDim);
        } catch (e) {
            // 忽略单个文本框创建失败的错误，继续处理其他文本框
            $.writeln("创建文本框失败: " + e.message);
        }
    }
}

// 从数据创建文本框
function createTextFrameFromData(data, page, config, imageDim) {
    // 提取文本框数据
    var xyxy = data.xyxy;
    var translation = data.translation || "";
    var angle = data.angle || 0;
    var vec = data.fontformat.vertical || false;
    var fontSize = data.fontformat ? data.fontformat.font_size : 12;
    var fontFamily = data.fontformat ? data.fontformat.font_family : "Arial";
    var detected_font_name = data._detected_font_name|| fontFamily;
    
    // 转换坐标为百分比
    var percentCoords = convertXYXYtoPercent(xyxy, imageDim.width, imageDim.height);
    
    // 转换百分比坐标为InDesign绝对坐标
    var pageWidth = page.bounds[3] - page.bounds[1];
    var pageHeight = page.bounds[2] - page.bounds[0];
    var absCoords = convertPercentToAbsoluteCoordinates(pageWidth,pageHeight, percentCoords);
    
    // 创建文本框
    var textFrame = page.textFrames.add();
    textFrame.geometricBounds = [absCoords.top, absCoords.left, absCoords.bottom, absCoords.right];
    
    // 如果需要包含字体字号信息
    var textContent = translation;
    if (config && config.includeFontInfo) {
        textContent = "{字体：" + detected_font_name + "}{字号：" + fontSize + "}" + translation;
    }
    textFrame.contents = textContent;
        // 应用文本方向
    if (vec) {
        textFrame.parentStory.storyPreferences.storyOrientation = StoryHorizontalOrVertical.VERTICAL;
    }
    
    // 应用文本属性
    if (config && config.applyStyles) {
        applyTextProperties(textFrame, fontFamily, detected_font_name, fontSize, angle, config);
    }

    // 自动调整文本框大小
    if (textFrame.overflows && (!config || !config.fitTextToFrame)) {
            textFrame.fit(FitOptions.FRAME_TO_CONTENT);
        }

    // 文本大小适合文本框
    if (config && config.fitTextToFrame) {
        try {
            TextFitter.processTextFrame(textFrame);
        } catch (e) {
            // 忽略错误，继续处理其他文本框
            $.writeln("调整文本大小失败: " + e.message);
        }
    }
}

// 应用文本属性
function applyTextProperties(textFrame, fontFamily, detected_font_name, fontSize, angle, config) {
    // 根据基准字号计算实际应用的字号，保留整数
    var actualFontSize = fontSize;
    if (config && config.baseFontSize && config.baseFontSize > 0) {
        actualFontSize = (fontSize / config.baseFontSize) * 9;
    }
    actualFontSize = Math.round(actualFontSize); // 保留整数
    
    // 应用字符样式
    applyCharacterStyle(textFrame, fontFamily, detected_font_name, config);
    
    // 应用段落样式
    applyParagraphStyle(textFrame, actualFontSize, config);
    
    // 应用旋转角度
    if (angle !== 0) {
        textFrame.rotationAngle = -angle;
    }
}

// 应用或创建字符样式
function applyCharacterStyle(textFrame, fontFamily, detected_font_name, config) {
    try {
        var allCharacterStyles = doc.characterStyles;
        var matchedStyle = null;
        
        // 遍历所有字符样式，进行模糊匹配
        for (var i = 0; i < allCharacterStyles.length; i++) {
            var style = allCharacterStyles[i];
            if (style.name.toLowerCase().indexOf(detected_font_name.toLowerCase()) !== -1) {
                matchedStyle = style;
                break;
            }
        }
        
        // 如果匹配到样式，直接应用
        if (matchedStyle) {
            textFrame.paragraphs[0].appliedCharacterStyle = matchedStyle;
        } else {
            // 没有匹配到，新建字符样式
            var newStyle = createCharacterStyleByName(detected_font_name);
            if (newStyle) {
                textFrame.paragraphs[0].appliedCharacterStyle = newStyle;
            } else {
                // 如果新建失败，尝试直接应用字体
                try {
                    textFrame.paragraphs[0].appliedFont = fontFamily;
                } catch (e) {
                    // 忽略错误
                }
            }
        }
    } catch (e) {
        // 忽略错误
    }
}

// 应用或创建段落样式
function applyParagraphStyle(textFrame, fontSize, config) {
    try {
        var allParagraphStyles = doc.paragraphStyles;
        var matchedStyle = null;
        var fontSizeStr = fontSize.toString();
        
        // 遍历所有段落样式，匹配字号
        for (var i = 0; i < allParagraphStyles.length; i++) {
            var style = allParagraphStyles[i];
            if (style.name === fontSizeStr) {
                matchedStyle = style;
                break;
            }
        }
        
        // 如果匹配到样式，直接应用
        if (matchedStyle) {
            textFrame.paragraphs[0].appliedParagraphStyle = matchedStyle;
        } else {
            // 没有匹配到，新建段落样式
            var newStyle = createParagraphStyleBySize(fontSize);
            if (newStyle) {
                textFrame.paragraphs[0].appliedParagraphStyle = newStyle;
            }
        }
    } catch (e) {
        // 忽略错误
    }
}

// 根据字体名创建字符样式
function createCharacterStyleByName(fontName) {
    try {
        var baseStyle = doc.characterStyles.itemByName("[无]");
        var newStyle = doc.characterStyles.add({ name: fontName, basedOn: baseStyle });
        newStyle.appliedFont = fontName;
        return newStyle;
    } catch (e) {
        // 样式已存在或其他错误，返回null
        return null;
    }
}

function createDefaultStyle() {
    try {
        var defaultStyle = doc.paragraphStyles.add({name: "~脚本基础~"});
        defaultStyle.basedOn = doc.paragraphStyles.item("[基本段落]");
        return defaultStyle;
    } catch(e) {
        // alert("创建默认段落样式失败：" + e);
        return null;
    }
}

// 根据字号创建段落样式
function createParagraphStyleBySize(fontSize) {
    try {
        var styleName = fontSize.toString();
        if (!doc.paragraphStyles.itemByName("~脚本基础~").isValid) {
            createDefaultStyle();
        }
        var baseParaStyle = doc.paragraphStyles.itemByName("~脚本基础~");
        var newStyle = doc.paragraphStyles.add({ name: styleName, basedOn: baseParaStyle });        
        newStyle.pointSize = fontSize;
        return newStyle;
    } catch (e) {
        // 样式已存在或其他错误，返回null
        return null;
    }
}

// 确保页面数足够
function ensureEnoughPages(requiredPageCount) {
    var currentPageCount = doc.pages.length;
    
    if (currentPageCount < requiredPageCount) {
        var pagesToAdd = requiredPageCount - currentPageCount;
        for (var i = 0; i < pagesToAdd; i++) {
            doc.pages.add();
        }
    }
}

// 构建图片完整路径
function buildImagePath(imageName) {
    if (!imageDirectory) {
        return imageName;
    }
    // 确保路径格式正确，处理不同操作系统的路径分隔符
    var path = imageDirectory;
    var separator = File.fs === "Windows" ? "\\" : "/";
    if (path.charAt(path.length - 1) !== '/' && path.charAt(path.length - 1) !== '\\') {
        path += separator;
    }
    // 替换路径中的分隔符为当前系统的分隔符
    path = path.replace(/[\/\\]/g, separator);
    return path + imageName;
}

// 将JSON中的xyxy坐标转换为百分比坐标
function convertXYXYtoPercent(xyxy, imageWidth, imageHeight) {
    // xyxy格式: [left, top, right, bottom]
    var left = xyxy[0] / imageWidth;
    var top = xyxy[1] / imageHeight;
    var right = xyxy[2] / imageWidth;
    var bottom = xyxy[3] / imageHeight;
    
    return { left: left, top: top, right: right, bottom: bottom };
}

// 将百分比坐标转换为InDesign绝对坐标
function convertPercentToAbsoluteCoordinates(pageWidth,pageHeight, percentCoords) {
    var left = pageWidth * percentCoords.left;
    var top = pageHeight * percentCoords.top;
    var right = pageWidth * percentCoords.right;
    var bottom = pageHeight * percentCoords.bottom;
    
    return { left: left, top: top, right: right, bottom: bottom };
}

// 获取或创建图片框
function getImageFrame(page) {
    var targetLayer = doc.layers.itemByName('Art').isValid ? 
        doc.layers.itemByName('Art') : 
        doc.layers.add({ name: 'Art' });
    
    var targetFrame = null;
    
    // 查找合适的矩形框的内部函数
    var findFrame = function(arr, ignoreLayer) {
        var returnVal = null;
        // 从后向前遍历，因为框架通常在底层
        for (var i = arr.length - 1; i >= 0 && returnVal === null; i--) {
            var item = arr[i];
            // 只选择非锁定图层上的矩形
            if (item && item instanceof Rectangle && !item.itemLayer.locked && 
                (ignoreLayer || item.itemLayer === targetLayer)) {
                returnVal = item;
            }
        }
        return returnVal;
    };
    
    // 检查页面是否应用了主页，如果有则覆盖主页框架
    if (page.appliedMaster !== null && page.appliedMaster.isValid) {
        var masterFrame = findFrame(page.masterPageItems, true); // 查找主页框架，不挑剔图层
        targetFrame = masterFrame && masterFrame.isValid ? 
            masterFrame.override(page) : null;
    }
    
    // 如果没有找到主页框架，检查页面上的项目
    if (targetFrame === null) {
        targetFrame = findFrame(page.allPageItems);
    }
    
    // 如果仍然没有找到，创建一个新的框架
    if (targetFrame === null) {
        var pageBounds = page.bounds;
        targetFrame = page.rectangles.add({
            geometricBounds: [pageBounds[0], pageBounds[1], pageBounds[2], pageBounds[3]],
            itemLayer: targetLayer
        });
    }
    
    return targetFrame;
}

// 置入图片到页面
function placeImageOnPage(imagePath, page) {
    var imageFile = new File(imagePath);
    if (!imageFile.exists) {
        throw new Error("图片文件不存在: " + imagePath);
    }
    
    var frame = getImageFrame(page);
    if (frame) {
        frame.place(imageFile);
        // 调整图片大小以适应框
        if (frame.graphics && frame.graphics.length > 0) {
            var graphic = frame.graphics[0];
            graphic.fit(FitOptions.FRAME_TO_CONTENT);
            frame.fit(FitOptions.CONTENT_TO_FRAME);
        }
        return true;
    }
    return false;
}

// 处理图片置入
function processImagePlacement(imageName, page) {
    try {
        var imagePath = buildImagePath(imageName);
        return placeImageOnPage(imagePath, page);
    } catch (e) {
        alert("置入图片失败: " + e.message);
        return false;
    }
}

// TextFitter对象：用于调整文本大小以适应文本框
var TextFitter = {
    // 默认配置参数
    config: {
        minFontSize: 6,
        maxFontSize: 30,
        lineHeightAdjustRange: 0.2, // ±20%
        trackingAdjustRange: 0.1,   // ±10%
        maxFontSizeForShortText: 36 // 短文本最大字号限制
    },
    
    // 处理单个文本框
    processTextFrame: function(textFrame) {
        // 检查文本框是否为空
        if (textFrame.contents === "" || textFrame.contents === null) {
            throw new Error("文本框为空");
        }
        
        // 获取文本框尺寸
        var frameWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];
        var frameHeight = textFrame.geometricBounds[2] - textFrame.geometricBounds[0];
        
        // 检查文本框是否过小
        if (frameWidth <= 0 || frameHeight <= 0) {
            throw new Error("文本框尺寸无效");
        }
        
        // 获取文本内容
        var textContent = textFrame['parentStory']['contents'];
        var textLength = textContent.length;
        
        // 根据文本长度调整最大字号限制
        var maxFontSize = this.config.maxFontSize;
        if (textLength <= 2) {
            maxFontSize = this.config.maxFontSizeForShortText;
        }
        
        // 获取当前文本属性
        var currentFontSize = this.getFontSize(textFrame);
        var currentLeading = this.getLeading(textFrame);
        var currentTracking = this.getTracking(textFrame);
        
        // 判断是否需要调整
        var overflow = textFrame.overflows;
        
        // 检查文本是否合适
        var fitStatus = this.isTextFitWell(textFrame);
        
        // 如果文本正好合适，不需要调整
        if (!overflow && fitStatus.fitWell) {
            return;
        }
        
        // 执行调整
        this.adjustTextProperties(textFrame, {
            minFontSize: this.config.minFontSize,
            maxFontSize: maxFontSize,
            lineHeightAdjustRange: this.config.lineHeightAdjustRange,
            trackingAdjustRange: this.config.trackingAdjustRange,
            fitStatus: fitStatus
        });
    },
    
    // 获取字体大小
    getFontSize: function(textFrame) {
        try {
            return textFrame['parentStory']['pointSize'];
        } catch (e) {
            return 12; // 默认值
        }
    },
    
    // 获取行距
    getLeading: function(textFrame) {
        try {
            var leading = textFrame['parentStory']['leading'];
            return (leading == Leading.AUTO) ? this.getFontSize(textFrame) * 1.2 : leading;
        } catch (e) {
            return this.getFontSize(textFrame) * 1.2; // 默认值
        }
    },
    
    // 获取字间距
    getTracking: function(textFrame) {
        try {
            return textFrame['parentStory']['tracking'];
        } catch (e) {
            return 0; // 默认值
        }
    },
    
    // 检查文本是否已经合适
    isTextFitWell: function(textFrame) {
        var result = {
            fitWell: true,
            overflow: false,
            tooMuchWhitespace: false,
            whitespaceRatio: 0
        };
        
        // 检查是否有溢出
        result.overflow = textFrame.overflows;
        if (result.overflow) {
            result.fitWell = false;
            return result;
        }
        
        // 检查最后一行是否接近底部（避免过多留白）
        try {
            var lines = textFrame.lines;
            if (lines.length > 0) {
                var lastLine = lines[lines.length - 1];
                var lastLineBounds = lastLine.baseline + lastLine.descent;
                var frameBounds = textFrame.geometricBounds;
                
                // 计算最后一行底部与文本框底部的距离                
                var frameHeight = frameBounds[2] - frameBounds[0];
                var distanceToBottom = frameHeight - lastLineBounds;
                
                // 计算留白比例
                result.whitespaceRatio = distanceToBottom / frameHeight;
                
                // 如果留白超过文本框高度的10%，则认为不合适
                if (result.whitespaceRatio > 0.1) {
                    result.fitWell = false;
                    result.tooMuchWhitespace = true;
                }
            }
        } catch (e) {
            // 出错时认为不合适
            result.fitWell = false;
            return result;
        }
        
        return result;
    },
    
    // 调整文本属性
    adjustTextProperties: function(textFrame, options) {
        var minFontSize = options.minFontSize;
        var maxFontSize = options.maxFontSize;
        var lineHeightAdjustRange = options.lineHeightAdjustRange;
        var trackingAdjustRange = options.trackingAdjustRange;
        var fitStatus = options.fitStatus;
        
        // 如果是因为留白过多而不合适，应该尝试放大文本
        var shouldEnlargeText = fitStatus.tooMuchWhitespace && !fitStatus.overflow;
        
        // 先尝试调整字体大小
        var fontSizeResult = this.adjustFontSize(textFrame, minFontSize, maxFontSize, shouldEnlargeText);
        
        // 如果仅调整字体大小就能解决问题，则返回
        var currentFitStatus = this.isTextFitWell(textFrame);
        if (!textFrame.overflows && currentFitStatus.fitWell) {
            return;
        }
        
        // 调整行间距
        var leadingResult = this.adjustLeading(textFrame, lineHeightAdjustRange);
        
        // 如果调整行间距能解决问题，则返回
        currentFitStatus = this.isTextFitWell(textFrame);
        if (!textFrame.overflows && currentFitStatus.fitWell) {
            return;
        }
        
        // 最后调整字间距
        var trackingResult = this.adjustTracking(textFrame, trackingAdjustRange);
        
        // 如果所有调整都无法解决问题，抛出异常
        currentFitStatus = this.isTextFitWell(textFrame);
        if (textFrame.overflows || !currentFitStatus.fitWell) {
            throw new Error("无法通过调整文本属性来适配文本框");
        }
    },
    
    // 调整字体大小
    adjustFontSize: function(textFrame, minSize, maxSize, shouldEnlarge) {
        var currentFontSize = this.getFontSize(textFrame);
        var overflow = textFrame.overflows;
        
        // 如果当前字体大小已经在范围内，且不需要调整，则不需要调整
        if (currentFontSize >= minSize && currentFontSize <= maxSize && !overflow && !shouldEnlarge) {
            return { changed: false, value: currentFontSize };
        }
        
        var low = minSize;
        var high = maxSize;
        var bestFit = currentFontSize;
        var found = false;
        
        if (shouldEnlarge) {
            // 如果需要放大文本，从当前大小到最大值之间查找
            low = currentFontSize;
            high = maxSize;
            this.applyLeading(textFrame, Leading.AUTO); 
        }
        
        // 二分法查找最佳字体大小
        while (low <= high) {
            var mid = Math.round((low + high) / 2);
            this.applyFontSize(textFrame, mid);
            
            if (!textFrame.overflows) {
                bestFit = mid;
                if (shouldEnlarge) {
                    low = mid + 1;  // 如果需要放大，继续寻找更大的值
                } else {
                    low = mid + 1;
                }
                found = true;
            } else {
                high = mid - 1;
            }
        }
        
        // 应用最佳字体大小
        if (found) {
            this.applyFontSize(textFrame, bestFit);
        } else {
            // 如果最小字体仍然溢出，则应用最小字体
            this.applyFontSize(textFrame, minSize);
        }
        
        return { changed: true, value: bestFit };
    },
    
    // 应用字体大小
    applyFontSize: function(textFrame, size) {
        try {
            textFrame['parentStory']['pointSize'] = size;
        } catch (e) {
            // 忽略错误
        }
    },
    
    // 调整行间距
    adjustLeading: function(textFrame, adjustRange) {
        var originalLeading = this.getLeading(textFrame);
        var overflow = textFrame.overflows;
        
        if (!overflow) {
            return { changed: false, value: originalLeading };
        }
        
        // 尝试减小行间距
        var minLeading = originalLeading * (1 - adjustRange);
        var maxLeading = originalLeading;
        
        var low = minLeading;
        var high = maxLeading;
        var bestLeading = originalLeading;
        var found = false;
        
        // 二分法查找最佳行间距
        while (low <= high) {
            var mid = (low + high) / 2;
            this.applyLeading(textFrame, mid);
            
            if (!textFrame.overflows) {
                bestLeading = mid;
                low = mid + 0.1;
                found = true;
            } else {
                high = mid - 0.1;
            }
        }
        
        // 应用最佳行间距
        if (found) {
            this.applyLeading(textFrame, bestLeading);
        } else {
            // 恢复原始行间距
            this.applyLeading(textFrame, originalLeading);
        }
        
        return { changed: found, value: bestLeading };
    },
    
    // 应用行间距
    applyLeading: function(textFrame, leading) {
        try {
            textFrame['parentStory']['leading'] = leading;
        } catch (e) {
            // 忽略错误
        }
    },
    
    // 调整字间距
    adjustTracking: function(textFrame, adjustRange) {
        var originalTracking = this.getTracking(textFrame);
        var overflow = textFrame.overflows;
        
        if (!overflow) {
            return { changed: false, value: originalTracking };
        }
        
        // 尝试减小字间距
        var minTracking = originalTracking * (1 - adjustRange);
        var maxTracking = originalTracking;
        
        var low = minTracking;
        var high = maxTracking;
        var bestTracking = originalTracking;
        var found = false;
        
        // 二分法查找最佳字间距
        while (low <= high) {
            var mid = (low + high) / 2;
            this.applyTracking(textFrame, mid);
            
            if (!textFrame.overflows) {
                bestTracking = mid;
                low = mid + 1;
                found = true;
            } else {
                high = mid - 1;
            }
        }
        
        // 应用最佳字间距
        if (found) {
            this.applyTracking(textFrame, bestTracking);
        } else {
            // 恢复原始字间距
            this.applyTracking(textFrame, originalTracking);
        }
        
        return { changed: found, value: bestTracking };
    },
    
    // 应用字间距
    applyTracking: function(textFrame, tracking) {
        try {
            textFrame.texts[0].tracking = tracking;
        } catch (e) {
            // 忽略错误
        }
    }
};

// 主函数调用
if (app.documents.length > 0) {
    showMainInterface();
} else {
    alert("请先打开一个InDesign文档。");
}
