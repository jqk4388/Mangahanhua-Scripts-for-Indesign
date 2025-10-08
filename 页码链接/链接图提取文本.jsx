
// 链接图提取文本
// 作者: 几千块
// 功能: 提取链接的PSD/TIFF图片中的文字图层，并在InDesign中创建对应的文本框
// 说明: 脚本通过BridgeTalk调用Photoshop进行文字提取，注意JSON.stringify在ES3环境中不可用，故使用自定义的字符串化函数,变量名不能有btn字符等会被bt转译
// 版本: 1.0
// 日期: 2025-09-20
var st = 2000;
// 主函数
function extractTextFromLinkedImages() {
    var doc = app.activeDocument;
    var tempFolder = Folder.temp;
    
    // 1. 弹出窗口询问用户执行范围
    var scope = promptUserForScope();
    if (scope === null) return; // 用户取消操作
    
    // 获取链接的PSD/TIFF图片
    var linkedImages = getLinkedPSDAndTIFFImages(doc, scope);
    if (linkedImages.length === 0) {
        alert("未找到链接的PSD或TIFF图片");
        return;
    }
    
    // 处理每个链接图片
    for (var i = 0; i < linkedImages.length; i++) {
        var image = linkedImages[i];
        var linkFile = image.itemLink.filePath;
        
        // 检查文件是否存在
        var fileObj = new File(linkFile);
        if (!fileObj.exists) {
            alert("文件不存在: " + linkFile);
            continue;
        }
        
        // 创建临时输出文件
        var outputFile = new File(tempFolder + "/output_text.txt");
        if (outputFile.exists) {
            outputFile.remove();
        }
        
        // 调用Photoshop提取文本
        if (extractTextWithPhotoshop(fileObj, outputFile)) {
            // 读取提取的文本并在InDesign中创建文本框
            createTextFramesFromOutput(doc, outputFile, image);
        }
    }
    
    alert("文本提取完成");
}

// 弹出对话框询问用户执行范围
function promptUserForScope() {
    var dialog = new Window("dialog", "选择执行范围");
    dialog.orientation = "column";
    dialog.alignChildren = "left";
    
    var text = dialog.add("statictext", undefined, "请选择执行范围:");
    
    var scopeGroup = dialog.add("group");
    scopeGroup.orientation = "column";
    scopeGroup.alignChildren = "left";
    
    var currentPageRadio = scopeGroup.add("radiobutton", undefined, "当前页面");
    currentPageRadio.value = true;
    var documentRadio = scopeGroup.add("radiobutton", undefined, "整个文档");
    
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var okButton = buttonGroup.add("button", undefined, "确定");
    var cancelButton = buttonGroup.add("button", undefined, "取消");
    
    var scope = null;
    
    okButton.onClick = function() {
        scope = documentRadio.value ? "document" : "page";
        dialog.close();
    };
    
    cancelButton.onClick = function() {
        scope = null;
        dialog.close();
    };
    
    dialog.show();
    
    return scope;
}

// 获取链接的PSD和TIFF图片
function getLinkedPSDAndTIFFImages(doc, scope) {
    var images = [];
    var allGraphics;
    
    if (scope === "page") {
        // 只获取当前页面的图形
        allGraphics = app.activeWindow.activePage.allGraphics;
    } else {
        // 获取整个文档的图形
        allGraphics = doc.allGraphics;
    }
    
    for (var i = 0; i < allGraphics.length; i++) {
        var graphic = allGraphics[i];
        // 检查是否为链接图像且格式为PSD或TIFF
        if (graphic.hasOwnProperty("itemLink") && graphic.itemLink) {
            var link = graphic.itemLink;
            var fileName = link.name.toLowerCase();
            if (fileName.match(/\.(psd|tif|tiff)$/)) {
                images.push(graphic);
            }
        }
    }
    
    return images;
}

// 使用BridgeTalk调用Photoshop执行文本提取脚本
function extractTextWithPhotoshop(imageFile, outputFile) {
    try {
        if (BridgeTalk.isRunning("photoshop")) {
            var bt = new BridgeTalk();
            bt.target = "photoshop";
            bt.body = psScript.toString() + "\rpsScript('" + imageFile.fullName + "', '" + outputFile.fullName + "');";
            
            // 接收 Photoshop 返回的结果
            bt.onResult = function(inBT) {
                // 处理返回的结果
                var result = inBT.body;
                // 这里可以添加对结果的处理逻辑
            };
            
            // 捕获通信或执行错误
            bt.onError = function(inBT) { 
                alert(inBT.body); // 显示错误信息
            };
            
            bt.send();
            // 等待一段时间确保Photoshop完成处理
            $.sleep(st);
            // Photoshop执行脚本函数
            function psScript(inputFilePath, outputFilePath) {
                try {
                    var inputFile = new File(inputFilePath);
                    var outputFile = new File(outputFilePath);
                    
                    if (!inputFile.exists) {
                        return "输入文件不存在: " + inputFilePath;
                    }
                    
                    open(inputFile);
                    var doc = app.activeDocument;
                    var textInfo = "";
                    
                    extractTextLayers(doc, "", doc);
                    // Json.stringify的替代实现，适用于ES3环境
                    function stringify(value) {
                        // 存储已处理的对象，防止循环引用（ES3中用数组模拟）
                        var processed = [];
                        
                        // ES3环境下判断是否为数组（替代Array.isArray）
                        function isArray(val) {
                            return Object.prototype.toString.call(val) === "[object Array]";
                        }
                        
                        // ES3环境下的indexOf替代方法
                        function indexOf(arr, val) {
                            for (var i = 0; i < arr.length; i++) {
                                if (arr[i] === val) {
                                    return i;
                                }
                            }
                            return -1;
                        }
                        
                        function serialize(val) {
                            var type = typeof val;
                            
                            // 处理null（ES3中typeof null会返回"object"，需要单独判断）
                            if (val === null) {
                                return "null";
                            }
                            
                            // 处理布尔值
                            if (type === "boolean") {
                                return val ? "true" : "false";
                            }
                            
                            // 处理数字
                            if (type === "number") {
                                return isFinite(val) ? val.toString() : "null";
                            }
                            
                            // 处理字符串
                            if (type === "string") {
                                // 转义特殊字符（ES3兼容写法）
                                return '"' + val
                                    .replace(/[\\]/g, '\\\\')
                                    .replace(/["]/g, '\\"')
                                    .replace(/[\b]/g, '\\b')
                                    .replace(/[\f]/g, '\\f')
                                    .replace(/[\n]/g, '\\n')
                                    .replace(/[\r]/g, '\\r')
                                    .replace(/[\t]/g, '\\t') + '"';
                            }
                            
                            // 处理数组（ES3中通过Object.prototype.toString判断）
                            if (isArray(val)) {
                                // 检查循环引用
                                if (indexOf(processed, val) !== -1) {
                                    return '"[Circular Reference]"';
                                }
                                processed[processed.length] = val; // 替代push()
                                
                                var elements = [];
                                for (var i = 0; i < val.length; i++) {
                                    elements[elements.length] = serialize(val[i]);
                                }
                                return "[" + elements.join(",") + "]";
                            }
                            
                            // 处理对象
                            if (type === "object") {
                                // 检查循环引用
                                if (indexOf(processed, val) !== -1) {
                                    return '"[Circular Reference]"';
                                }
                                processed[processed.length] = val; // 替代push()
                                
                                var props = [];
                                for (var key in val) {
                                    // 仅序列化自有属性（ES3兼容）
                                    if (val.hasOwnProperty(key)) {
                                        props[props.length] = serialize(key) + ":" + serialize(val[key]);
                                    }
                                }
                                return "{" + props.join(",") + "}";
                            }
                            
                            // 处理不支持的类型（函数、正则等）
                            return '"[Unsupported Type: ' + type + ']"';
                        }
                        
                        return serialize(value);
                    }
                    function extractTextLayers(layer, prefix, doc) {
                        for (var i = 0; i < layer.layers.length; i++) {
                            var subLayer = layer.layers[i];
                            if (subLayer.typename == "ArtLayer" && subLayer.kind == LayerKind.TEXT) {
                                var textItem = subLayer.textItem;
                                var bounds = subLayer.bounds;
                                var info = {
                                    ame: subLayer.name,
                                    co: textItem.contents.replace(/[\r\n]/g, ""), // 移除换行符,
                                    le: bounds[0] / doc.width,
                                    op: bounds[1] / doc.height,
                                    ig: bounds[2] / doc.width,
                                    om: bounds[3] / doc.height,
                                };
                                textInfo += stringify(info) + "\n";
                            } else if (subLayer.typename == "LayerSet") {
                                extractTextLayers(subLayer, prefix + subLayer.name + "/", doc);
                            }
                        }
                    }
                    
                    var file = new File(outputFile);
                    file.encoding = "UTF-8";
                    file.open("w");
                    file.write(textInfo);
                    file.close();
                    doc.close(SaveOptions.DONOTSAVECHANGES);
                    
                    return "success";
                } catch (e) {
                    return "Error: " + e.message;
                }
            }
            return outputFile.exists && outputFile.length > 0;
        } else {
            alert("Photoshop未运行，请先启动Photoshop");
            return false;
        }
    } catch (e) {
        alert("调用Photoshop时出错: " + e.message);
        return false;
    }
}


// 根据输出文件创建文本框
function createTextFramesFromOutput(doc, outputFile, image) {
    if (!outputFile.exists) return;
    
    // 读取输出文件内容
    outputFile.open("r");
    outputFile.encoding = "UTF-8";
    var content = outputFile.read();
    outputFile.close();
    
    // 解析文本信息
    var lines = content.split("\\n");
    var parentPage = image.parentPage;
    
    if (!parentPage) {
        // 如果图形不在页面上，尝试找到其父页面
        var parent = image.parent;
        while (parent && parent.hasOwnProperty("parentPage") && !parent.parentPage) {
            parent = parent.parent;
        }
        if (parent && parent.hasOwnProperty("parentPage")) {
            parentPage = parent.parentPage;
        }
    }
    
    if (!parentPage) {
        alert("无法确定图形所在的页面");
        return;
    }
    
    // 获取图像的位置信息
    var imageBounds = parentPage.bounds;
    var imageLeft = imageBounds[1];
    var imageTop = imageBounds[0];
    var imageWidth = imageBounds[3] - imageBounds[1];
    var imageHeight = imageBounds[2] - imageBounds[0];
    function trim(str) {
    // 先判断输入是否为字符串，非字符串则转换为字符串
    if (typeof str !== 'string') {
        str = String(str);
    }
    return str.replace(/^\s+/, '').replace(/\s+$/, '');
}
    // 为每个文本条目创建文本框
    for (var i = 0; i < lines.length; i++) {
        if (trim(lines[i]) === "") continue;
        
        try {
            //去掉换行符，eval中不能换行
            var line = lines[i].replace(/\n/g, "\\r");
            var textInfo = eval("(" + line + ")"); // 使用eval解析JSON（因为是ES3环境）
            
            // 计算文本框位置（相对于页面位置）
            var left = textInfo.le * imageWidth;
            var top = textInfo.op * imageHeight;
            var right = textInfo.ig  * imageWidth;
            var bottom = textInfo.om * imageHeight;
            
            // 创建文本框
            var textFrame = parentPage.textFrames.add();
            textFrame.geometricBounds = [top, left, bottom, right];
            textFrame.contents = textInfo.co;
            textFrame.parentStory.storyPreferences.storyOrientation = StoryHorizontalOrVertical.VERTICAL;
            textFrame.fit(FitOptions.FRAME_TO_CONTENT);

            
            // 设置基本文本属性
            textFrame.textFramePreferences.verticalJustification = VerticalJustification.TOP_ALIGN;
            
        } catch (e) {
            // 忽略解析错误的行
            continue;
        }
    }
}

// 执行主函数
extractTextFromLinkedImages();