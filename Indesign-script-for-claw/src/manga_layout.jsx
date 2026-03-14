/*
 * Manga Layout Automation Script (Headless)
 * 自动化漫画排版脚本 - 无界面版本
 * 
 * 功能：
 * 1. 创建新文档并设置页面尺寸
 * 2. 新建图层结构
 * 3. 导入样式和复合字体
 * 4. 导入图片
 * 5. 导入LabelPlus格式的txt稿件
 * 6. 应用样式匹配
 * 7. 执行断句
 * 
 * 作者：几千块
 * 版本：1.0.0
 */

// ==================== 依赖引入 ====================
// 引入JSON解析库（ES3兼容）
#include "../../Library/json2.js"

// ==================== 全局变量 ====================
var SCRIPT_VERSION = "1.0.0";
var SCRIPT_NAME = "漫画全自动嵌字";
var config = null;
var doc = null;
var logMessages = [];
var errorMessages = [];

// ==================== 主入口函数 ====================
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;// 设置无界面模式

// 步骤超时时间（毫秒），默认3分钟
var STEP_TIMEOUT = 3*60000;

// 步骤执行结果结构
function StepResult(success, message, data) {
    this.success = success;
    this.message = message || "";
    this.data = data || null;
}

// 带超时的步骤执行器
function executeStepWithTimeout(stepName, stepFunc, timeoutMs) {
    timeoutMs = timeoutMs || STEP_TIMEOUT;
    var startTime = new Date().getTime();
    
    try {
        logMessage("[执行] " + stepName + " (超时: " + (timeoutMs / 1000) + "秒)");
        var result = stepFunc();
        var elapsed = new Date().getTime() - startTime;
        
        if (result && result.success !== undefined) {
            result.elapsed = elapsed;
            if (result.success) {
                logMessage("[成功] " + stepName + " (耗时: " + (elapsed / 1000).toFixed(2) + "秒)");
            } else {
                logMessage("[失败] " + stepName + " - " + result.message);
            }
            return result;
        }
        
        // 如果函数没有返回StepResult，包装一个成功结果
        return new StepResult(true, "完成", result);
        
    } catch (e) {
        var elapsed = new Date().getTime() - startTime;
        var errorMsg = e.message + " (行: " + e.line + ")";
        logMessage("[异常] " + stepName + " - " + errorMsg);
        return new StepResult(false, errorMsg, null);
    }
}

// 主函数
function main() {
    var stepResults = {};
    var startTime = new Date();
    
    logMessage("===== " + SCRIPT_NAME + " v" + SCRIPT_VERSION + " 开始执行 =====");
    logMessage("开始时间: " + startTime.toLocaleString());
    
    // ========== 步骤1: 读取配置文件 ==========
    var step1 = executeStepWithTimeout("步骤1: 读取配置文件", function() {
        config = loadConfig();
        if (!config) {
            return new StepResult(false, "无法加载配置文件");
        }
        return new StepResult(true, "配置文件加载成功");
    });
    stepResults.step1 = step1;
    
    if (!step1.success) {
        return buildFinalResult(false, "配置文件加载失败，无法继续", stepResults, startTime);
    }
    
    // ========== 步骤2-5: 核心步骤，失败直接退出 ==========
    
    // 步骤2: 创建新文档
    var step2 = executeStepWithTimeout("步骤2: 创建新文档", function() {
        var result = createDocumentWithResult();
        if (result.success) {
            doc = result.data;
        }
        return result;
    });
    stepResults.step2 = step2;
    
    if (!step2.success) {
        return buildFinalResult(false, "创建文档失败，退出执行", stepResults, startTime);
    }
    
    // 如果从模板新建，跳过步骤3-5
    if (isFromTemplate) {
        logMessage("[跳过] 步骤3-5: 从模板新建，跳过页面设置、图层创建、主页图框创建");
        stepResults.step3 = new StepResult(true, "跳过（从模板新建）");
        stepResults.step4 = new StepResult(true, "跳过（从模板新建）");
        stepResults.step5 = new StepResult(true, "跳过（从模板新建）");
    } else {
        // 步骤3: 设置页面尺寸和文档属性
        var step3 = executeStepWithTimeout("步骤3: 设置页面尺寸和属性", function() {
            return setupDocumentPreferencesWithResult();
        });
        stepResults.step3 = step3;
        
        if (!step3.success) {
            return buildFinalResult(false, "页面设置失败，退出执行", stepResults, startTime);
        }
        
        // 步骤4: 创建图层结构
        var step4 = executeStepWithTimeout("步骤4: 创建图层结构", function() {
            return createLayersWithResult();
        });
        stepResults.step4 = step4;
        
        if (!step4.success) {
            return buildFinalResult(false, "创建图层失败，退出执行", stepResults, startTime);
        }
        
        // 步骤5: 创建主页图框
        var step5 = executeStepWithTimeout("步骤5: 创建主页图框", function() {
            return createMasterFramesWithResult();
        });
        stepResults.step5 = step5;
        
        if (!step5.success) {
            return buildFinalResult(false, "创建主页图框失败，退出执行", stepResults, startTime);
        }
    }
    
    // ========== 步骤6-10: 可选步骤，失败跳过 ==========
    
    // 步骤6: 导入样式和复合字体
    if (config.styleImport.enabled && config.styleImport.styletemplatePath) {
        var step6 = executeStepWithTimeout("步骤6: 导入样式和复合字体", function() {
            return importStylesAndFontsWithResult();
        });
        stepResults.step6 = step6;
        
        if (!step6.success) {
            logMessage("[警告] 步骤6失败，跳过继续执行: " + step6.message);
            errorMessages.push("步骤6: " + step6.message);
        }
    } else {
        logMessage("[步骤6] 跳过样式导入（未启用或未指定模板路径）");
        stepResults.step6 = new StepResult(true, "跳过");
    }
    
    // 步骤7: 导入图片
    if (config.imageImport.enabled && config.imageImport.artFolderPath) {
        var step7 = executeStepWithTimeout("步骤7: 导入图片", function() {
            return importImagesWithResult();
        });
        stepResults.step7 = step7;
        
        if (!step7.success) {
            logMessage("[警告] 步骤7失败，跳过继续执行: " + step7.message);
            errorMessages.push("步骤7: " + step7.message);
        }
    } else {
        logMessage("[步骤7] 跳过图片导入（未启用或未指定图片路径）");
        stepResults.step7 = new StepResult(true, "跳过");
    }
    
    // 步骤8: 导入LabelPlus文本
    if (config.textImport.enabled && config.textImport.lpTxtPath) {
        var step8 = executeStepWithTimeout("步骤8: 导入LabelPlus文本", function() {
            return importLabelPlusTextWithResult();
        });
        stepResults.step8 = step8;
        
        if (!step8.success) {
            logMessage("[警告] 步骤8失败，跳过继续执行: " + step8.message);
            errorMessages.push("步骤8: " + step8.message);
        }
    } else {
        logMessage("[步骤8] 跳过文本导入（未启用或未指定文本路径）");
        stepResults.step8 = new StepResult(true, "跳过");
    }
    
    // 步骤9: 执行样式匹配
    if (config.fontMapping.enabled) {
        var step9 = executeStepWithTimeout("步骤9: 执行样式匹配", function() {
            return runStyleMatchingWithResult();
        });
        stepResults.step9 = step9;
        
        if (!step9.success) {
            logMessage("[警告] 步骤9失败，跳过继续执行: " + step9.message);
            errorMessages.push("步骤9: " + step9.message);
        }
    } else {
        logMessage("[步骤9] 跳过样式匹配（未启用或未指定配置文件）");
        stepResults.step9 = new StepResult(true, "跳过");
    }
    
    // 步骤10: 执行断句
    if (config.segmentation.enabled) {
        var step10 = executeStepWithTimeout("步骤10: 执行断句", function() {
            return runSegmentationWithResult();
        });
        stepResults.step10 = step10;
        
        if (!step10.success) {
            logMessage("[警告] 步骤10失败，跳过继续执行: " + step10.message);
            errorMessages.push("步骤10: " + step10.message);
        }
    } else {
        logMessage("[步骤10] 跳过断句（未启用）");
        stepResults.step10 = new StepResult(true, "跳过");
    }
    
    // ========== 步骤11: 保存文档（最后执行）==========
    if (config.output.saveDocument) {
        var step11 = executeStepWithTimeout("步骤11: 保存文档", function() {
            return saveDocumentWithResult();
        });
        stepResults.step11 = step11;
        
        if (!step11.success) {
            logMessage("[错误] 步骤11保存文档失败: " + step11.message);
            errorMessages.push("步骤11: " + step11.message);
        }
    } else {
        logMessage("[步骤11] 跳过文档保存（未启用）");
        stepResults.step11 = new StepResult(true, "跳过");
    }
    
    // 返回最终结果
    return buildFinalResult(true, "排版流程完成", stepResults, startTime);
}

// 构建最终返回结果
function buildFinalResult(success, message, stepResults, startTime) {
    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;
    
    logMessage("===== 执行" + (success ? "完成" : "终止") + " =====");
    logMessage("结束时间: " + endTime.toLocaleString());
    logMessage("总耗时: " + duration.toFixed(2) + " 秒");
    
    // 写入日志
    if (config && config.output && config.output.logEnabled) {
        writeLogFile();
    }
    
    // 输出到控制台
    $.writeln(logMessages.join("\n"));
    
    return {
        success: success,
        message: message,
        duration: duration,
        steps: stepResults,
        errors: errorMessages.length > 0 ? errorMessages : null,
        document: doc
    };
}

// ==================== 配置加载函数 ====================
function loadConfig() {
    try {
        // 获取脚本所在目录
        var scriptPath = File($.fileName).parent.fsName;
        var configPath = scriptPath + "\\manga_layout_config.json";
        
        // 兼容Mac路径
        if ($.os.indexOf("Windows") === -1) {
            configPath = scriptPath + "/manga_layout_config.json";
        }
        
        var configFile = new File(configPath);
        if (!configFile.exists) {
            throw new Error("配置文件不存在: " + configPath);
        }
        
        configFile.encoding = "UTF-8";
        configFile.open("r");
        var configText = configFile.read();
        configFile.close();
        
        var parsedConfig = JSON.parse(configText);
        return parsedConfig;
        
    } catch (e) {
        logMessage("加载配置文件失败: " + e.message);
        return null;
    }
}

// ==================== 文档创建函数 ====================
// 全局标记：是否从模板新建
var isFromTemplate = false;

function createDocumentWithResult() {
    try {
        var newDoc = null;
        
        // 检查是否配置了模板文档
        if (config.templateDocument && 
            config.templateDocument.enabled && 
            config.templateDocument.indtPath) {
            
            var templatePath = config.templateDocument.indtPath;
            var templateFile = new File(templatePath);
            
            if (!templateFile.exists) {
                return new StepResult(false, "模板文件不存在: " + templatePath);
            }
            
            // 从模板新建文档
            logMessage("从模板新建文档: " + templatePath);
            newDoc = app.open(templateFile, true); // true = 打开为副本
            
            if (!newDoc || !newDoc.isValid) {
                return new StepResult(false, "从模板创建文档失败: 文档对象无效");
            }
            
            isFromTemplate = true;
            return new StepResult(true, "从模板创建文档成功", newDoc);
        }
        
        // 默认：创建空白文档
        newDoc = app.documents.add();
        if (!newDoc || !newDoc.isValid) {
            return new StepResult(false, "创建文档失败: 文档对象无效");
        }
        
        isFromTemplate = false;
        return new StepResult(true, "文档创建成功", newDoc);
        
    } catch (e) {
        return new StepResult(false, "创建文档失败: " + e.message);
    }
}

// ==================== 文档设置函数 ====================
function setupDocumentPreferencesWithResult() {
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        var settings = config.documentSettings;
        
        // 设置单位
        if (settings.unit === "mm") {
            app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;
        } else if (settings.unit === "pt") {
            app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
        } else {
            app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;
        }
        
        // 设置页面尺寸
        doc.documentPreferences.pageWidth = settings.pageWidth;
        doc.documentPreferences.pageHeight = settings.pageHeight;
        
        // 设置装订方向
        if (settings.pageBinding === "rightToLeft") {
            doc.documentPreferences.pageBinding = PageBindingOptions.RIGHT_TO_LEFT;
        } else {
            doc.documentPreferences.pageBinding = PageBindingOptions.LEFT_TO_RIGHT;
        }
        
        // 设置出血
        doc.documentPreferences.documentBleedTopOffset = settings.bleedTop;
        doc.documentPreferences.documentBleedBottomOffset = settings.bleedBottom;
        doc.documentPreferences.documentBleedInsideOrLeftOffset = settings.bleedInside;
        doc.documentPreferences.documentBleedOutsideOrRightOffset = settings.bleedOutside;
        
        // 设置边距
        if (settings.margin) {
            doc.marginPreferences.top = settings.margin.top;
            doc.marginPreferences.bottom = settings.margin.bottom;
            doc.marginPreferences.left = settings.margin.inside;
            doc.marginPreferences.right = settings.margin.outside;
        }
        
        // 设置标尺原点
        doc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
        
        // 如果指定了总页数，添加页面
        if (settings.totalPages > 1) {
            for (var i = 1; i < settings.totalPages; i++) {
                doc.pages.add();
            }
        }
        
        return new StepResult(true, "页面设置完成");
        
    } catch (e) {
        return new StepResult(false, "设置文档属性失败: " + e.message);
    }
}

// ==================== 图层创建函数 ====================
function createLayersWithResult() {
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        var layersConfig = config.layers;
        
        // 获取图层配置并按顺序创建
        var layerOrder = ["art", "design", "clean", "pageNumber", "text", "sfx"];
        
        // 首先重命名默认图层为Art图层
        if (doc.layers.length > 0) {
            var defaultLayer = doc.layers[0];
            if (layersConfig.art) {
                defaultLayer.name = layersConfig.art.name;
                defaultLayer.visible = layersConfig.art.visible;
                defaultLayer.locked = layersConfig.art.locked;
            }
        }
        
        // 创建其他图层（从下往上）
        var layerStack = [];
        for (var i = 0; i < layerOrder.length; i++) {
            var key = layerOrder[i];
            var layerConfig = layersConfig[key];
            if (layerConfig && key !== "art") {
                layerStack.push(layerConfig);
            }
        }
        
        // 反向添加图层（确保正确的堆叠顺序）
        for (var j = layerStack.length - 1; j >= 0; j--) {
            var cfg = layerStack[j];
            var newLayer = doc.layers.add({
                name: cfg.name
            });
            newLayer.visible = cfg.visible;
            newLayer.locked = cfg.locked;
        }
        
        return new StepResult(true, "图层创建完成，共 " + doc.layers.length + " 个图层");
        
    } catch (e) {
        return new StepResult(false, "创建图层失败: " + e.message);
    }
}

// ==================== 主页图框创建函数 ====================
function createMasterFramesWithResult() {
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        // 创建主页
        var masterSpread = doc.masterSpreads[0];
        var artLayer = doc.layers.itemByName(config.layers.art.name);
        doc.activeLayer = artLayer;
        
        var frameCount = 0;        
        // 为每个主页页面创建图框
        for (var i = 0; i < masterSpread.pages.length; i++) {
            var masterPage = masterSpread.pages[i];
            
            // 创建图框
            var artFrame = masterPage.rectangles.add({
                name: 'ArtFrame',
                layer: artLayer,
                geometricBounds: masterPage.bounds,
                contentType: ContentType.GRAPHIC_TYPE
            });
            
            // 设置填充和描边为无
            artFrame.fillColor = doc.swatches.itemByName("None");
            artFrame.strokeColor = doc.swatches.itemByName("None");
            
            // 设置适合选项
            artFrame.frameFittingOptions.fittingAlignment = AnchorPoint.CENTER_ANCHOR;
            artFrame.frameFittingOptions.fittingOnEmptyFrame = EmptyFrameFittingOptions.FILL_PROPORTIONALLY;
            
            // 扩展图框到出血线
            expandFrameToBleed(artFrame, masterPage);
            frameCount++;
        }
        
        return new StepResult(true, "主页图框创建完成，共 " + frameCount + " 个");
        
    } catch (e) {
        return new StepResult(false, "创建主页图框失败: " + e.message);
    }
}

// 扩展图框到出血线
function expandFrameToBleed(frame, page) {
    try {
        var bleedTop = config.documentSettings.bleedTop;
        var bleedBottom = config.documentSettings.bleedBottom;
        var bleedInside = config.documentSettings.bleedInside;
        var bleedOutside = config.documentSettings.bleedOutside;
        
        var pageBounds = page.bounds;
        var isRightPage = page.side === PageSideOptions.RIGHT_HAND;
        
        // 根据页面位置设置出血
        if (isRightPage) {
            frame.geometricBounds = [
                pageBounds[0] - bleedTop,
                pageBounds[1],
                pageBounds[2] + bleedBottom,
                pageBounds[3] + bleedOutside
            ];
        } else {
            frame.geometricBounds = [
                pageBounds[0] - bleedTop,
                pageBounds[1] - bleedOutside,
                pageBounds[2] + bleedBottom,
                pageBounds[3]
            ];
        }
    } catch (e) {
        logMessage("扩展图框到出血线时出错: " + e.message);
    }
}

// ==================== 样式导入函数 ====================
function importStylesAndFontsWithResult() {
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        var templatePath = config.styleImport.styletemplatePath;
        var templateFile = new File(templatePath);
        
        if (!templateFile.exists) {
            return new StepResult(false, "模板文件不存在: " + templatePath);
        }
        
        // 打开源文档（不显示）
        var sourceDoc = app.open(templateFile, false);
        if (!sourceDoc || !sourceDoc.isValid) {
            return new StepResult(false, "无法打开模板文件: " + templatePath);
        }
        
        var importCount = 0;
        
        try {
            // 导入段落和字符样式
            if (config.styleImport.importParagraphStyles) {
                doc.importStyles(
                    ImportFormat.textStylesFormat, 
                    templateFile, 
                    GlobalClashResolutionStrategy.loadAllWithOverwrite
                );
                importCount++;
            }
            
            // 导入对象样式
            if (config.styleImport.importObjectStyles) {
                doc.importStyles(
                    ImportFormat.OBJECT_STYLES_FORMAT, 
                    templateFile, 
                    GlobalClashResolutionStrategy.loadAllWithOverwrite
                );
                importCount++;
            }
            
            // 复制复合字体
            if (config.styleImport.importCompositeFonts) {
                var fontCount = copyCompositeFontsWithCount(sourceDoc, doc);
                importCount++;
                logMessage("导入复合字体: " + fontCount + " 个");
            }
            
        } finally {
            sourceDoc.close(SaveOptions.NO);
        }
        
        return new StepResult(true, "样式导入完成，共导入 " + importCount + " 类样式");
        
    } catch (e) {
        return new StepResult(false, "导入样式失败: " + e.message);
    }
}

// 复制复合字体（返回复制的数量）
function copyCompositeFontsWithCount(sourceDoc, targetDoc) {
    var copiedCount = 0;
    try {
        var sourceFonts = sourceDoc.compositeFonts;
        var targetFonts = targetDoc.compositeFonts;
        
        for (var i = 1; i < sourceFonts.length; i++) {
            var srcFont = sourceFonts[i];
            var newFont = null;
            
            try {
                newFont = targetFonts.add({ name: srcFont.name });
            } catch (e) {
                continue; // 字体已存在，跳过
            }
            
            // 复制字体条目
            var srcFontEntries = srcFont.compositeFontEntries;
            var newFontEntries = newFont.compositeFontEntries;
            
            for (var j = 0; j < srcFontEntries.length; j++) {
                var srcEntry = srcFontEntries[j];
                var newEntry = newFontEntries[j];
                if (!newEntry) continue;
                
                try {
                    newEntry.name = srcEntry.name;
                    newEntry.appliedFont = srcEntry.appliedFont;
                    newEntry.fontStyle = srcEntry.fontStyle;
                    
                    if (j !== 0) {
                        newEntry.horizontalScale = srcEntry.horizontalScale;
                        newEntry.verticalScale = srcEntry.verticalScale;
                        newEntry.baselineShift = srcEntry.baselineShift;
                        newEntry.relativeSize = srcEntry.relativeSize;
                        newEntry.scaleOption = srcEntry.scaleOption;
                    }
                    
                    newEntry.customCharacters = srcEntry.customCharacters;
                } catch (ee) {
                    // 忽略单个属性复制错误
                }
            }
            copiedCount++;
        }
    } catch (e) {
        logMessage("复制复合字体时出错: " + e.message);
    }
    return copiedCount;
}

// ==================== 图片导入函数 ====================
function importImagesWithResult() {
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        var artFolder = new Folder(config.imageImport.artFolderPath);
        if (!artFolder.exists) {
            return new StepResult(false, "图片文件夹不存在: " + artFolder.fsName);
        }
        
        // 获取图片文件列表
        var extensions = config.imageImport.fileExtensions;
        var artFiles = [];
        var allFiles = artFolder.getFiles();
        
        for (var i = 0; i < allFiles.length; i++) {
            var file = allFiles[i];
            if (file instanceof File) {
                var ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
                for (var k = 0; k < extensions.length; k++) {
                    if (extensions[k] === ext) {
                        artFiles.push(file);
                        break;
                    }
                }
            }
        }
        
        if (artFiles.length === 0) {
            return new StepResult(false, "未找到任何图片文件");
        }
        
        // 按文件名排序
        artFiles.sort(function(a, b) {
            return a.name.localeCompare(b.name);
        });
        
        logMessage("找到 " + artFiles.length + " 个图片文件");
        
        // 找出文件名模式
        var filenames = [];
        for (var i = 0; i < artFiles.length; i++) {
            filenames.push(artFiles[i].name);
        }
        var pattern = findCommonPattern(filenames);
        
        // 根据图片数量添加页面
        var requiredPages = artFiles.length;
        if (doc.pages.length < requiredPages) {
            var pagesToAdd = requiredPages - doc.pages.length;
            for (var p = 0; p < pagesToAdd; p++) {
                doc.pages.add();
            }
        }
        
        // 放置图片
        var artLayer = doc.layers.itemByName(config.layers.art.name);
        var placedCount = 0;
        var failedCount = 0;
        
        for (var j = 0; j < artFiles.length; j++) {
            var artFile = artFiles[j];
            var pageNum = extractPageNumber(artFile.name, pattern);
            
            // 确定目标页面
            var targetPage;
            if (config.imageImport.placeByFileName && pageNum > 0) {
                targetPage = doc.pages.itemByName(pageNum.toString());
                if (!targetPage.isValid) {
                    targetPage = doc.pages[Math.min(pageNum - 1, doc.pages.length - 1)];
                }
            } else {
                targetPage = doc.pages[j];
            }
            
            // 放置图片
            var placeResult = placeImageOnPageWithResult(artFile, targetPage, artLayer);
            if (placeResult.success) {
                placedCount++;
            } else {
                failedCount++;
                logMessage("放置图片失败 " + artFile.name + ": " + placeResult.message);
            }
        }
        
        artLayer.locked = true; // 锁定Art图层，防止误操作
        
        var msg = "图片导入完成: 成功 " + placedCount + " 个";
        if (failedCount > 0) {
            msg += ", 失败 " + failedCount + " 个";
        }
        
        // 如果全部失败，返回失败
        if (placedCount === 0 && artFiles.length > 0) {
            return new StepResult(false, "所有图片放置失败");
        }
        
        return new StepResult(true, msg, { placed: placedCount, failed: failedCount });
        
    } catch (e) {
        return new StepResult(false, "导入图片失败: " + e.message);
    }
}

// 找出文件名共同模式
function findCommonPattern(filenames) {
    if (filenames.length < 2) return null;
    // 找出共同前缀
    var prefix = filenames[0];
    for (var i = 1; i < filenames.length; i++) {
        var len = Math.min(prefix.length, filenames[i].length);
        var j = 0;
        while (j < len && prefix[j] === filenames[i][j]) j++;
        prefix = prefix.substring(0, j);
    }
    // 找出共同后缀
    var reversed = [];
    for (var i = 0; i < filenames.length; i++) {
        reversed.push(filenames[i].split('').reverse().join(''));
    }
    var revPrefix = reversed[0];
    for (var k = 1; k < reversed.length; k++) {
        var len2 = Math.min(revPrefix.length, reversed[k].length);
        var m = 0;
        while (m < len2 && revPrefix[m] === reversed[k][m]) m++;
        revPrefix = revPrefix.substring(0, m);
    }
    var suffix = revPrefix.split('').reverse().join('');
    return { prefix: prefix, suffix: suffix };
}

// 从文件名提取页码，支持从变化的数字中提取
function extractPageNumber(filename, pattern) {
    if (pattern) {
        var prefix = pattern.prefix;
        var suffix = pattern.suffix;
        if (filename.indexOf(prefix) === 0 && filename.lastIndexOf(suffix) === filename.length - suffix.length) {
            var numStr = filename.substring(prefix.length, filename.length - suffix.length);
            if (/^\d+$/.test(numStr)) {
                return parseInt(numStr, 10);
            }
        }
    }
    // fallback
    var regex = /(\d{2,4})(?=[a-zA-Z]?\.([A-Za-z]{3,4})$)/;
    var match = regex.exec(filename);
    if (match) {
        return parseInt(match[1], 10);
    }
    return 0;
}

// 在页面上放置图片（返回结果）
function placeImageOnPageWithResult(imageFile, page, layer) {
    try {
        // 查找或创建图框
        var frame = findOrCreateFrame(page, layer);
        
        // 放置图片
        frame.place(imageFile);
        
        // 应用缩放
        if (config.imageImport.scaleFactor !== 100) {
            var scaleFactor = config.imageImport.scaleFactor / 100;
            var graphic = frame.graphics[0];
            if (graphic && graphic.isValid) {
                var scaleMatrix = app.transformationMatrices.add({
                    horizontalScaleFactor: scaleFactor,
                    verticalScaleFactor: scaleFactor
                });
                var anchorPoint = getAnchorPoint(config.imageImport.anchorPoint);
                graphic.transform(CoordinateSpaces.INNER_COORDINATES, anchorPoint, scaleMatrix);
            }
        }
        
        // 应用适合选项
        applyFitOption(frame);
        
        return new StepResult(true, "放置成功", imageFile.name);
        
    } catch (e) {
        return new StepResult(false, e.message, imageFile.name);
    }
}

// 查找或创建图框
function findOrCreateFrame(page, layer) {
    // 查找现有图框
    var pageItems = page.allPageItems;
    for (var i = pageItems.length - 1; i >= 0; i--) {
        var item = pageItems[i];
        if (item instanceof Rectangle && item.itemLayer.name === layer.name) {
            return item;
        }
    }
    
    // 检查主页图框
    if (page.appliedMaster !== null && page.appliedMaster.isValid) {
        var masterItems = page.masterPageItems;
        for (var j = 0; j < masterItems.length; j++) {
            var masterItem = masterItems[j];
            if (masterItem instanceof Rectangle) {
                return masterItem.override(page);
            }
        }
    }
    
    // 创建新图框
    var bounds = getPageBoundsWithBleed(page);
    return page.rectangles.add({
        geometricBounds: bounds,
        itemLayer: layer
    });
}

// 获取页面边界（含出血）
function getPageBoundsWithBleed(page) {
    var pb = page.bounds;
    var bleedTop = config.documentSettings.bleedTop;
    var bleedBottom = config.documentSettings.bleedBottom;
    var bleedOutside = config.documentSettings.bleedOutside;
    var isRightPage = page.side === PageSideOptions.RIGHT_HAND;
    
    if (isRightPage) {
        return [pb[0] - bleedTop, pb[1], pb[2] + bleedBottom, pb[3] + bleedOutside];
    } else {
        return [pb[0] - bleedTop, pb[1] - bleedOutside, pb[2] + bleedBottom, pb[3]];
    }
}

// 获取锚点
function getAnchorPoint(anchorName) {
    var anchorMap = {
        "center": AnchorPoint.CENTER_ANCHOR,
        "topCenter": AnchorPoint.TOP_CENTER_ANCHOR,
        "topLeft": AnchorPoint.TOP_LEFT_ANCHOR,
        "topRight": AnchorPoint.TOP_RIGHT_ANCHOR,
        "bottomCenter": AnchorPoint.BOTTOM_CENTER_ANCHOR,
        "bottomLeft": AnchorPoint.BOTTOM_LEFT_ANCHOR,
        "bottomRight": AnchorPoint.BOTTOM_RIGHT_ANCHOR,
        "leftCenter": AnchorPoint.LEFT_CENTER_ANCHOR,
        "rightCenter": AnchorPoint.RIGHT_CENTER_ANCHOR
    };
    return anchorMap[anchorName] || AnchorPoint.CENTER_ANCHOR;
}

// 应用适合选项
function applyFitOption(frame) {
    try {
        var fitOption = config.imageImport.fitOptions;
        switch (fitOption) {
            case "fillProportionally":
                frame.fit(FitOptions.FILL_PROPORTIONALLY);
                break;
            case "frameToContent":
                frame.fit(FitOptions.FRAME_TO_CONTENT);
                break;
            case "contentToFrame":
                frame.fit(FitOptions.CONTENT_TO_FRAME);
                break;
            case "centerContent":
                frame.fit(FitOptions.CENTER_CONTENT);
                break;
            default:
                frame.fit(FitOptions.FILL_PROPORTIONALLY);
        }
    } catch (e) {
        // 忽略适合选项错误
    }
}

// ==================== LabelPlus文本导入函数 ====================
function importLabelPlusTextWithResult() {
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        var txtPath = config.textImport.lpTxtPath;
        var txtFile = new File(txtPath);
        
        if (!txtFile.exists) {
            return new StepResult(false, "LabelPlus文本文件不存在: " + txtPath);
        }
        
        // 读取并解析文本
        var entries = parseLabelPlusTxt(txtFile);
        logMessage("解析到 " + entries.length + " 个文本条目");
        
        if (entries.length === 0) {
            return new StepResult(false, "未解析到任何文本条目");
        }
        
        // 获取Text图层，或者“文字”图层，或者“汉化”图层
        var textLayer = doc.layers.itemByName(config.layers.text.name);
        if (!textLayer || !textLayer.isValid) {
            // 尝试获取“文字”图层
            textLayer = doc.layers.itemByName("文字");
            if (!textLayer || !textLayer.isValid) {
                // 尝试获取“汉化”图层
                textLayer = doc.layers.itemByName("汉化");
                if (!textLayer || !textLayer.isValid) {
                    return new StepResult(false, "文字图层不存在");
                }
            }
        }
        
        doc.activeLayer = textLayer;
        doc.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;
        
        // 插入文本
        var insertedCount = 0;
        var failedCount = 0;
        
        for (var i = 0; i < entries.length; i++) {
            var insertResult = insertTextEntryWithResult(entries[i], textLayer);
            if (insertResult.success) {
                insertedCount++;
            } else {
                failedCount++;
            }
        }
        
        var msg = "文本导入完成: 成功 " + insertedCount + " 个";
        if (failedCount > 0) {
            msg += ", 失败 " + failedCount + " 个";
        }
        
        return new StepResult(true, msg, { inserted: insertedCount, failed: failedCount });
        
    } catch (e) {
        return new StepResult(false, "导入LabelPlus文本失败: " + e.message);
    }
}

// 解析LabelPlus格式文本
function parseLabelPlusTxt(txtFile) {
    var entries = [];
    
    try {
        txtFile.encoding = "UTF-8";
        txtFile.open("r");
        var content = txtFile.read();
        txtFile.close();
        
        // 应用文本替换
        content = applyReplacements(content);
        
        var lines = content.split("\n");
        var currentPageName = "";
        var currentPageNum = 0;
        var pageOffset = config.textImport.pageOffset || 0;
        
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            
            // 匹配页面名
            var pageMatch = line.match(/>>>>>>>>\[(.*)\]<<<<<<<<$/);
            if (pageMatch) {
                currentPageName = pageMatch[1];
                currentPageNum = extractPageNumber(currentPageName, null);
                continue;
            }
            
            // 匹配台词编号、位置和分组
            var textMatch = line.match(/----------------\[(\d+)\]----------------\[(.*?),(.*?),(.*?)\]/);
            if (textMatch) {
                var pageNumber = parseInt(textMatch[1], 10);
                var baseX = parseFloat(textMatch[2]);
                var baseY = parseFloat(textMatch[3]);
                var group = parseInt(textMatch[4], 10);
                
                // 收集文本内容
                var textContent = [];
                for (var j = i + 1; j < lines.length; j++) {
                    if (lines[j].match(/----------------\[\d+\]----------------/) || 
                        lines[j].match(/>>>>>>>>\[(.*)\]<<<<<<<<$/)) {
                        break;
                    }
                    if (lines) {
                        textContent.push(lines[j]);
                    }
                }
                
                // 根据模式处理
                var isSingleLine = config.textImport.singleLineMode;
                try {
                    //从textContent[0]中提取大括号{}内的内容，包括大括号本身，包括多个大括号连用
                        var braceMatch="";
                        var braceMatchArr = textContent[0].match(/\{.*?\}/g);
                        if (braceMatchArr) {
                            braceMatch = braceMatchArr.join("");
                        }
                } catch (error) {                    
                }
                
                if (isSingleLine && textContent.length > 1) {
                    // 单行模式：每行创建一个文本框
                    for (var k = 0; k < textContent.length; k++) {
                        if (braceMatch){
                            if (k==0) {
                                var entry = {
                                    pageImage: currentPageNum + pageOffset,
                                    pageNumber: pageNumber,
                                    position: [baseX - k * 0.03, baseY + k * 0.03],
                                    group: group,
                                    text: textContent[k]
                                };
                                entries.push(entry);
                            }                    
                            if (k>0) {
                                var entry = {
                                    pageImage: currentPageNum + pageOffset,
                                    pageNumber: pageNumber,
                                    position: [baseX - k * 0.03, baseY + k * 0.03],
                                    group: group,
                                    text: braceMatch + textContent[k]
                                };
                                entries.push(entry);}
                        }else {
                            var entry = {
                                pageImage: currentPageNum + pageOffset,
                                pageNumber: pageNumber,
                                position: [baseX - k * 0.03, baseY + k * 0.03],
                                group: group,
                                text: textContent[k]
                            };
                            entries.push(entry);
                        }
                    }                                    
                } else {
                    // 多行模式：合并为一个文本框
                    entries.push({
                        pageImage: currentPageNum + pageOffset,
                        pageNumber: pageNumber,
                        position: [baseX, baseY],
                        group: group,
                        text: textContent.join("\n")
                    });
                }
                
                i += textContent.length;
            }
        }
        
    } catch (e) {
        logMessage("解析LabelPlus文本时出错: " + e.message);
    }
    
    return entries;
}

// 应用文本替换
function applyReplacements(content) {
    var replacements = config.textImport.replacements;
    if (!replacements) return content;
    
    for (var key in replacements) {
        if (replacements.hasOwnProperty(key) && key) {
            var regex = new RegExp(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
            content = content.replace(regex, replacements[key]);
        }
    }
    
    return content;
}

// 插入文本条目（返回结果）
function insertTextEntryWithResult(entry, layer) {
    try {
        // 获取目标页面
        var pageIndex = entry.pageImage - 1;
        if (pageIndex < 0 || pageIndex >= doc.pages.length) {
            return new StepResult(false, "页码超出范围: " + entry.pageImage);
        }
        
        var page = doc.pages[pageIndex];
        
        // 转换坐标
        var pageWidth = page.bounds[3] - page.bounds[1];
        var pageHeight = page.bounds[2] - page.bounds[0];
        var x = pageWidth * entry.position[0];
        var y = pageHeight * entry.position[1];
        
        // 获取文本框尺寸
        var frameWidth = config.textImport.textFrameSize.width;
        var frameHeight = config.textImport.textFrameSize.height;
        
        // 创建文本框
        var textFrame = page.textFrames.add({
            geometricBounds: [
                y, 
                x - frameWidth / 2, 
                y + frameHeight, 
                x + frameWidth / 2
            ]
        });
        
        textFrame.contents = entry.text;
        
        // 设置文本方向为垂直
        textFrame.parentStory.storyPreferences.storyOrientation = StoryHorizontalOrVertical.VERTICAL;
        
        // 如果溢流，自动调整大小
        if (textFrame.overflows) {
            textFrame.fit(FitOptions.FRAME_TO_CONTENT);
        }
        
        // 应用对象样式
        applyStyleRules(textFrame, entry);
        
        return new StepResult(true, "插入成功");
        
    } catch (e) {
        return new StepResult(false, e.message);
    }
}
    try{
        textFrame.parentStory.storyPreferences.storyOrientation = StoryHorizontalOrVertical.VERTICAL;
        
        // 如果溢流，自动调整大小
        if (textFrame.overflows) {
            textFrame.fit(FitOptions.FRAME_TO_CONTENT);
        }
        
        // 应用对象样式
        applyStyleRules(textFrame, entry);
        
    } catch (e) {
        logMessage("插入文本条目时出错: " + e.message);
    }

// 应用样式规则
function applyStyleRules(textFrame, entry) {
    try {
        var styleRules = config.textImport.styleRules;
        if (!styleRules || styleRules.length === 0) return;
        
        var appliedStyle = null;
        var text = textFrame.contents;
        
        // 遍历样式规则
        for (var i = 0; i < styleRules.length; i++) {
            var rule = styleRules[i];
            if (!rule.enabled) continue;
            
            // 默认匹配规则
            if (rule.match === "默认匹配") {
                appliedStyle = rule.style;
                continue;
            }
            
            // 检查文本是否包含匹配字符
            if (text.indexOf(rule.match) !== -1) {
                appliedStyle = rule.style;
                break;
            }
        }
        
        // 应用样式
        if (appliedStyle) {
            var objectStyle = doc.objectStyles.itemByName(appliedStyle);
            if (objectStyle.isValid) {
                textFrame.appliedObjectStyle = objectStyle;
            }
        }
        
    } catch (e) {
        logMessage("应用样式规则时出错: " + e.message);
    }
}

// ==================== 样式匹配函数 ====================
function runStyleMatchingWithResult() {
    var scriptPath = File($.fileName).parent.fsName;
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        // 读取字体映射配置文件
        var mapConfigFile;
        if (config.fontMapping.mapconfig) {
            mapConfigFile = new File(config.fontMapping.mapconfig);
        } else {
            mapConfigFile = new File(scriptPath + "../../../Library/字符样式匹配配置.ini");
        }    
        
        if (!mapConfigFile.exists) {
            return new StepResult(false, "字体映射配置文件不存在: " + mapConfigFile.fsName);
        }
        
        // 步骤1: 构建 [样式名称] -> [样式对象] 的哈希表
        var charStyleHashMap = buildCharStyleHashMap();
        var paraStyleHashMap = buildParaStyleHashMap();
        logMessage("字符样式数量: " + getObjectKeyCount(charStyleHashMap) + " 个");
        logMessage("段落样式数量: " + getObjectKeyCount(paraStyleHashMap) + " 个");
        
        // 步骤2: 构建 [字体名] -> [样式对象] 的哈希表
        var fontToStyleHashMap = buildFontToStyleHashMap(mapConfigFile, charStyleHashMap);
        logMessage("字体映射哈希表: " + getObjectKeyCount(fontToStyleHashMap) + " 个");
        
        // 获取所有文本框
        var textFrames = getAllTextFramesForStyleMatch();
        logMessage("找到 " + textFrames.length + " 个文本框需要处理");
        
        if (textFrames.length === 0) {
            return new StepResult(true, "没有找到需要样式匹配的文本框");
        }
        
        // 步骤3: 应用样式匹配
        var matchResult = applyStyleMappingOptimizedWithCount(textFrames, fontToStyleHashMap, paraStyleHashMap);
        
        // 清除花括号内容
        clearBracketsInTextFrames(textFrames);
        
        return new StepResult(true, "样式匹配完成: 处理 " + textFrames.length + " 个文本框, 匹配 " + matchResult.matched + " 个", matchResult);
        
    } catch (e) {
        return new StepResult(false, "执行样式匹配时出错: " + e.message);
    }
}

// 获取对象键数量
function getObjectKeyCount(obj) {
    var count = 0;
    for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
            count++;
        }
    }
    return count;
}

// 步骤1: 构建 [样式名称] -> [样式对象] 的哈希表（字符样式）
function buildCharStyleHashMap() {
    var styleHashMap = {};
    
    function collectStyles(styles, groupName) {
        for (var i = 0; i < styles.length; i++) {
            if (styles[i].constructor.name === "CharacterStyleGroup") {
                // 递归处理样式组
                collectStyles(styles[i].characterStyles, styles[i].name);
                collectStyles(styles[i].characterStyleGroups, styles[i].name);
            } else if (styles[i].constructor.name === "CharacterStyle") {
                // 构建样式名: "样式名(组名)" 或 "样式名"
                var styleName = styles[i].name;
                var fullName = groupName ? styleName + "(" + groupName + ")" : styleName;
                styleHashMap[fullName] = styles[i];
            }
        }
    }
    
    collectStyles(doc.characterStyles, "");
    collectStyles(doc.characterStyleGroups, "");
    
    // 输出所有字符样式名，便于调试
    var allStyleNames = [];
    for (var key in styleHashMap) {
        if (styleHashMap.hasOwnProperty(key)) {
            allStyleNames.push(key);
        }
    }
    
    return styleHashMap;
}

// 步骤1: 构建 [样式名称] -> [样式对象] 的哈希表（段落样式）
function buildParaStyleHashMap() {
    var styleHashMap = {};
    
    function collectStyles(styles, groupName) {
        for (var i = 0; i < styles.length; i++) {
            if (styles[i].constructor.name === "ParagraphStyleGroup") {
                collectStyles(styles[i].paragraphStyles, styles[i].name);
                collectStyles(styles[i].paragraphStyleGroups, styles[i].name);
            } else if (styles[i].constructor.name === "ParagraphStyle") {
                var styleName = styles[i].name;
                // 跳过基本段落样式
                if (styleName === "[基本段落]" || styleName === "[无]") continue;
                var fullName = groupName ? styleName + "(" + groupName + ")" : styleName;
                styleHashMap[fullName] = styles[i];
                // 同时存储纯样式名以便模糊匹配
                styleHashMap[styleName] = styles[i];
            }
        }
    }
    
    collectStyles(doc.paragraphStyles, "");
    collectStyles(doc.paragraphStyleGroups, "");
    
    return styleHashMap;
}

// 步骤2: 构建 [字体名] -> [样式对象] 的哈希表
function buildFontToStyleHashMap(configFile, charStyleHashMap) {
    var fontToStyleMap = {};
    var configData = {};
    
    // 读取配置文件
    try {
        configFile.encoding = "UTF-8";
        configFile.open("r");
        
        while (!configFile.eof) {
            var line = configFile.readln();
            
            // 去除行首尾空格
            while (line.charAt(0) === " " || line.charAt(0) === "\t") {
                line = line.substring(1);
            }
            while (line.length > 0 && 
                   (line.charAt(line.length - 1) === " " || line.charAt(line.length - 1) === "\t")) {
                line = line.substring(0, line.length - 1);
            }
            
            if (line === "" || line.charAt(0) === "#" || line.charAt(0) === ";") continue;
            
            var eq = line.indexOf("=");
            if (eq > 0) {
                var key = line.substring(0, eq);
                var val = line.substring(eq + 1);
                configData[key] = val;
            }
        }
        configFile.close();
        
    } catch (e) {
        logMessage("读取配置文件失败: " + e.message);
        return fontToStyleMap;
    }
    
    // 构建字体名 -> 样式对象 的映射
    var matchedCount = 0;
    var fuzzyMatchedCount = 0;
    var unmatchedStyles = [];
    var totalCount = 0;
    
    // 遍历configData中所有键，找出font_X格式但不包含font_style_的键
    for (var key in configData) {
        if (configData.hasOwnProperty(key) && 
            key.indexOf("font_") === 0 && 
            key.indexOf("font_style_") === -1) {
            var fontName = configData[key];
            var idx = key.substring(5); // 提取索引号
            var styleName = configData["font_style_" + idx];
            
            totalCount++;
            
            if (fontName && styleName && styleName !== "[无]") {
                // 通过样式名在哈希表中查找样式对象
                var styleObj = charStyleHashMap[styleName];
                
                if (styleObj) {
                    // 精确匹配成功
                    fontToStyleMap[fontName] = styleObj;
                    matchedCount++;
                } else {
                    // 尝试模糊匹配：在哈希表中查找包含该样式名的键
                    for (var hashKey in charStyleHashMap) {
                        if (charStyleHashMap.hasOwnProperty(hashKey)) {
                            // 检查配置的样式名是否是哈希键的子串，或反之
                            if (hashKey.indexOf(styleName) !== -1 || styleName.indexOf(hashKey) !== -1) {
                                styleObj = charStyleHashMap[hashKey];
                                fontToStyleMap[fontName] = styleObj;
                                fuzzyMatchedCount++;
                                break;
                            }
                        }
                    }
                    
                    if (!styleObj) {
                        unmatchedStyles.push(styleName);
                    }
                }
            }
        }
    }
    
    // logMessage("字体映射统计: 总计 " + totalCount + " 条, 精确匹配 " + matchedCount + " 条, 模糊匹配 " + fuzzyMatchedCount + " 条, 失败 " + unmatchedStyles.length + " 条");
    // if (unmatchedStyles.length > 0) {
    //     logMessage("未匹配的样式名列表: " + unmatchedStyles.join(", "));
    // }
    
    return fontToStyleMap;
}

// 获取所有文本框用于样式匹配
function getAllTextFramesForStyleMatch() {
    var textFrames = [];
    
    for (var i = 0; i < doc.pages.length; i++) {
        var page = doc.pages[i];
        var frames = getAllTextFramesFromPage(page);
        for (var j = 0; j < frames.length; j++) {
            textFrames.push(frames[j]);
        }
    }
    
    return textFrames;
}

// 获取页面中所有文本框，包括编组中的文本框（跳过隐藏图层）
function getAllTextFramesFromPage(page) {
    var textFrames = [];
    
    function collectTextFrames(item) {
        if (item.itemLayer && !item.itemLayer.visible) return;
        
        if (item.constructor.name === "TextFrame") {
            textFrames.push(item);
        } else if (item.constructor.name === "Group") {
            for (var i = 0; i < item.allPageItems.length; i++) {
                collectTextFrames(item.allPageItems[i]);
            }
        }
    }
    
    for (var i = 0; i < page.allPageItems.length; i++) {
        collectTextFrames(page.allPageItems[i]);
    }
    
    // 过滤掉空白文本框
    var validFrames = [];
    for (var i = 0; i < textFrames.length; i++) {
        if (textFrames[i].contents !== null && String(textFrames[i].contents).replace(/^\s+|\s+$/g, '') !== "") {
            validFrames.push(textFrames[i]);
        }
    }
    return validFrames;
}

// 步骤3: 应用样式匹配（使用哈希表直接查找，返回匹配数量）
function applyStyleMappingOptimizedWithCount(textFrames, fontToStyleHashMap, paraStyleHashMap) {
    var matchedCount = 0;
    try {
        for (var i = 0; i < textFrames.length; i++) {
            var story = textFrames[i].parentStory;
            if (!story) continue;
            var txt = story.contents;
            var frameMatched = false;
            
            // 从文本中提取字体名 {字体：xxx}
            var fontMatch = txt.match(/\{字体：([^}]+)\}/);
            if (fontMatch && fontMatch[1]) {
                var extractedFontName = fontMatch[1];
                
                // 直接通过哈希表查找样式对象
                var styleObj = fontToStyleHashMap[extractedFontName];
                
                // 如果精确匹配失败，尝试模糊匹配
                if (!styleObj) {
                    for (var configFontName in fontToStyleHashMap) {
                        if (fontToStyleHashMap.hasOwnProperty(configFontName)) {
                            if (extractedFontName.indexOf(configFontName) !== -1 || 
                                configFontName.indexOf(extractedFontName) !== -1) {
                                styleObj = fontToStyleHashMap[configFontName];
                                break;
                            }
                        }
                    }
                }
                
                // 应用字符样式
                if (styleObj) {
                    try {
                        story.appliedCharacterStyle = styleObj;
                        frameMatched = true;
                    } catch (e) {}
                }
            }
            
            // 处理字号-段落样式
            var sizeMatch = txt.match(/\{字号：([^}]+)\}/);
            if (sizeMatch && sizeMatch[1]) {
                var sizeName = sizeMatch[1];
                var sizeNum = parseFloat(sizeName);
                var baseFontSize = config.fontMapping.baseFontSize || 9;
                
                // 计算缩放后的字号
                if (!isNaN(sizeNum) && baseFontSize > 0) {
                    var scaledSize = sizeNum * (9 / baseFontSize);
                    
                    // 在段落样式哈希表中查找最接近的样式
                    var bestStyle = findClosestStyleInHashMap(paraStyleHashMap, scaledSize);
                    if (bestStyle) {
                        try {
                            story.appliedParagraphStyle = bestStyle;
                            frameMatched = true;
                        } catch (e) {}
                    }
                }
            }
            
            if (frameMatched) {
                matchedCount++;
            }
        }
        
    } catch (e) {
        logMessage("应用样式映射时出错: " + e.message);
    }
    return { matched: matchedCount, total: textFrames.length };
}

// 在段落样式哈希表中查找最接近指定字号的样式
function findClosestStyleInHashMap(styleHashMap, targetSize) {
    var best = null;
    var minDiff = 99999;
    
    for (var styleName in styleHashMap) {
        if (styleHashMap.hasOwnProperty(styleName)) {
            // 从样式名中提取字号（假设样式名以数字开头，如 "9.2" 或 "10(宋体)"）
            var numMatch = styleName.match(/^[\d.]+/);
            if (numMatch) {
                var styleSize = parseFloat(numMatch[0]);
                if (!isNaN(styleSize)) {
                    var diff = Math.abs(styleSize - targetSize);
                    if (diff < minDiff) {
                        minDiff = diff;
                        best = styleHashMap[styleName];
                    }
                }
            }
        }
    }
    
    return best;
}

// 清除文本框中的花括号内容
function clearBracketsInTextFrames(textFrames) {
    try {
        var savedFindPrefs = app.findGrepPreferences.properties;
        var savedChangePrefs = app.changeGrepPreferences.properties;

        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;

        app.findGrepPreferences.findWhat = "\\{.*?\\}";
        app.changeGrepPreferences.changeTo = "";

        for (var i = 0; i < textFrames.length; i++) {
            var story = textFrames[i].parentStory;
            if (!story) continue;
            story.changeGrep();
        }

        app.findGrepPreferences.properties = savedFindPrefs;
        app.changeGrepPreferences.properties = savedChangePrefs;
    } catch (e) {}
}

// 模糊匹配字体名
function isFontMatch(text, configFontName) {
    // 如果文本中直接包含配置的字体名，直接返回 true
    if (text.indexOf(configFontName) !== -1) {
        return true;
    }
    
    var fontPattern = "{字体：" + configFontName + "}";
    if (text.indexOf(fontPattern) !== -1) {
        return true;
    }
    
    // 模糊匹配逻辑：配置文件中的字体名是实际字体名的子串
    // 先提取文本中所有花括号内的字体声明
    var fontMatches = text.match(/\{字体：([^}]+)\}/g);
    if (fontMatches) {
        for (var i = 0; i < fontMatches.length; i++) {
            // 提取花括号内的字体名
            var extractedFont = fontMatches[i].replace(/\{字体：/, "").replace(/\}/, "");
            // 检查配置的字体名是否是提取字体名的子串，或反之
            if (extractedFont.indexOf(configFontName) !== -1 || configFontName.indexOf(extractedFont) !== -1) {
                return true;
            }
        }
    }
    
    return false;
}

// 查找段落样式
// 支持格式: "样式名" 或 "样式名(分组名)"
function findParagraphStyle(styleName) {
    try {
        // 解析样式名和分组名
        var groupName = "";
        var actualStyleName = styleName;
        var parenIdx = styleName.lastIndexOf("(");
        if (parenIdx > 0 && styleName.charAt(styleName.length - 1) === ")") {
            groupName = styleName.substring(parenIdx + 1, styleName.length - 1);
            actualStyleName = styleName.substring(0, parenIdx);
        }
        
        // 如果有分组名，直接在该分组中查找
        if (groupName) {
            var group = doc.paragraphStyleGroups.itemByName(groupName);
            if (group.isValid) {
                var style = group.paragraphStyles.itemByName(actualStyleName);
                if (style.isValid) return style;
            }
        }
        
        // 尝试直接查找（无分组）
        var style = doc.paragraphStyles.itemByName(actualStyleName);
        if (style.isValid) return style;
        
        // 遍历所有分组查找
        return findStyleInGroups(doc.paragraphStyleGroups, actualStyleName, groupName, "paragraph");
    } catch (e) {
        return null;
    }
}

// 在样式组中递归查找样式
// targetGroupName: 目标分组名，如果为空则在所有分组中查找
function findStyleInGroups(groups, styleName, targetGroupName, styleType) {
    if (!groups || groups.length === 0) return null;
    
    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        
        // 如果指定了目标分组名，只在该分组中查找
        if (targetGroupName && group.name !== targetGroupName) {
            continue;
        }
        
        // 在当前组中查找
        var styles = styleType === "character" ? group.characterStyles : group.paragraphStyles;
        for (var j = 0; j < styles.length; j++) {
            if (styles[j].name === styleName) {
                return styles[j];
            }
        }
        
        // 递归查找子组
        var subGroups = styleType === "character" ? group.characterStyleGroups : group.paragraphStyleGroups;
        var found = findStyleInGroups(subGroups, styleName, targetGroupName, styleType);
        if (found) return found;
    }
    
    return null;
}

// 查找最接近的段落样式（按字号）
function findClosestParagraphStyle(targetSize) {
    try {
        var best = null;
        var minDiff = 99999;
        
        // 遍历所有段落样式
        var styles = doc.paragraphStyles;
        for (var i = 0; i < styles.length; i++) {
            var style = styles[i];
            var styleName = style.name;
            
            // 跳过基本段落样式
            if (styleName === "[基本段落]" || styleName === "[无]") continue;
            
            // 从样式名中提取字号
            var numMatch = styleName.match(/^[\d.]+/);
            if (numMatch) {
                var styleSize = parseFloat(numMatch[0]);
                if (!isNaN(styleSize)) {
                    var diff = Math.abs(styleSize - targetSize);
                    if (diff < minDiff) {
                        minDiff = diff;
                        best = style;
                    }
                }
            }
        }
        
        // 也检查段落样式组
        best = findClosestStyleInGroups(doc.paragraphStyleGroups, targetSize, best, minDiff);
        
        return best;
    } catch (e) {
        return null;
    }
}

// 在段落样式组中查找最接近的样式
function findClosestStyleInGroups(groups, targetSize, currentBest, currentMinDiff) {
    if (!groups || groups.length === 0) return currentBest;
    
    var best = currentBest;
    var minDiff = currentMinDiff;
    
    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        var styles = group.paragraphStyles;
        
        for (var j = 0; j < styles.length; j++) {
            var styleName = styles[j].name;
            if (styleName === "[基本段落]" || styleName === "[���]") continue;
            
            var numMatch = styleName.match(/^[\d.]+/);
            if (numMatch) {
                var styleSize = parseFloat(numMatch[0]);
                if (!isNaN(styleSize)) {
                    var diff = Math.abs(styleSize - targetSize);
                    if (diff < minDiff) {
                        minDiff = diff;
                        best = styles[j];
                    }
                }
            }
        }
        
        // 递归子组
        best = findClosestStyleInGroups(group.paragraphStyleGroups, targetSize, best, minDiff);
    }
    
    return best;
}

// 清除文本框中的花括号内容
function clearBracketsInTextFrames(textFrames) {
    try {
        // 保存当前查找/替换设置
        var savedFindPrefs = app.findGrepPreferences.properties;
        var savedChangePrefs = app.changeGrepPreferences.properties;
        
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        
        app.findGrepPreferences.findWhat = "\\{.*?\\}";
        app.changeGrepPreferences.changeTo = "";
        
        for (var i = 0; i < textFrames.length; i++) {
            var story = textFrames[i].parentStory;
            if (story) {
                try {
                    story.changeGrep();
                } catch (e) {}
            }
        }
        
        // 恢复查找/替换设置
        app.findGrepPreferences.properties = savedFindPrefs;
        app.changeGrepPreferences.properties = savedChangePrefs;
        
    } catch (e) {
        logMessage("清除花括号时出错: " + e.message);
    }
}

// ==================== 断句函数 ====================
function runSegmentationWithResult() {
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        var range = config.segmentation.range;
        var textFrames = [];
        
        // 获取需要处理的文本框
        if (range === "currentSelection") {
            for (var i = 0; i < app.selection.length; i++) {
                if (app.selection[i].constructor.name === "TextFrame") {
                    textFrames.push(app.selection[i]);
                }
            }
        } else if (range === "currentPage") {
            var currentPage = app.activeWindow.activePage;
            textFrames = currentPage.textFrames.everyItem().getElements();
        } else {
            // entireDocument
            var allTextFrames = doc.textFrames.everyItem().getElements();
            for (var j = 0; j < allTextFrames.length; j++) {
                var tf = allTextFrames[j];
                try {
                    if (tf.parentPage != null && 
                        tf.parentPage.parent.constructor.name !== "MasterSpread") {
                        textFrames.push(tf);
                    }
                } catch (e) {
                    // 忽略错误
                }
            }
        }
        
        if (textFrames.length === 0) {
            return new StepResult(true, "没有找到需要断句的文本框");
        }
        
        // 调用Python脚本进行断句
        var segResult = callPythonSegmentationWithResult(textFrames);
        return segResult;
        
    } catch (e) {
        return new StepResult(false, "执行断句时出错: " + e.message);
    }
}

// 调用Python断句脚本（返回结果）
function callPythonSegmentationWithResult(textFrames) {
    try {
        // 创建临时文件
        var tempFolder = Folder.temp;
        var inputFile = new File(tempFolder.fsName + "/jieba_temp_input.txt");
        var outputFile = new File(tempFolder.fsName + "/jieba_temp_output.txt");
        
        // 写入文本内容
        inputFile.encoding = "UTF-8";
        inputFile.open("w");
        for (var i = 0; i < textFrames.length; i++) {
            var content = textFrames[i].parentStory.contents;
            content = content.replace(/[\r\n]/g, "");
            inputFile.write(content + "\n");
        }
        inputFile.close();
        
        // 获取Python脚本路径
        var scriptPath = File($.fileName).parent.fsName;
        var pythonScript;
        if (config.segmentation.pythonScriptPath) {
            pythonScript = new File(config.segmentation.pythonScriptPath);
        } else {
            pythonScript = new File(scriptPath + "../../../断句-CN/jieba_pytojs.pyw");
        }
        
        // 检查Python脚本是否存在
        if (!pythonScript.exists) {
            // 清理临时文件
            if (inputFile.exists) inputFile.remove();
            return new StepResult(false, "Python断句脚本不存在: " + pythonScript.fsName);
        }
        
        // 执行Python脚本
        pythonScript.execute('"' + inputFile.fsName + '" "' + outputFile.fsName + '"');
        
        // 等待输出文件（最多30秒）
        var maxWait = 30000;
        var waited = 0;
        while (!outputFile.exists && waited < maxWait) {
            delay(500);
            waited += 500;
        }
        
        if (!outputFile.exists) {
            // 清理临时文件
            if (inputFile.exists) inputFile.remove();
            return new StepResult(false, "Python断句脚本执行超时（30秒）");
        }
        
        // 读取结果
        outputFile.encoding = "UTF-8";
        outputFile.open("r");
        var segmentedContent = outputFile.read();
        outputFile.close();
        
        // 应用结果
        var contentLines = segmentedContent.split("\n");
        var processedCount = 0;
        for (var j = 0; j < textFrames.length && j < contentLines.length; j++) {
            var processedContent = contentLines[j].replace(/\\r/g, "\n");
            textFrames[j].parentStory.contents = processedContent || '';
            processedCount++;
        }
        
        // 清理临时文件
        if (inputFile.exists) inputFile.remove();
        if (outputFile.exists) outputFile.remove();
        
        return new StepResult(true, "断句完成: 处理 " + processedCount + " 个文本框", { processed: processedCount });
        
    } catch (e) {
        // 清理临时文件
        try {
            var tempFolder = Folder.temp;
            var inputFile = new File(tempFolder.fsName + "/jieba_temp_input.txt");
            var outputFile = new File(tempFolder.fsName + "/jieba_temp_output.txt");
            if (inputFile.exists) inputFile.remove();
            if (outputFile.exists) outputFile.remove();
        } catch (cleanErr) {}
        
        return new StepResult(false, "调用Python断句脚本时出错: " + e.message);
    }
}

// ==================== 文档保存函数 ====================
function saveDocumentWithResult() {
    try {
        if (!doc || !doc.isValid) {
            return new StepResult(false, "文档对象无效");
        }
        
        var saveFile = null;
        var saveFolder = null;
        var filePrefix = "";
        
        // 1. 尝试从图片文件夹获取保存路径
        if (config.imageImport.enabled && config.imageImport.artFolderPath) {
            saveFolder = new Folder(config.imageImport.artFolderPath);
        }
        
        // 2. 尝试从翻译稿文件名获取前缀
        if (config.textImport.enabled && config.textImport.lpTxtPath) {
            var txtFile = new File(config.textImport.lpTxtPath);
            if (txtFile.exists) {
                // 提取文件名前缀（去掉扩展名和可能的 _LPoutput_数据 等后缀）
                var txtName = txtFile.name;
                // 移除扩展名
                var dotIdx = txtName.lastIndexOf(".");
                if (dotIdx > 0) {
                    txtName = txtName.substring(0, dotIdx);
                }
                // 移除常见的后缀模式
                txtName = txtName.replace(/_LPoutput_数据$/, "");
                txtName = txtName.replace(/_LPoutput$/, "");
                txtName = txtName.replace(/_LP$/, "");
                txtName = txtName.replace(/\(校对\)$/, "");
                txtName = txtName.replace(/\(翻译\)$/, "");
                filePrefix = txtName;
            }
        }
        
        // 3. 确定最终保存路径
        if (saveFolder && saveFolder.exists && filePrefix) {
            // 优先保存到图片文件夹，使用翻译稿文件名前缀
            var savePath = saveFolder.fsName;
            if ($.os.indexOf("Windows") !== -1) {
                saveFile = new File(savePath + "\\" + filePrefix + ".indd");
            } else {
                saveFile = new File(savePath + "/" + filePrefix + ".indd");
            }
            logMessage("保存到图片文件夹: " + saveFile.fsName);
        } else if (config.output.savePath) {
            // 使用配置中的保存路径
            saveFile = new File(config.output.savePath);
            logMessage("使用配置中的保存路径: " + saveFile.fsName);
        } else {
            // fallback 到 output.indd（脚本所在目录）
            var scriptPath = File($.fileName).parent.fsName;
            if ($.os.indexOf("Windows") !== -1) {
                saveFile = new File(scriptPath + "\\output.indd");
            } else {
                saveFile = new File(scriptPath + "/output.indd");
            }
            logMessage("使用默认保存路径: " + saveFile.fsName);
        }
        
        // 保存文档
        doc.save(saveFile);
        logMessage("文档已保存: " + saveFile.fsName);
        
        var savedFiles = [saveFile.fsName];
        
        // 如果需要保存IDML
        if (config.output.saveAsIdml) {
            var idmlPath = saveFile.fsName.replace(/\.indd$/, ".idml");
            doc.exportFile(ExportFormat.INDESIGN_MARKUP, new File(idmlPath));
            logMessage("IDML已导出: " + idmlPath);
            savedFiles.push(idmlPath);
        }
        
        // 如果需要关闭文档
        if (config.output.closeAfterSave) {
            doc.close();
            doc = null;
        }
        
        return new StepResult(true, "文档保存成功: " + saveFile.fsName, { files: savedFiles });
        
    } catch (e) {
        return new StepResult(false, "保存文档失败: " + e.message);
    }
}

// ==================== 工具函数 ====================
// 延时函数
function delay(ms) {
    if (!ms || ms <= 0) return;
    
    var start = new Date().getTime();
    var now = start;
    while (now - start < ms) {
        now = new Date().getTime();
    }
}

// 日志记录函数
function logMessage(message) {
    var timestamp = new Date().toLocaleTimeString();
    var logEntry = "[" + timestamp + "] " + message;
    logMessages.push(logEntry);
    $.writeln(logEntry);
}

// 写入日志文件
function writeLogFile() {
    try {
        var logPath = config.output.logPath;
        if (!logPath) {
            var desktopPath = Folder.desktop.fsName;
            var timestamp = new Date().toISOString().replace(/[:.]/g, "-").substring(0, 19);
            logPath = desktopPath + "\\manga_layout_log_" + timestamp + ".txt";
            if ($.os.indexOf("Windows") === -1) {
                logPath = desktopPath + "/manga_layout_log_" + timestamp + ".txt";
            }
        }
        
        var logFile = new File(logPath);
        logFile.encoding = "UTF-8";
        logFile.open("w");
        
        // 写入所有日志消息
        logFile.write(logMessages.join("\n"));
        
        // 如果有错误，额外写入错误信息
        if (errorMessages.length > 0) {
            logFile.write("\n\n===== 错误信息 =====\n");
            logFile.write(errorMessages.join("\n"));
        }
        
        logFile.close();
        
    } catch (e) {
        $.writeln("写入日志文件失败: " + e.message);
    }
}

// ==================== 执行主函数 ====================
var result = main();

// 返回执行结果（供VBS调用）
result;
