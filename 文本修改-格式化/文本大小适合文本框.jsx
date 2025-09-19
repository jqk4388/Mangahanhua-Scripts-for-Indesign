#target indesign

// 文本大小适合文本框
// 作者: 几千块
// 功能: 自动调整文本属性以适应固定尺寸的文本框

var TextFitter = {
    // 默认配置参数
    config: {
        minFontSize: 6,
        maxFontSize: 30,
        lineHeightAdjustRange: 0.2, // ±20%
        trackingAdjustRange: 0.1,   // ±10%
        maxFontSizeForShortText: 36 // 短文本最大字号限制
    },
    
    // 主函数
    main: function() {
        var doc = app.activeDocument;
        var selection = doc.selection;
        
        if (selection.length === 0) {
            alert("请选择至少一个文本框");
            return;
        }
        
        var textFrames = this.getTextFramesFromSelection(selection);
        
        if (textFrames.length === 0) {
            alert("未找到可处理的文本框");
            return;
        }
        
        // 显示预估信息
        // if (!this.showPreview(textFrames)) {
        //     return;
        // }
        
        // 开始处理
        app.scriptPreferences.enableRedraw = false;
        appUndoSupportLevel =UndoModes.FAST_ENTIRE_SCRIPT;
        
        var results = {
            success: 0,
            failed: 0,
            errors: []
        };
        
        for (var i = 0; i < textFrames.length; i++) {
            try {
                this.processTextFrame(textFrames[i]);
                results.success++;
            } catch (e) {
                results.failed++;
                results.errors.push({
                    index: i,
                    name: textFrames[i].name || "文本框" + (i+1),
                    error: e.message
                });
            }
        }
        
        app.scriptPreferences.enableRedraw = true;
        
        // 显示结果
        this.showResults(results);
    },
    
    // 从选择中提取文本框（包括嵌套在群组中的）
    getTextFramesFromSelection: function(selection) {
        var textFrames = [];
        
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            
            switch (item.constructor.name) {
                case "TextFrame":
                    if (!item.locked) {
                        textFrames.push(item);
                    }
                    break;
                    
                case "Group":
                    textFrames = textFrames.concat(this.getTextFramesFromGroup(item));
                    break;
                    
                case "PageItem":
                    // 检查是否包含文本内容
                    if (item.hasOwnProperty("texts") && !item.locked) {
                        textFrames.push(item);
                    }
                    break;
            }
        }
        
        return textFrames;
    },
    
    // 从群组中提取文本框
    getTextFramesFromGroup: function(group) {
        var textFrames = [];
        var allPageItems = group.allPageItems;
        
        for (var i = 0; i < allPageItems.length; i++) {
            var item = allPageItems[i];
            if (item instanceof TextFrame && !item.locked) {
                textFrames.push(item);
            }
        }
        
        return textFrames;
    },
    
    // 显示预估信息
    showPreview: function(textFrames) {
        var info = "将处理 " + textFrames.length + " 个文本框\n\n";
        info += "调整策略:\n";
        info += "- 优先调整字号 (" + this.config.minFontSize + "pt - " + this.config.maxFontSize + "pt)\n";
        info += "- 字号达极限后调整行间距(±" + (this.config.lineHeightAdjustRange * 100) + "%)\n";
        info += "- 行间距达极限后调整字间距(±" + (this.config.trackingAdjustRange * 100) + "%)\n\n";
        info += "确定要继续吗？";
        
        return confirm(info);
    },
    
    // 处理单个文本框
    processTextFrame: function(textFrame) {
        // 检查文本框是否为空
        if (textFrame.contents === "" || textFrame.contents === null) {
            throw new Error("文本框为空");
        }
        
        // 保存原始状态用于撤销
        var originalState = this.saveState(textFrame);
        
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
    
    // 保存当前状态
    saveState: function(textFrame) {
        return {
            fontSize: this.getFontSize(textFrame),
            leading: this.getLeading(textFrame),
            tracking: this.getTracking(textFrame)
        };
    },
    
    // 获取字体大小
    getFontSize: function(textFrame) {
        try {
            return textFrame.texts[0].pointSize;
        } catch (e) {
            return 12; // 默认值
        }
    },
    
    // 获取行距
    getLeading: function(textFrame) {
        try {
            var leading = textFrame.texts[0].leading;
            return (leading == Leading.AUTO) ? this.getFontSize(textFrame) * 1.2 : leading;
        } catch (e) {
            return this.getFontSize(textFrame) * 1.2; // 默认值
        }
    },
    
    // 获取字间距
    getTracking: function(textFrame) {
        try {
            return textFrame.texts[0].tracking;
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
            textFrame.texts[0].leading = leading;
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
    },
    
    // 显示处理结果
    showResults: function(results) {
        var message = "文本适配完成\n\n";
        message += "成功: " + results.success + " 个文本框\n";
        message += "失败: " + results.failed + " 个文本框\n\n";
        
        if (results.errors.length > 0) {
            message += "错误详情:\n";
            for (var i = 0; i < Math.min(results.errors.length, 5); i++) {
                var error = results.errors[i];
                message += "- " + error.name + ": " + error.error + "\n";
            }
            
            if (results.errors.length > 5) {
                message += "... 还有 " + (results.errors.length - 5) + " 个错误\n";
            }
        }
        
        // alert(message);
    }
};

// 运行主函数
try {
    TextFitter.main();
} catch (e) {
    alert("脚本执行出错: " + e.message);
}