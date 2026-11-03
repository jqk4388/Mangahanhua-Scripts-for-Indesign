// 判断文本框是否应跳过：已锁定、对象隐藏或所在图层隐藏
function shouldSkipFrame(tf) {
    try {
        if (tf.locked) return true
        if (tf.visible === false) return true
        if (tf.itemLayer && tf.itemLayer.visible === false) return true
    } catch (e) {
        return true // 属性读取失败时也跳过，避免中断脚本
    }
    return false
}

// 查找或删除空白文本框：弹出对话框询问是否删除所有空白文本框，或仅查找并定位第一个空白文本框。
function findOrDeleteEmptyTextFrames() {
    if (!app.documents.length) {
        alert('当前没有打开的文档。')
        return false
    }

    var doc = app.activeDocument
    var deleteAll = confirm('是否删除所有空白文本框？\n点击“确定”将删除所有空白文本框；点击“取消”只查找并定位第一个空白文本框。')

    if (deleteAll) {
        // 倒序遍历 frames，以便安全删除
        var deletedCount = 0
        for (var p = 0; p < doc.pages.length; p++) {
            var frames = doc.pages[p].textFrames
            for (var f = frames.length - 1; f >= 0; f--) {
                try {
                    var tf = frames[f]
                    if (!shouldSkipFrame(tf) && tf.contents === '' && tf.overflows === false) {
                        tf.remove()
                        deletedCount++
                    }
                } catch (e) {
                    // 忽略删除错误，继续处理其他文本框
                }
            }
        }
        if (deletedCount > 0) {
            alert('已删除 ' + deletedCount + ' 个空白文本框。')
            return true
        } else {
            alert('未找到空白文本框。')
            return false
        }
    } else {
        // 仅查找并定位第一个空白文本框
        for (var i = 0; i < doc.pages.length; i++) {
            var page = doc.pages[i]
            for (var j = 0; j < page.textFrames.length; j++) {
                try {
                    var tf2 = page.textFrames[j]
                    if (!shouldSkipFrame(tf2) && tf2.contents === '' && tf2.overflows === false) {
                        app.selection = tf2
                        alert('找到空白文本框，位于页：' + page.name + '。')
                        return true
                    }
                } catch (e) {
                    // 跳过无法读取的文本框
                }
            }
        }
        alert('未找到空白文本框。')
        return false
    }
}

findOrDeleteEmptyTextFrames()
