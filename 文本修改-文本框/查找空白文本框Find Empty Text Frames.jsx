// 查找或删除空白文本框：弹出对话框询问是否删除所有空白文本框，或仅查找并定位第一个空白文本框。
function findOrDeleteEmptyTextFramesˀ() {
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
                    if (tf.contents === '' && tf.overflows === false) {
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
                var tf2 = page.textFrames[j]
                if (tf2.contents === '' && tf2.overflows === false) {
                    app.selection = tf2
                    alert('找到空白文本框，位于页：' + page.name + '。')
                    return true
                }
            }
        }
        alert('未找到空白文本框。')
        return false
    }
}

findOrDeleteEmptyTextFramesˀ()
