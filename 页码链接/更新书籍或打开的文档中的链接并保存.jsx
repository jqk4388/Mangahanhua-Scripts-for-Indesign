// 检查是否有打开的书籍
if (app.books.length === 0) {
    // 没有打开书籍，遍历所有已打开文档
    if (app.documents.length === 0) {
        alert("没有打开的文档或书籍。");
    } else {
        for (var d = 0; d < app.documents.length; d++) {
            var doc = app.documents[d];
            var links = doc.links;
            for (var j = 0; j < links.length; j++) {
                var link = links[j];
                if (link.status === LinkStatus.LINK_OUT_OF_DATE || link.status === LinkStatus.LINK_MISSING) {
                    try {
                        link.update();
                    } catch (e) {
                        // 忽略无法更新的链接
                    }
                }
            }
            doc.save();
        }
        alert("所有文档链接已更新并保存。");
    }
} else {
    var book = app.activeBook;
    var bookContent, doc, i, j, links, link;

    // 遍历书籍中的所有文档
    for (i = 0; i < book.bookContents.length; i++) {
        bookContent = book.bookContents[i];
        // 只处理已存在的文档
        if (bookContent.status === "INCLUDED") {
            try {
                doc = app.open(File(bookContent.fullName), false); // 不显示界面
                // 更新所有链接
                links = doc.links;
                for (j = 0; j < links.length; j++) {
                    link = links[j];
                    if (link.status === LinkStatus.LINK_OUT_OF_DATE || link.status === LinkStatus.LINK_MISSING) {
                        try {
                            link.update();
                        } catch (e) {
                            // 忽略无法更新的链接
                        }
                    }
                }
                doc.save();
                doc.close();
            } catch (e) {
                // 忽略无法打开或保存的文档
            }
        }
    }
    // 保存书籍
    try {
        book.save();
    } catch (e) {
        // 忽略保存书籍时的错误
    }
    alert("所有文档链接已更新并保存，书籍已保存。");
}
