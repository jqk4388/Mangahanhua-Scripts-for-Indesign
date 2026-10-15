#include "../Library/KTUlib.jsx"

// 收集页面上可被选择的顶层元素
// 排除：本体锁定/隐藏的元素，以及所在图层锁定/隐藏的元素
function collectSelectablePageItems(page) {
    var result = [];
    var items = page.pageItems;
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var layer = item.itemLayer;
        if (item.locked) continue;       // 本体锁定
        if (!item.visible) continue;     // 本体隐藏
        if (layer.locked) continue;      // 所在图层锁定
        if (!layer.visible) continue;    // 所在图层隐藏
        result.push(item);
    }
    return result;
}

// function for selecting all items on right page

function selectAllOnRightPage() {
    if (theDoc.documentPreferences.pageBinding == PageBindingOptions.leftToRight && app.activeWindow.activeSpread.pages.length > 1) { // if the book's laid out left-to-right AND there are multiple pages on the spread
        app.selection = collectSelectablePageItems(app.activeWindow.activeSpread.pages[1]) // select everything on the first page of the spread
    } else {
        app.selection = collectSelectablePageItems(app.activeWindow.activeSpread.pages[0]) // otherwise, select everything on the second page of the spread
    }
}

// Send script to KTUDoScriptAsUndoable, just for form's sake,
// and do it within a try/catch, again, just for form's sake
try { 
    KTUDoScriptAsUndoable(selectAllOnRightPage, "Select All On Right Page")
} catch(err) {
    alert("Error: " + err.description) 
}
