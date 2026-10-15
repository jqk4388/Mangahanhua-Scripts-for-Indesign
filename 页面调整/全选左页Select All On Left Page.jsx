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

// function for selecting all items on left page
// have to declare function as a variable in order to pass it, jeez
var selectAllOnLeftPage = function () {
    if (app.activeWindow.activeSpread.pages.length == 1) { // if there's only one page in the spread
        app.selection = collectSelectablePageItems(app.activeWindow.activeSpread.pages[0]) // select everything on the first page of the spread
    } else { // if there are multiple pages in the spread
        if (theDoc.documentPreferences.pageBinding == PageBindingOptions.leftToRight) { // if the book's laid out left-to-right OR there's only one page on the spread
            app.selection = collectSelectablePageItems(app.activeWindow.activeSpread.pages[0]) // select everything on the first page of the spread
        } else {
            app.selection = collectSelectablePageItems(app.activeWindow.activeSpread.pages[1]) // otherwise, select everything on the second page of the spread
        }
    }
}

// Send script to KTUDoScriptAsUndoable, just for form's sake,
// and do it within a try/catch, again, just for form's sake
try { 
    KTUDoScriptAsUndoable(selectAllOnLeftPage, "Select All On Left Page")
} catch(err) {
    alert("Error: " + err.description) 
}
