function getTextStructure () {
    var dialogPages = []; // dialogPages will be an array of arrays. Each sub-array will contain zero or more strings, each string corresponding to a single dialog bubble
    for (var i = 0; i < app.activeDocument.pages.length; i++) { // loop over pages
        dialogPages.push([]); // Add an empty array for every page we iterate over. 
        for (var j = 0; j < app.activeDocument.pages[i].textFrames.length; j++) { // loop over textFrames
            if (app.activeDocument.pages[i].textFrames[j].itemLayer.name == 'Text') { // Check if the textFrame is on the 'Dialog' layer
                dialogPages[i].push(app.activeDocument.pages[i].textFrames[j].contents);
            }
        }
    }
    return dialogPages;
}

var myDialog = getTextStructure();

function findFirstTextFrameWithSpace() {
    for (var i = 0; i < app.activeDocument.pages.length; i++) { // Iterate through each page
        var page = app.activeDocument.pages[i];
        if (page.textFrames.length > 0) { // Check if the page has any text frames
            var textFrame = page.textFrames.lastItem();
            if (textFrame && !textFrame.locked) { // Ensure text frame is valid and unlocked
                return textFrame;
            }
        }
    }
    return null; // Return null if no suitable text frame is found
}

var targetTextFrame = findFirstTextFrameWithSpace();
targetTextFrame.contents +="文档中的全部文本："
if (targetTextFrame) {
    for (var i = 0; i < myDialog.length; i++) {
        for (var j = 0; j < myDialog[i].length; j++) {
            targetTextFrame.contents += myDialog[i][j] + "\n";
        }
    }
} else {
    alert("Error: No suitable text frame found on any page.");
}
