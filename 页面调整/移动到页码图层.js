// Get the active document
var doc = app.activeDocument;

// Check if the "页码" layer exists, if not create a new layer named "Page Numbers"
var targetLayer;
try {
    targetLayer = doc.layers.itemByName("Page Numbers");
    targetLayer.name; // This will throw an error if the layer does not exist
} catch (e) {
    targetLayer = doc.layers.add({ name: "Page Numbers" });
}

// Move the selected text frames to the target layer
var selectedItems = app.selection;

// If no items are selected, find text using Grep expression ~N and move to target layer
if (selectedItems.length === 0) {
    // Enable Grep search
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = "~N"; // Grep expression for page numbers

    // Find all matching text
    var foundItems = doc.findGrep();

    // Move found text items to target layer
    for (var i = 0; i < foundItems.length; i++) {
        if (foundItems[i] instanceof TextFrame) {
            foundItems[i].itemLayer = targetLayer;
        } else if (foundItems[i] instanceof Character || foundItems[i] instanceof Word || foundItems[i] instanceof Line || foundItems[i] instanceof Text) {
            // If it's text within a text frame, move the parent text frame
            if (foundItems[i].parentTextFrames.length > 0) {
                foundItems[i].parentTextFrames[0].itemLayer = targetLayer;
            }
        }
    }

    // Clear find preferences
    app.findGrepPreferences = NothingEnum.nothing;
} else {
    // Move the selected text frames to the target layer
    for (var i = 0; i < selectedItems.length; i++) {
        if (selectedItems[i] instanceof TextFrame) {
            selectedItems[i].itemLayer = targetLayer;
        }
    }
}