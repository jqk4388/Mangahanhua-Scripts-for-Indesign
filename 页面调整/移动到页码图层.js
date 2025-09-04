// Get the active document
var doc = app.activeDocument;

// Check if the "页码" layer exists, if not create a new layer named "Pages"
var targetLayer;
try {
    targetLayer = doc.layers.itemByName("Page Numbers");
    targetLayer.name; // This will throw an error if the layer does not exist
} catch (e) {
    targetLayer = doc.layers.add({ name: "Page Numbers" });
}

// Move the selected text frames to the target layer
var selectedItems = app.selection;
for (var i = 0; i < selectedItems.length; i++) {
    if (selectedItems[i] instanceof TextFrame) {
        selectedItems[i].itemLayer = targetLayer;
    }
}