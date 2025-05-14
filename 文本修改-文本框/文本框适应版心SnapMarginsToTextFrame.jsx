//DESCRIPTION:Snap margins to text frame

/*
About Script: 
Modify the document margins easily and visually instead of fiddling about with numbers in the Margins and Columns dialog. 
This script will modify the document's margins throughout to match the selected text frame.

To Use: 
Draw or select a text frame on a right-hand (odd-numbered) page and run the script.
The document's margins will be modified to match the selected frame.
Any text frames that abut the margins will be adjusted and resized as needed.

More Info: 
For more information about this script and how to use it, please visit https://www.id-extras.com/snap-margins-to-text-frame/
 Copyright (c) by Ariel Walden, www.Id-Extras.com, 2020
Version 1.0,0


/////////////////////////////////////////////
//   In Memoriam
//   RAPHAEL MYER WALDEN
//   1935-2019
/////////////////////////////////////////////
*/

gScriptName = "Snap Margins to TextFrame";

// The following line (a) makes the script run faster; and (b) allows it to be undone with a single click.
app.doScript(premain, undefined, undefined,UndoModes.ENTIRE_SCRIPT, gScriptName);

// Let's make sure the script units are points.
// But once the script finishes, or if any errors occur, let's make sure to set the units back to what they were.

function premain(){
	// Check that at least one document is open.
	if (app.documents.length == 0){
		alert("Open at least one document to run this script.", gScriptName);
		return;
	}
	// Let's make sure the script units are points.
	// But once the script finishes, or if any errors occur, let's make sure to set the units back to what they were.
	var e, oldUnits = app.scriptPreferences.measurementUnit;
	app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
	try{
		main();
	}
	catch(e){
		alert("Something's gone wrong. The script must now quit. Margins have not been modified.", gScriptName);
	}
	finally{
		app.scriptPreferences.measurementUnit = oldUnits;
	}
}


function main(){
	var s = app.selection[0];
	var allPages;
	if (s instanceof TextFrame == false){
		alert("Select a text frame and then rerun the script.", gScriptName);
		return;
	}
	if (s.rotationAngle != 0){
		alert("Cannot set margins as the selected text frame is rotated. Select a non-rotated text frame and rerun the script.", gScriptName);
		return;
	}
	// If available, let's rely on InDesign's old Layout Adjustment feature to resize and reposition page items as we modify the margins...
	var myLayoutAdjustmentStatus;
	if (document.hasOwnProperty("layoutAdjustmentPreferences")){
		myLayoutAdjustmentStatus = document.layoutAdjustmentPreferences.enableLayoutAdjustment;
		document.layoutAdjustmentPreferences.enableLayoutAdjustment = true;
	}
	// But in version 2019 onwards, the old Layout Adjustment feature was replaced with the new Adjust Layout feature.
	// This does not  do quite the same thing, and because it relies on AI, the results are not sufficiently predictable for our use.
	// Therefore, disable the feature.
	if (document.hasOwnProperty("adjustLayoutPreferences")){
		myLayoutAdjustmentStatus = document.adjustLayoutPreferences.properties;
		document.adjustLayoutPreferences.enableAdjustLayout = false;
		document.adjustLayoutPreferences.allowFontSizeAndLeadingAdjustment = false;
		document.adjustLayoutPreferences.allowLockedObjectsToAdjust = false;
		document.adjustLayoutPreferences.enableAutoAdjustMargins = false;
	}
	var b = s.geometricBounds;
	var p = s.parentPage;
	// Check that the selected frame is on a right-hand page (for documents with facing pages).
	// Really, this is an artificial constraint. It just makes things a little bit simpler later on.
	if (p.side == PageSideOptions.LEFT_HAND){
		alert("To apply mirror margins, select a text frame on the right-hand page.", gScriptName);
		return;
	}
	// Remind the user what the script does, and give them a last chance to bail out!
	if (! confirm("Margins will be adjusted to match the selected text frame throughout the document, including master pages. Do you want to continue?", undefined, gScriptName)){
		return;
	}
	var z = p.bounds;
	// The top margin is the distance from the top of the selected frame to the top of the page.
	b[0] -= z[0];
	// The left margin, the distance from left of frame to left of page.
	b[1] -= z[1];
	// The bottom margin is the distance from the bottom of the page to the bottom of the frame.
	b[2] = z[2] - b[2];
	// And finally, the right margin is the distance from the right edge of the page to the right edge of the selected frame.
	b[3] = z[3] - b[3];
	
	if (document.hasOwnProperty("adjustLayoutPreferences")){
		// If we're having to use our own layout adjustment, then we've first got to do the regular pages, then the master pages
		allPages = app.activeDocument.pages.everyItem().getElements();
		allPages = allPages.concat(app.activeDocument.masterSpreads.everyItem().pages.everyItem().getElements());
		fit2019(allPages, b);
		document.adjustLayoutPreferences.properties = myLayoutAdjustmentStatus;			
	}
	else{
		// If we're using InDesign's old layout adjustment, it's more efficient to change the master pages first.
		allPages = app.activeDocument.masterSpreads.everyItem().pages.everyItem().getElements();
		allPages = allPages.concat(app.activeDocument.pages.everyItem().getElements());
		fit(allPages, b);
		document.layoutAdjustmentPreferences.enableLayoutAdjustment = myLayoutAdjustmentStatus;	
	}
    alert("All margins were successfully modified to match the current selection.", gScriptName);
}

// How simple life was with the old layout adjustment feature! Switch it on, modify the page margins, and let InDesign do the rest...
function fit(thePages, theBounds){
	var i;
	for (i = 0; i < thePages.length; i++){
		with (thePages[i].marginPreferences){
			top = theBounds[0];
			bottom = theBounds[2];
			left = theBounds[1];
			right = theBounds[3];
		}
	}
}

// But with CC2019, the old Layout Adjustment was removed. We need to reprogram our own...
function fit2019(thePages, theBounds){
	var i;
	for (i = 0; i < thePages.length; i++){
		adjustMargins(thePages[i], theBounds);
	}
}

// Let's try to emulate InDesign's old Layout Adjustment feature.
// The idea is to keep track of the positions of all objects on the page and check if they touch the margins on any side.
// The margins are then adjusted, and finally if any of the sides of the object touched a margin, its bounds are adjusted as well.
// An added complication is that pages can have columns. So we need to check whether an item abuts not just the page margins, but any of the columns as well.
function adjustMargins(aPage, newBounds){
	var i, j;
	var p = aPage.marginPreferences;
	var pLeft, pRight;
	var pageBounds = aPage.bounds;
	var oldMargins = [];
	var newMargins = [];
	var oldColumnPositions = [];
	var newColumnPositions = [];
	// Get a collection of all items on the page
	var theItems = aPage.pageItems.everyItem();
	// Quickly get the position of all items on the page in one go.
	var itemBounds = theItems.geometricBounds;
	// Now that we've collected the item bounds for all items, it's best for speed to convert the collection of page items to an array.
	// For more info on this topic: http://www.indiscripts.com/post/2010/06/on-everyitem-part-1
	theItems = theItems.getElements();	
	var theBounds;
	var oldBounds;
	var snap = {};
	// For single-sided and right-hand pages the left and right margins are correct. But for left-hand pages they are reversed.
	if (aPage.side == PageSideOptions.LEFT_HAND){
		pLeft = p.right;
		pRight = p.left;
	}
	else{
		pLeft = p.left;
		pRight = p.right;
	}
	oldMargins[0] = pageBounds[0] + p.top;
	oldMargins[1] = pageBounds[1] + pLeft;
	oldMargins[2] = pageBounds[2] - p.bottom;
	oldMargins[3] = pageBounds[3] - pRight;
	oldColumnPositions = p.columnsPositions; // Columns positions
	// Avoid crashes and unexpected behavior
	aPage.marginPreferences.top = newBounds[0];
	aPage.marginPreferences.bottom = newBounds[2];
	aPage.marginPreferences.left = newBounds[1];
	aPage.marginPreferences.right = newBounds[3];
	p = aPage.marginPreferences;
	// Single-sided and right-hand pages the left and right margins are correct. But for left-hand pages they are reversed.
	if (aPage.side == PageSideOptions.LEFT_HAND){
		pLeft = p.right;
		pRight = p.left;
	}
	else{
		pLeft = p.left;
		pRight = p.right;
	}	
	newMargins[0] = pageBounds[0] + p.top;
	newMargins[1] = pageBounds[1] + pLeft;
	newMargins[2] = pageBounds[2] - p.bottom;
	newMargins[3] = pageBounds[3] - pRight;
	newColumnPositions = p.columnsPositions;
	// Column positions are always measured from the visible left margin, whatever side page this is.
	for (i = 0; i < oldColumnPositions.length; i++) oldColumnPositions[i] += oldMargins[1];
	for (i = 0; i < newColumnPositions.length; i++) newColumnPositions[i] += newMargins[1];
	// If margins and columns are the same, do nothing and return
	if (String(oldMargins) == String(newMargins) && String(oldColumnPositions) == String(newColumnPositions)) return;
	// Adjust page items
	for (i = 0; i < theItems.length; i++){
		// Ignore locked items
		if (theItems[i].locked || theItems[i].itemLayer.locked) continue;
		// The snap object will keep track of which edge of the object has been moved.
		// It' properties will be top, bottom, left, and right, all booleans.
		snap = {};
		theBounds = itemBounds[i];
		oldBounds = [].concat(theBounds);
		// If the top edge of the page item is within 1pt of the top margin, adjust its top edge to be flush with the new margin,
		// and set snap.top to true so we know that we have adjusted the top edge.
		if (epsilon(theBounds[0], oldMargins[0])){
			theBounds[0] = newMargins[0];
			snap.top = true;
		}
		// Ditto bottom margin.
		if (epsilon(theBounds[2], oldMargins[2])){
			theBounds[2] = newMargins[2];
			snap.bottom = true;
		}
		// Next, loop through all the column positions. If the left edge of the item was within 1pt of any of the old column positions, adjust the new position of its left edge,
		// and set snap.left to true so we know that we have adjusted the left edge.
		// Ditto for the right edge.
		for (j = 0; j < newColumnPositions.length && j < oldColumnPositions.length; j++){
			if (!snap.left && epsilon(theBounds[1], oldColumnPositions[j])){
				theBounds[1] = newColumnPositions[j];
				snap.left = true;
			}
			if (!snap.right && epsilon(theBounds[3], oldColumnPositions[j])){
				theBounds[3] = newColumnPositions[j];
				snap.right = true;
			}
		}
		// The way the old Layout Adjustment feature worked was as follows:
		// If only one edge of an item abutted a margin, then the entire item is moved so that it remains sitting on that margn, but its size is not changed.
		// But if two opposite edges of an item abutted a margin, then the object was both resized and moved,
		// Currently, theBounds array represents a resized object in all cases.
		// The following lines check whether only one of a pair of edges has been modified. If so, theBounds array is adjusted so that the size of the item is not changed.
		if (snap.top && !snap.bottom){
			theBounds[2] = theBounds[0] + (oldBounds[2] - oldBounds[0]);
		}
		if (snap.bottom && !snap.top){
			theBounds[0] = theBounds[2] - (oldBounds[2] - oldBounds[0]);
		}
		if (snap.left && !snap.right){
			theBounds[3] = theBounds[1] + (oldBounds[3] - oldBounds[1]);
		}
		if (snap.right && !snap.left){
			theBounds[1] = theBounds[3] - (oldBounds[3] - oldBounds[1]);
		}
		theItems[i].geometricBounds = theBounds;
	}

	// If the difference between a and b is smaller than 1, return true, else return false.
	function epsilon(a, b){
		return (Math.abs(a - b) < 1);
	}
}

	