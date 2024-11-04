//设置毫米单位。
app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;

// set the ruler to "spread", again so our math works out.
oldOrigin = app.activeDocument.viewPreferences.rulerOrigin // save old ruler origin
app.activeDocument.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN

// Set up right-to-left binding order
app.activeDocument.documentPreferences.pageBinding = PageBindingOptions.RIGHT_TO_LEFT

// Initialize first layer as Art layer
app.activeDocument.layers[0].name = 'Art'
var artLayer = app.activeDocument.layers[0]

// Create other standard layers
var designLayer = app.activeDocument.layers.add({name:'Design'}) 
var retouchingLayer = app.activeDocument.layers.add({name:'Clean'}) 
var pageNumberLayer = app.activeDocument.layers.add({name:'Page Numbers'}) 
var textLayer = app.activeDocument.layers.add({name:'Text'})
var SFXLayer = app.activeDocument.layers.add({name:'SFX'})

// This reuses the function from the "match art frame to page size plus bleed" script
// 给图框设置成文档设置的出血量
function matchFrameToPageSize(theFrame) {
    if (app.activeDocument.documentPreferences.pageBinding == PageBindingOptions.leftToRight) { // if the book's laid out left-to-right
        if (theFrame.parentPage.index % 2 == 0) { // if we’re on a left-side page
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - app.activeDocument.documentPreferences.documentBleedTopOffset, // Same, but for right-side pages
                theFrame.parentPage.bounds[1] - app.activeDocument.documentPreferences.documentBleedBottomOffset, 
                theFrame.parentPage.bounds[2] + app.activeDocument.documentPreferences.documentBleedOutsideOrRightOffset, 
                theFrame.parentPage.bounds[3] + 0];
        } else { // we must be on a right-side page
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - app.activeDocument.documentPreferences.documentBleedTopOffset, // Adjust the dimensions to give 1/8" bleed on right-side pages
                theFrame.parentPage.bounds[1] - 0, // 
                theFrame.parentPage.bounds[2] + app.activeDocument.documentPreferences.documentBleedOutsideOrRightOffset, 
                theFrame.parentPage.bounds[3] + app.activeDocument.documentPreferences.documentBleedInsideOrLeftOffset];
        }
       
        } 
    else { // if the book is laid out right-to-left
        if (theFrame.parentPage.index % 2 == 0) { // if we’re on a right-side page
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - app.activeDocument.documentPreferences.documentBleedTopOffset, // Adjust the dimensions to give 1/8" bleed
                theFrame.parentPage.bounds[1] + 0, // on right-side pages
                theFrame.parentPage.bounds[2] + app.activeDocument.documentPreferences.documentBleedOutsideOrRightOffset, 
                theFrame.parentPage.bounds[3] + app.activeDocument.documentPreferences.documentBleedInsideOrLeftOffset];
        } else { // we must be on a left-side page
            theFrame.geometricBounds = [
                theFrame.parentPage.bounds[0] - app.activeDocument.documentPreferences.documentBleedTopOffset, // Same, but for right-side pages
                theFrame.parentPage.bounds[1] - app.activeDocument.documentPreferences.documentBleedBottomOffset, 
                theFrame.parentPage.bounds[2] + app.activeDocument.documentPreferences.documentBleedOutsideOrRightOffset, 
                theFrame.parentPage.bounds[3] + 0];
        }
    }
}


// Set up master page object for convenience
theMaster = app.activeDocument.masterSpreads[0]
//查找第一页有没有图框，没有则去主页里新建一个图框
getFrame(app.activeDocument.pages[0], artLayer);

function getFrame(page, layer) {
    var targetFrame = null;

    var findFrame = function(arr, ignoreLayer) {
        var returnVal = null;
        // iterate backwards because it's most likely on the bottom
        for (var i = arr.length; i > -1; i-- && returnVal === null) {
            // only grab an object if it's a Rectangle on a non-locked layer
            if (arr[i] && arr[i] instanceof Rectangle &&
                !arr[i].itemLayer.locked &&
                (ignoreLayer || arr[i].itemLayer == layer)) { // if using an element from a master, don't be picky about the layer
                returnVal = arr[i];
            }
        }
        return returnVal;
    }

    // 检查是否有主页，覆盖主页并使用那里的框架
    if (page.appliedMaster !== null && page.appliedMaster.isValid) {
        var masterFrame = findFrame(page.masterPageItems, true); // 查找一个图框，但不挑剔它所在的图层
        targetFrame = masterFrame && masterFrame.isValid ?
            masterFrame.override(page) : null;
    };
        // 检查 Art 图层底部是否有矩形
    if (targetFrame === null) targetFrame = findFrame(page.allPageItems);
        // 如果没有图框，也没有主页图框，主页创建一个图框
    if (targetFrame === null) 
        newtwoboxs();
    return targetFrame;
}



/// THIS IS EXTREMELY FRAGILE AND EXPERIMENTAL, but: 
// this is also the beginning of a script that will place artwork automatically in a layout


var w = new Window ("palette"); // must be defined at top level
var myMessage = w.add ("statictext"); 
myMessage.text = "Placing artwork...";

// Regular expression to find single-page art with page numbers.
// Matches "P04.tif", "44.psd", etc.
var singlePgMatchExpr = /[pP]?\d+\.(tif|TIF|psd|PSD|tiff|TIFF)/ // Regular expression to find single-page art with page numbers

// Regular expression to find two-page art files with two page numbers across a spread.
// Matches "P04-05.tif", "44-45.psd", etc.
var doublePgMatchExpr = /[pP]?\d+[\-\_][pP]?\d+\.(tif|TIF|psd|PSD|tiff|TIFF)/ // Regular expression to find single-page art with page numbers

/* Function for recursively pushing matching art files onto artFileArr. It's not great to have a recursive function modifying an array outside of its own scope,
but I don't feel like figuring out how to make fileList more cleanly recursive right now. */

// create an external variable we'll use to keep track of the highest-numbered page
var lastPage = 0

function newtwoboxs() {
    for (var i = 0; i < theMaster.pages.length; i++) {
        thisRect = theMaster.pages[i].rectangles.add(
            app.activeDocument.layers.itemByName('Art'), // art layer
            LocationOptions.AT_BEGINNING, // inside the beginning of the list of objects 
            theMaster.pages[i], // with the current page as the reference object
            {
                name: 'ArtFrame',
                layer: app.activeDocument.layers.itemByName('Art'),
                geometricBounds: theMaster.pages[i].bounds,
                contentType: ContentType.GRAPHIC_TYPE, // make sure to specify that this is a graphic container
            });
        //加载了空框架之后设置适合选项
        thisRect.frameFittingOptions.fittingAlignment = AnchorPoint.CENTER_ANCHOR;
        thisRect.frameFittingOptions.fittingOnEmptyFrame = EmptyFrameFittingOptions.FILL_PROPORTIONALLY;
        //设置空框架的描边为无颜色
        thisRect.strokeWeight = 0;
        matchFrameToPageSize(thisRect); // 图框加大到出血线
    }
    return i;
}

function findArtFiles(fileList) { // add files whose names match singlePgMatchExpr to the artFileArr array
    var foundFiles = []
    for (var i=0; i < fileList.length; i++) {
        $.writeln('total files to eval: ' + fileList.length)
        $.writeln('evaluating: ' + fileList[i].name + ', iteration ' + i)
        if (fileList[i].constructor == Folder) { // if the current element has the Folder object as a constructor, e.g., it is a folder…
            $.writeln('found Folder: ' + fileList[i].name)
            foundFiles = concat(foundFiles, findArtFiles(fileList[i].getFiles())) // recurse and continue, concatenating as we go
        } else {
            if (fileList[i].displayName.match(singlePgMatchExpr)) { // if current element's filename matches our expression
                foundFiles.push(fileList[i]) // add file to the array we'll be returning
                // check if filename is bigger than the last biggest one
                if (parseInt(fileList[i].displayName.match(/\d+/)[0], 10) > lastPage) {
                    lastPage = parseInt(fileList[i].displayName.match(/\d+/)[0], 10) // if it is, parse the match as an int and assign it to lastPage
                }
            }
        }
    }
    return foundFiles
}

function addPagesUpToSignature(highestPage) {
    signaturePages = highestPage + (16 - (highestPage % 16))
    for (i=0; i < signaturePages - 1 ; i++) { // add one less than the number of pages in signaturePages because the document already has one page.
        app.activeDocument.pages.add(
            LocationOptions.AT_END,
            {
                appliedMaster: app.activeDocument.masterSpreads[0]
            }
        )
    }
}
 
function placeAllArtwork() {
    if (!w.pbar) { // if the progress bar doesn't exist
        w.pbar = w.add('progressbar', undefined, 0, artFileArr.length); 
    } else {
        w.pbar.value = 0;
        w.update();
    }
    w.pbar.preferredSize.width = 300;
    w.show(); // Show our progress bar window

    for (var i=0; i < artFileArr.length; i++) { // iterate over array and place every file.
            // this insane line converts the matched to an integer, then back to a string, 
            //in order to strip any leading "0" characters from the match. It does this because 
            //the itemByName method will only match correctly if the _string_ value matches the page
            //number.
        var currentPage = app.activeDocument.pages.itemByName(parseInt(artFileArr[i].displayName.match(/\d+/)[0], 10).toString())
        var artFrame = currentPage.masterPageItems[0].override(currentPage) // detatch master art frame instance on current page and make it editable
        // app.activeWindow.activePage = currentPage // jump view to page being placed
        artFrame.place(artFileArr[i])

        w.pbar.value = i + 1; // increment progress bar
        w.update() // Have to call this, or the progress bar won't update.
    }
    w.close()   
}

// intialize the artFileArr by prompting for a folder to start with, getting the contents, and passing them to our art-finding script
// var artFileArr = findArtFiles(Folder.selectDialog('选择无字图的文件夹。').getFiles()) 


// addPagesUpToSignature(lastPage)

// placeAllArtwork()

oldOrigin = app.activeDocument.viewPreferences.rulerOrigin // save old ruler origin

// Spread Origin is necessary for some reason, otherwise I can't figure out how to refer to dimensions on both pages
// app.activeDocument.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN 
app.activeDocument.viewPreferences.rulerOrigin = RulerOrigin.SPINE_ORIGIN

// Create a new layer for putting Guides on, if it doesn't already exist
if (app.activeDocument.layers.itemByName('Guides') == null ) {
    app.activeDocument.layers.add({name:'Guides'}) 
}


// Add top margin guide
theMaster.guides.add(
    {
        label: 'Top',
        orientation: HorizontalOrVertical.HORIZONTAL,
        location: theMaster.pages[0].bounds[0] + app.activeDocument.marginPreferences.top
    }
)

// Add bottom margin guide
theMaster.guides.add(
    {
        label: 'Bottom',
        orientation: HorizontalOrVertical.HORIZONTAL,
        location: theMaster.pages[0].bounds[2] - app.activeDocument.marginPreferences.bottom
    }
)

// Add recto inside guide
theMaster.guides.add(
    {
        label: 'RectoInside',
        orientation: HorizontalOrVertical.VERTICAL,
        location: app.activeDocument.marginPreferences.left
    }
)


// Add recto outside guide
theMaster.guides.add(
    {
        label: 'RectoOutside',
        orientation: HorizontalOrVertical.VERTICAL,
        location: theMaster.pages[0].bounds[1] - app.activeDocument.marginPreferences.right
    }
)

// Add verso inside guide
theMaster.guides.add(
    {
        label: 'VersoInside',
        orientation: HorizontalOrVertical.VERTICAL,
        location: - app.activeDocument.marginPreferences.left
    }
)

// Add verso outside guide
theMaster.guides.add(
    {
        label: 'VersoOutside',
        orientation: HorizontalOrVertical.VERTICAL,
        location: -theMaster.pages[1].bounds[3] + app.activeDocument.marginPreferences.right
    }
)

// PAGE NUMBERS SETUP INCLUDES
// - Create B-Master Page
// - Create Paragraph Style for Page Numbers
// - Place Page Number textFrames on both sides of B-Master

// Add B-Master and set it to inherit from A-Master
var theBMaster = app.activeDocument.masterSpreads.add(2, {
    namePrefix:"P",
    baseName:"页码",
    appliedMaster: theMaster, // set it to inherit from A-Master  
})

// Create page number paragraph style
pageNumberStyle = app.activeDocument.paragraphStyles.add({
    name: "Page Numbers",
    justification: Justification.AWAY_FROM_BINDING_SIDE,
})

// Calculate placement of verso page number box based on placement of named guides
var versoTopLeftX = theMaster.guides.itemByName('VersoOutside').location + 30
var versoTopLeftY = theMaster.guides.itemByName('Bottom').location
var versoBottomLeftX = theMaster.guides.itemByName('VersoOutside').location
var versoBottomLeftY = theMaster.guides.itemByName('Bottom').location + 15

// Calculate placement of recto page number box based on placement of named guides
var rectoTopLeftX = theMaster.guides.itemByName('RectoOutside').location
var rectoTopLeftY = theMaster.guides.itemByName('Bottom').location
var rectoBottomLeftX = theMaster.guides.itemByName('RectoOutside').location - 30
var rectoBottomLeftY = theMaster.guides.itemByName('Bottom').location + 15


// Add text boxes with page numbers to B-Master
var pageNumberL = theBMaster.textFrames.add(app.activeDocument.layers[0], LocationOptions.AT_END, theBMaster, {
    contents: SpecialCharacters.AUTO_PAGE_NUMBER,
    geometricBounds: [versoTopLeftY,versoTopLeftX,versoBottomLeftY,versoBottomLeftX],
})
pageNumberL.texts[0].appliedParagraphStyle = pageNumberStyle;

var pageNumberR = theBMaster.textFrames.add(app.activeDocument.layers[0], LocationOptions.AT_END, theBMaster, {
    contents: SpecialCharacters.AUTO_PAGE_NUMBER,
    geometricBounds: [rectoTopLeftY,rectoTopLeftX,rectoBottomLeftY,rectoBottomLeftX],
})
pageNumberR.texts[0].appliedParagraphStyle = pageNumberStyle;

// When we're done, change the scriptPreferences MeasurementUnit to its default, so
// we don't accidentally break any other scripts.
app.scriptPreferences.measurementUnit = AutoEnum.AUTO_VALUE

oldOrigin = app.activeDocument.viewPreferences.rulerOrigin // save old ruler origin


// When we're done, change the scriptPreferences MeasurementUnit to its default, so
// we don't accidentally break any other scripts.
app.scriptPreferences.measurementUnit = AutoEnum.AUTO_VALUE
// And restore old ruler origin
app.activeDocument.viewPreferences.rulerOrigin = oldOrigin 
app.activeDocument.pages[0].select()

function concat(arr1, arr2){
    for (var j = 0; j < arr2.length; j++)
        arr1.push(arr2[j]);
    return arr1;
}
