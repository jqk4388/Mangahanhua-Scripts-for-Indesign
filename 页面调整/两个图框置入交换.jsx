//DESCRIPTION: SwapImages
// Based on Gerald Singelmann's Swap images startupscript
// http://indesign-faq.de/en/swapping_images_en
//

var mySelection = app.selection;
if (mySelection.length == 2 && mySelection[0].allGraphics.length > 0 && mySelection[1].allGraphics.length > 0)
{
if (mySelection[0].allGraphics[0].itemLink.status == LinkStatus.NORMAL && mySelection[1].allGraphics[0].itemLink.status == LinkStatus.NORMAL) {
var firstLink = mySelection[0].allGraphics[0].itemLink.filePath;
var secondLink = mySelection[1].allGraphics[0].itemLink.filePath;
mySelection[0].place(secondLink);
mySelection[1].place(firstLink);
} else {
alert("This cannot work with images whose status is not up-to-date");
}
} else {
alert("Please select two image frames");
}