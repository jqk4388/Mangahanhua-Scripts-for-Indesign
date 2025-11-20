//ConnectCenters.jsx
//An InDesign JavaScript by SPI
//
//Connects the center points of the selected objects.
//
//by Olav Martin Kvern
//Distributed by Silicon Publishing, Inc.
//
//Our web site says, "Silicon Publishing, Inc. provides electronic publishing solutions and customizations 
//that automate the distribution of information between multiple sources and destinations." But what does
//that really mean? We make custom software for publishing. Some of it is for intense, data-driven, high-
//volume applications--think: directories, catalogs, and customized itineraries. Some of it is for web-to-print 
//or print-to-web applications--think: business cards, flyers, brochures, data sheets. Some of it is none of the 
//above. Use your imagination--if there's some chance that we can make your work easier, drop us a line! 
//sales@siliconpublishing.com
//
//Make sure that user interaction is turned on.
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
main();
function main(){
	var objectList = new Array;
	if(app.documents.length > 0){
		if(app.selection.length > 0){
			for(var counter = 0; counter < app.selection.length; counter++){
				switch(app.selection[counter].constructor.name){
					case "Rectangle":
					case "Oval":
					case "Polygon":
					case "GraphicLine":
						objectList.push(app.selection[counter]);
						break;
				}
			}
			if(objectList.length > 1){
				connectCenters(objectList);
			}
		}
	}
}
function connectCenters(objectList){
	var pointArray = new Array();
	for(var counter = 0; counter < objectList.length; counter++){
		pointArray.push(getCenterPoint(objectList[counter]));
	}
	var parent = objectList[0].parent;
	var polygon = parent.polygons.add();
	polygon.paths.item(0).entirePath = pointArray;
}
function getCenterPoint(pageItem){
	var x = pageItem.geometricBounds[1] + ((pageItem.geometricBounds[3] - pageItem.geometricBounds[1])/2);
	var y = pageItem.geometricBounds[0] + ((pageItem.geometricBounds[2] - pageItem.geometricBounds[0])/2);
	return new Array(x, y);
}