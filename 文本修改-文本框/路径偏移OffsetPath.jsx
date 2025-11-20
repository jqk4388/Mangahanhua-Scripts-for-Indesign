//OffsetPaths.jsx
//An InDesign JavaScript by SPI
//
//Use InDesign's text wrap feature to create offset/inset paths.
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
	if(app.documents.length > 0){
		if(app.documents.item(0).selection.length > 0){
			var objectList = new Array();
			for(var counter = 0; counter < app.documents.item(0).selection.length; counter++){
				switch(app.documents.item(0).selection[counter].constructor.name){
					case "GraphicLine":
					case "Rectangle":
					case "Oval":
					case "Polygon":
						objectList.push(app.documents.item(0).selection[counter]);
					break;
				}
			}
			if(objectList.length > 0){
				displayDialog(objectList);
			}
		}
	}
}
function displayDialog(objectList){
	var dialog = app.dialogs.add({name:"路径偏移"});
	with(dialog.dialogColumns.add()){
		with(dialogRows.add()){
			with(dialogColumns.add()){
				staticTexts.add({staticLabel:"偏移距离:"});
			}
			with(dialogColumns.add()){
				var offsetField = measurementEditboxes.add({editValue:0});
			}
			with(dialogColumns.add()){
				staticTexts.add({staticLabel:"pts"});
			}
		}
	}
	var result = dialog.show();
	if(result){
		var offset = offsetField.editValue;
		dialog.destroy();
		offsetPaths(objectList, offset);
	}
	else{
		dialog.destroy();
	}
}
function offsetPaths(objectList, offset){
	//Store the current measurement units.
	var xUnits = app.documents.item(0).viewPreferences.horizontalMeasurementUnits;
	var yUnits = app.documents.item(0).viewPreferences.verticalMeasurementUnits;
	//Set measurement units to points.
	app.documents.item(0).viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
	app.documents.item(0).viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
	for(var counter = 0; counter < objectList.length; counter++){
		offsetPath(objectList[counter], offset);
	}
	//Set measurement units back to the original units.
	app.documents.item(0).viewPreferences.horizontalMeasurementUnits = xUnits;
	app.documents.item(0).viewPreferences.verticalMeasurementUnits = yUnits;
}
function offsetPath(pageItem, offset){
	pageItem.textWrapPreferences.textWrapMode = TextWrapModes.CONTOUR;
	var offsetArray = makeOffsetArray(pageItem, offset);
	pageItem.textWrapPreferences.textWrapOffset = offsetArray;
	var textWrapPaths = pageItem.textWrapPreferences.paths;
	var newPageItem = pageItem.parent.polygons.add();
	for(var counter = 0; counter < textWrapPaths.length; counter++){
		if(counter > 0){
			newPathItem.paths.add();
		}
		newPageItem.paths.item(counter).pathType = textWrapPaths.item(counter).pathType;
		newPageItem.paths.item(counter).entirePath = textWrapPaths.item(counter).entirePath;
	}
	pageItem.textWrapPreferences.textWrapMode = TextWrapModes.NONE;
}
function makeOffsetArray(pageItem, offset){
	var arrayLength;
	if(pageItem.textWrapPreferences.textWrapOffset.length != undefined){
		arrayLength = pageItem.textWrapPreferences.textWrapOffset.length;
		var array = new Array(arrayLength);
		for(var counter = 0; counter < arrayLength; counter++){
			array[counter] = offset;
		}
		return array;
	}
	else{
		return offset
	}
}