//SaveSelection.jsx
//An InDesign JavaScript by SPI
//
//Saves a selection as a series of page item IDs to a named script label on the document.
//Maintains a list of saved selections in the document. Does not save text selections.
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
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	var objectList = new Array;
	if (app.documents.length != 0){
		if (app.selection.length != 0){
			for(var counter = 0;counter < app.selection.length; counter++){
				switch (app.selection[counter].constructor.name){
					case "Rectangle":
					case "Oval":
					case "Polygon":
					case "TextFrame":
					case "Group":
					case "Button":
					case "GraphicLine":
						objectList.push(app.selection[counter]);
						break;
				}
			}
			if (objectList.length != 0){
				displayDialog(objectList);
			}
			else{
				alert ("请选择一个页面项后重试。");
			}
		}
		else{
			alert ("请选择一个对象后重试。");
		}
	}
	else{
 	   alert ("请打开一个文档、选择一个对象后重试。");
	}
}
function displayDialog(objectList){
	var selectionNameField;
	var fieldWidth = 200;
	var labelWidth = 120;
	var dialog = app.dialogs.add({name:"保存选择集"});
	var savedSelections = app.documents.item(0).extractLabel("保存选择集");
	with(dialog.dialogColumns.add()){
		if(savedSelections != ""){
			with(dialogRows.add()){
				savedSelections = savedSelections.split(",");
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:"已保存选择集:", minWidth:labelWidth});
				}
				with(dialogColumns.add()){
					var dropdown = dropdowns.add({stringList:savedSelections, selectedIndex:0, minWidth:fieldWidth});
				}
			}
			with(dialogRows.add()){
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:"新选择集名称:", minWidth:labelWidth});
				}
				with(dialogColumns.add()){
					selectionNameField = textEditboxes.add({editContents:"", minWidth:fieldWidth});
				}
			}
		}
		else{
			with(dialogRows.add()){
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:"新选择集名称:", minWidth:labelWidth});
				}
				with(dialogColumns.add()){
					selectionNameField = textEditboxes.add({editContents:"", minWidth:fieldWidth});
				}
			}
		}
	}
	var result = dialog.show();
	if(result == true){
		var selectionName;
		switch(true){
			//Saved selections exist, and user did not enter a new selection name.
			case ((savedSelections != "")&&(selectionNameField.editContents == "")):
				selectionName = savedSelections[dropdown.selectedIndex];
				break;
			//Saved selections exist, and user entered a new selection name.
			case ((savedSelections != "")&&(selectionNameField.editContents != "")):
				selectionName = selectionNameField.editContents;
				savedSelections.push(selectionName);
				app.documents.item(0).insertLabel("保存选择集", savedSelections.toString());
				break;
			//Saved selections do not exist, and user entered a new selection name.
			case ((savedSelections == "")&&(selectionNameField.editContents !="")):
				selectionName = selectionNameField.editContents;
				savedSelections = new Array(selectionName);
				app.documents.item(0).insertLabel("保存选择集", savedSelections.toString());
				break;
			//Saved selections did not exist, and user clicked OK without entering a new selection name.
			case ((savedSelections == "")&&(selectionNameField.editContents == "")):
				dialog.destroy();
				return;
				break;
		}
		dialog.destroy();
		saveSelection(objectList, selectionName);
	}
	else{
		dialog.destroy();
	}
}
function saveSelection(objectList, selectionName){
	var objectIds = new Array;
	for(var counter = 0; counter < objectList.length; counter++){
		objectIds.push(objectList[counter].id);
	}
	objectIds = objectIds.toString();
	app.documents.item(0).insertLabel(selectionName, objectIds);
}