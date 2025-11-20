//RestoreSelection.jsx
//
//Restores a selection saved using the SaveSelection.jsx script.
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
	if (app.documents.length != 0){
		displayDialog();
	}
	else{
 	   alert ("请打开一个文档后重试。");
	}
}
function displayDialog(){
	var selectionNameField;
	var dialog = app.dialogs.add({name:"选择集还原"});
	var savedSelections = app.documents.item(0).extractLabel("保存选择集");
	if(savedSelections == ""){
		return;
	}
	savedSelections = savedSelections.split(",");
	with(dialog.dialogColumns.add()){
		with(dialogRows.add()){
			with(dialogColumns.add()){
				staticTexts.add({staticLabel:"选择区域已保存:"});
			}
			with(dialogColumns.add()){
				var dropdown = dropdowns.add({stringList:savedSelections, selectedIndex:0});
			}
		}
		with(dialogRows.add()){
			var zoomToSelectionCheckbox = checkboxControls.add({staticLabel:"缩放至选择区域", checkedState:false});
		}
	}
	var result = dialog.show();
	if(result == true){
		var selectionName = savedSelections[dropdown.selectedIndex];
		var zoomToSelection = zoomToSelectionCheckbox.checkedState;
		dialog.destroy();
		restoreSelection(selectionName, zoomToSelection);
	}
	else{
		dialog.destroy();
	}
}
function restoreSelection(selectionName, zoomToSelection){
	var pageItem;
	var idList = app.documents.item(0).extractLabel(selectionName);
	if(idList != ""){
		var objectList = new Array;
		idList = idList.split(",");
		for(var counter = 0; counter < idList.length; counter++){
			pageItem = app.documents.item(0).pageItems.itemByID(Number(idList[counter]));
			if(pageItem.isValid == true){
				objectList.push(pageItem);
			}
		}
		if(objectList.length > 0){
			try{
				app.documents.item(0).select(objectList);
				if(zoomToSelection == true){
					app.activeWindow.zoomPercentage = 200;
				}
			}
			catch(error){
				alert("无法完成选择。对象是否已移到其他页面或已被锁定？");
			}
		}
		else{
			alert("抱歉，保存的 ID 中没有与现有页面项匹配的任何对象。");
		}
	}
}
