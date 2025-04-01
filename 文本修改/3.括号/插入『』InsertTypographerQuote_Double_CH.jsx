//DESCRIPTION:Insert typographer's quote


#targetengine "setquotes"

/*

About Script:
The InDesign script "InsertTypographerQuote" adds defined typographic quotes before and after the selected text.

Setup:

Find the UTF-8 codes for the typographic quotes you want to use. 
Edit the script and change the values for the following variables:

SetQuotes.QUOTES_START
SetQuotes.QUOTES_END

Some common codes for german, english and swiss typographic quotes are included. 
If you are looking for other special charaters this link might help:


To Use: 
1. Open a document and select some text.
2. Double click the script "InsertTypographerQuote" from the scripts palette.
3. The typographic quotes before and after the selection are added.

Undo the action if needed by pressing Cmd + Z (Win: Strg + Z)

Contact InDesignScript.de:
http://www.indesignscript.de
Email: info@indesignscript.de
Twitter: @InDesignScript
    
*/




Array.prototype.exists = function (x) {
    for (var i = 0; i < this.length; i++) {
        if (this[i] == x) return true;
    }
    return false;
}



function SetQuotes() {
 
}

// Valid text types for selection
SetQuotes.SELECTION_ALLOWED = ["Text", "TextStyleRange", "Word", "Paragraph", "TextColumn","Line","Character"];

// Original typographes quotes setting for document
SetQuotes.TYPOGRAPHERS_QUOTES_VALUE = undefined;


// DIRECT quotation marks (double quote)

// German  
//SetQuotes.QUOTES_START = "\u201E"; // 99 unten
//SetQuotes.QUOTES_END = "\u201C"; // 66 oben

// Swiss 
//SetQuotes.QUOTES_START = "\u00BB"; //  >> 
//SetQuotes.QUOTES_END = "\u00AB"; // <<

//French
//SetQuotes.QUOTES_START = "\u00AB";  //  << 
//SetQuotes.QUOTES_END = "\u00BB";  // >>

// English
//SetQuotes.QUOTES_START = "\u201C"; //  "
//SetQuotes.QUOTES_END = "\u201D"; // "


// INDIRECT quotation marks (simple quote)

// German 
//SetQuotes.QUOTES_START = "\u201A"; // 9 unten
//SetQuotes.QUOTES_END = "\u2018"; // 6 oben

// Chinese 
SetQuotes.QUOTES_START = "\u300E"; //  『
SetQuotes.QUOTES_END = "\u300F"; // 』

//French
//SetQuotes.QUOTES_START = "\u2039";  //  < 
//SetQuotes.QUOTES_END = "\u203A";  // >

// English
//SetQuotes.QUOTES_START = "\u2018"; 
//SetQuotes.QUOTES_END = "\u2019"; 


// Straight
//SetQuotes.QUOTES_START = "\u0027"; 
//SetQuotes.QUOTES_END = "\u0027";

// +++++ Changes END +++++ //



SetQuotes.prototype.init = function() {
	
	var success = false;
	if (app.documents.length > 0) {

         if (app.selection.length > 0) {
            
            var selOK = SetQuotes.checkSelection();
                
                if (selOK) {
                    
                    var curSel = app.selection[0];
                    
                    SetQuotes.TYPOGRAPHERS_QUOTES_VALUE = app.activeDocument.textPreferences.typographersQuotes;
                    app.activeDocument.textPreferences.typographersQuotes = false;

                    success = true;
                   
                    
                 } else {                
                             alert("Selection not valid, Please select some text and try again" );
                 }            
              
            } else {
                alert("Text not selected, Please select some text and try again" );
            }

	} else {
		alert("Open at least one document to run this script." );
	}

	return success;
    
}


SetQuotes.prototype.reset= function() {
    
    app.activeDocument.textPreferences.typographersQuotes = SetQuotes.TYPOGRAPHERS_QUOTES_VALUE;
    
}
	


/*
    Only types of text may be selected by the user
    
  */

SetQuotes.checkSelection = function() {
   
    var selOK = true;
    
    var curSelection = app.selection;
    if (curSelection.length > 1) return false;
    
    var curItem = curSelection[0];
    var curItemType = curItem.constructor.name;
    var allowedItems = SetQuotes.SELECTION_ALLOWED;
    
    var isAllowed = allowedItems.exists(curItemType);
    if (!isAllowed) {
           selOK = false;
    }
    
    return selOK;
}


// Main
SetQuotes.prototype.setQuotes = function () {

    //var curDoc = app.activeDocument;
    var curSel = app.selection[0];

    var countInsertionPoints = curSel.insertionPoints.length;
    var firstInsertionPoint = curSel.insertionPoints.firstItem();
    var lastInsertionPoint = curSel.insertionPoints.lastItem();
    
    var lastSelectedCharacterIsLineBreak = SetQuotes.getLastCharacterIsLineBreak(curSel);

    // ignore the linebreak character
    if (lastSelectedCharacterIsLineBreak) {
       
        lastInsertionPoint = curSel.insertionPoints.item(countInsertionPoints-2);
        
    }
    

    if (firstInsertionPoint.isValid && lastInsertionPoint.isValid) {
    
        lastInsertionPoint.contents = SetQuotes.QUOTES_START;
        
        // InDesign selection is extended, if new content is added
        curSel = app.selection[0];
        
        if (lastSelectedCharacterIsLineBreak) {
            lastInsertionPoint = curSel.insertionPoints.item(countInsertionPoints-2);        
        } else {
            lastInsertionPoint = curSel.insertionPoints.lastItem();
        }
        
        lastInsertionPoint.contents = SetQuotes.QUOTES_END;
        
        var quoteStartCharacter = curSel.characters.lastItem();
        
        // first character of selection
        var firstCharacterOfSelection = curSel.characters.firstItem();
        
        // move character to the beginning of the selection
        quoteStartCharacter.move(LocationOptions.AT_BEGINNING, firstCharacterOfSelection);
       
        
    } else {
        alert("Error placing quotes please, try again");
    }
	
}


SetQuotes.getLastCharacterIsLineBreak = function (curSel) {
    
    lastCharacterIsLineBreak = false;
    
    if (curSel.isValid) {
        
        var lastCharacter = curSel.characters.lastItem();
        
        if (lastCharacter.contents == "\r") {
            lastCharacterIsLineBreak = true;
        }
        
    }
    
    return lastCharacterIsLineBreak;
}


function runSetQuotes() {
	var mySetQuotes = new SetQuotes();
	var success = mySetQuotes.init();
	if (success) {
        mySetQuotes.setQuotes();
        mySetQuotes.reset();
    }
    
}

// User can undo in one step
try {
    app.doScript("runSetQuotes()", ScriptLanguage.JAVASCRIPT , [], UndoModes.ENTIRE_SCRIPT, "InsertTypographerQuote.jsx");
} catch (e) {
     alert("Error in running script please, try again");
}

