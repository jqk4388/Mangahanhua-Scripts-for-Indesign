//DESCRIPTION:Insert unicode character

/*
About Script :
This script will insert a sequence of one or more Unicode characters at the current insertion point in InDesign
This script is a bit odd in that it is configured by renaming a  copy of the UnicodeInjector.jsx file.
 
The idea is that you make one or more copies of this script,
rename them according to the desired goal, and place these copies into your scripts folder.

For example, if you rename a copy of this script to:
   U+0061 U+0062 Insert Unicode chars.jsx

Copy it into your InDesign Scripts folder, each time you set your text cursor in some text, then double-click the script 
 (or hit a keyboard shortcut if you assign one), it will insert the
 unicode characters U+0061 and U+0062 (that is, 'A' and 'B' ) 

To Use: 
Make a copy of the UnicodeInjector.jsx file.
Look up the Unicode characters you desire to insert (e.g. assume we want to insert a hair space, U+200A).
Rename the copy so it contains the desired unicode(s) as well as some descriptive text - e.g.  U+200A Insert a hair space.jsx
Move the 'U+200A Insert a hair space.jsx' script to your InDesign Scripts folder. 
Go to InDesign - the script is ready for use. Position your text cursor
by double-clicking in any text frame, then double-click the script name on the Scripts Palette; the unicode character(s) will be inserted.

You can assign a keyboard shortcut to the script from Keyboard Shortcuts... menu. 

More info
---------
This script will ignore any descriptive text in the file name.
So in the above example, the text 'Insert Unicode chars' is 
ignored - only the U+0061 and U+0062 is interpreted.

The script understands the following notations:
 U+nnnn (U, +, then 1 to 4 hexadecimal digits)
   Examples: U+0041, u+20eF, u+61
 0xnnnn (0, x, then 1 to 4 hexadecimal digits)
   Examples: 0x0041, 0x20Ef, 0X61
 0dnnnnn (0, d, followed by 1 to 5 decimal digits)
   Examples: 0d65, 0d8431, 0D97 


// ----------------
About Rorohiko:
Rorohiko specialises in making printing, publishing and web workflows more efficient.
This script is a free sample of the custom solutions we create for our customers.
If your workflow is hampered by boring or repetitive tasks, inquire at sales@rorohiko.com
The scripts we write for our customers repay for themselves within weeks or months.

(C) 2011-2019 Rorohiko Ltd.
All rights reserved.
By Kris Coppieters
kris@rorohiko.com
*/

// This is the 'dummy' script name used when you run the script from the
// ExtendScript Toolkit
const kDebugSampleScriptName = "U+0061 U+0062 Insert Unicode chars.jsx";

const kUnicodeRegEx = /^(.*?u\+([0-9a-f]{1,4}))(.*)$/i;
const kHexaRegEx = /^(.*?0x([0-9a-f]{1,4}))(.*)$/i;
const kDeciRegEx = /^(.*?0d([1-9][0-9]*))(.*)$/i;

const kErr_NoError = 0;
const kErr_NoSelection = 1;
const kErr_MissingCodes = 2;
const kErr_NoDocument = 3;

var error = kErr_NoError;


do
{
     if (app.documents.length == 0) 
    {
        error = kErr_NoDocument;        
        break;
    }
    
    if (app.selection == null)
    {
        error = kErr_NoSelection;
        break;
    }

    if (! (app.selection instanceof Array))
    {
        error = kErr_NoSelection;
        break;
    }

    if (app.selection.length != 1)
    {
        error = kErr_NoSelection;
        break;
    }

    if (!(
        app.selection[0] instanceof InsertionPoint ||
        app.selection[0] instanceof Text ||
        app.selection[0] instanceof Word ||
        app.selection[0] instanceof Character ||
        app.selection[0] instanceof Paragraph ||
        app.selection[0] instanceof TextStyleRange
    ))
    {
        error = kErr_NoSelection;
        break;
    }

    var fileName; 
    var err;
    try
    {
        if (app.activeScript instanceof File)
        {
            fileName = app.activeScript.name;
        }
        else 
        {
            fileName = kDebugSampleScriptName;
        }
    }
    catch (err)
    {
        fileName = kDebugSampleScriptName;
    }
    
    var fileNameWasPrefixedWithCode;
    var charCode;
    var numberBase;
    var numericalCharCode;
    var matchingRegEx;
    
    error = kErr_MissingCodes;

    // Start a loop that strips code prefixes from the filename 
    // one by one, until all codes have been processed
    do
    {
        fileNameWasPrefixedWithCode = false;
    
        if (fileName.match(kUnicodeRegEx) != null)
        {
            matchingRegEx = kUnicodeRegEx;
            numberBase = 16;
        }
        else if (fileName.match(kHexaRegEx) != null)
        {
            matchingRegEx = kHexaRegEx;
            numberBase = 16;
        }
        else if (fileName.match(kDeciRegEx) != null)
        {
            matchingRegEx = kDeciRegEx;
            numberBase = 10;
        }
        else
        {
            matchingRegEx = null;
        }

        if (matchingRegEx != null)
        {
            charCode = fileName.replace(matchingRegEx,"$2");

            if (charCode != "")
            {
                numericalCharCode = parseInt(charCode, numberBase);
                if (! isNaN(numericalCharCode) && numericalCharCode != 0)
                {
                    error = kErr_NoError;
                    app.selection[0].contents = String.fromCharCode(numericalCharCode);

                    fileNameWasPrefixedWithCode = true;
                    // Strip off the code we just processed, and then loop back to try 
                    // to get the next code, if any
                    fileName = fileName.replace(matchingRegEx,"$3");
                }
            }
        }
    }
    while (fileNameWasPrefixedWithCode);
}
while (false);

switch (error)
{
case kErr_NoSelection:
    alert("Place the text cursor in the text frame, without selecting the text.");
    break;
case kErr_MissingCodes:
    alert("Rename the script to configure it. To view configuration instructions, right-click the script and choose 'Edit Script'.");
    break;
case kErr_NoDocument:
    alert("Open at least one document to run this script.");
    break;

}