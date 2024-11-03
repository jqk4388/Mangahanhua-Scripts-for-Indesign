/*

Links GREP Relink
Copyright 2023 William Campbell
All Rights Reserved
https://www.marspremedia.com/contact

Permission to use, copy, modify, and/or distribute this software
for any purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.

*/

(function () {

    var title = "Links GREP Relink";

    if (!/indesign/i.test(app.name)) {
        alert("Script for InDesign", title, false);
        return;
    }

    var count;
    var doc; // Document
    var doneMessage;
    var error; // Error
    var folderLinks; // Folder
    var progress;

    // Reusable UI variables.
    var g; // group
    var gc1; // group (column)
    var gc2; // group (column)
    var p; // panel
    var w; // window

    // Permanent UI variables.
    var btnCancel;
    var btnFolderLinks;
    var btnOk;
    var cbIgnoreCase;
    var inpChange;
    var inpFind;
    var txtFolderLinks;

    // SETUP

    // Script requires open document.
    if (!app.documents.length) {
        alert("Open a document", title, false);
        return;
    }
    doc = app.activeDocument;

    // CREATE PROGRESS WINDOW

    progress = new Window("palette", "Progress", undefined, {
        "closeButton": false
    });
    progress.t = progress.add("statictext");
    progress.t.preferredSize.width = 450;
    progress.b = progress.add("progressbar");
    progress.b.preferredSize.width = 450;
    progress.display = function (message) {
        message && (this.t.text = message);
        this.show();
        this.update();
    };
    progress.increment = function () {
        this.b.value++;
    };
    progress.set = function (steps) {
        this.b.value = 0;
        this.b.minvalue = 0;
        this.b.maxvalue = steps;
    };

    // CREATE USER INTERFACE

    w = new Window("dialog", title);
    w.alignChildren = "fill";

    // Panel 'Link file name'.
    p = w.add("panel", undefined, "Link file name");
    p.alignChildren = "left";
    p.margins = [24, 24, 24, 18];
    p.spacing = 24;
    // Group of 2 columns.
    g = p.add("group");
    // Groups, columns 1 and 2.
    gc1 = g.add("group");
    gc1.orientation = "column";
    gc1.alignChildren = "left";
    gc2 = g.add("group");
    gc2.orientation = "column";
    gc2.alignChildren = "left";
    // Rows.
    gc1.add("statictext", undefined, "Find what:").preferredSize.height = 23;
    inpFind = gc2.add("edittext");
    inpFind.preferredSize = [360, 23];
    gc1.add("statictext", undefined, "Change to:").preferredSize.height = 23;
    inpChange = gc2.add("edittext");
    inpChange.preferredSize = [360, 23];
    // Checkbox ignore case.
    cbIgnoreCase = p.add("checkbox", undefined, "Ignore letter case");

    // Panel 'Search for links in folder'.
    p = w.add("panel", undefined, "Search for links in folder");
    p.alignChildren = "left";
    p.margins = [24, 24, 24, 18];
    g = p.add("group");
    btnFolderLinks = g.add("button", undefined, "Folder...");
    txtFolderLinks = g.add("statictext", undefined, "", {
        truncate: "middle"
    });
    txtFolderLinks.preferredSize.width = 350;

    // Action Buttons
    g = w.add("group");
    g.alignment = "center";
    btnOk = g.add("button", undefined, "OK");
    btnCancel = g.add("button", undefined, "Cancel");

    // Panel Copyright
    p = w.add("panel");
    p.add("statictext", undefined, "Copyright 2023 William Campbell");

    // SET UI VALUES

    // Links folder set to location of first document link (if any).
    if (doc.links.length) {
        txtFolderLinks.text = Folder.decode(new File(doc.links[0].filePath).path);
        folderLinks = new Folder(txtFolderLinks.text);
    }

    // UI ELEMENT EVENT HANDLERS

    btnFolderLinks.onClick = function () {
        var f = Folder.selectDialog("Select Links folder", txtFolderLinks.text);
        if (f) {
            txtFolderLinks.text = Folder.decode(f.fullName);
        }
    };

    btnOk.onClick = function () {
        if (!inpFind.text) {
            alert("Enter a value for 'Find what'");
            return;
        }
        folderLinks = new Folder(txtFolderLinks.text);
        if (!(folderLinks && folderLinks.exists)) {
            alert("Select Links folder", " ", false);
            return;
        }
        w.close(1);
    };

    btnCancel.onClick = function () {
        w.close(0);
    };

    // DISPLAY THE DIALOG

    if (w.show() == 1) {
        doneMessage = "";
        try {
            app.doScript(process, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, title);
            doneMessage = count + " links relinked";
        } catch (e) {
            error = error || e;
            doneMessage = "An error has occurred.\nLine " + error.line + ": " + error.message;
        }
        progress.close();
        doneMessage && alert(doneMessage, title, error);
    }

    //====================================================================
    //               END PROGRAM EXECUTION, BEGIN FUNCTIONS
    //====================================================================

    function process() {
        var file;
        var i;
        var link;
        var nameNew;
        var pattern;
        progress.display("Initializing...");
        count = 0;
        try {
            pattern = new RegExp(inpFind.text, "g" + (cbIgnoreCase.value ? "i" : ""));
            progress.set(doc.links.length);
            for (i = 0; i < doc.links.length; i++) {
                progress.increment();
                link = doc.links[i];
                if (inpFind.text) {
                    nameNew = link.name.replace(pattern, inpChange.text);
                    progress.display(link.name + " -> " + nameNew);
                    file = new File(folderLinks.fullName + "/" + nameNew);
                } else {
                    // 'Find what' is empty.
                    // Relink same name in specified folder.
                    progress.display(link.name);
                    file = new File(folderLinks.fullName + "/" + link.name);
                }
                if (file.exists) {
                    link.relink(file);
                    link.update();
                    count++;
                }
            }
        } catch (e) {
            error = e;
            throw e;
        }
    }

})();