/*

Links Rename Add Page Number
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

    var title = "Links Rename Add Page Number";

    if (!/indesign/i.test(app.name)) {
        alert("Script for InDesign", title, false);
        return;
    }

    // Script variables.
    var count;
    var doc;
    var doneMessage;
    var error;
    var progress;

    // Reusable UI variables.
    var g; // group
    var p; // panel
    var w; // window

    // Permanent UI variables.
    var btnCancel;
    var btnOk;
    var grpExtensions;
    var inpExtensions;
    var rbLinksAll;
    var rbLinksExtensions;
    var rbPrefix;
    var rbSuffix;

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
    p = w.add("panel");
    p.alignChildren = "left";
    p.margins = [24, 18, 24, 24];
    g = p.add("group");
    g.add("statictext", undefined, "Add:");
    g = g.add("group");
    g.margins = [0, 6, 0, 0];
    rbPrefix = g.add("radiobutton", undefined, "Prefix");
    rbSuffix = g.add("radiobutton", undefined, "Suffix");
    g = p.add("group");
    g.add("statictext", undefined, "Links:");
    g = g.add("group");
    g.margins = [0, 6, 0, 0];
    rbLinksAll = g.add("radiobutton", undefined, "All");
    rbLinksExtensions = g.add("radiobutton", undefined, "Extensions");
    grpExtensions = p.add("group");
    grpExtensions.add("statictext", undefined, "Extensions:");
    inpExtensions = grpExtensions.add("edittext");
    inpExtensions.preferredSize.width = 140;

    // Action Buttons
    g = w.add("group");
    g.alignment = "center";
    btnOk = g.add("button", undefined, "OK");
    btnCancel = g.add("button", undefined, "Cancel");

    // Panel Copyright
    p = w.add("panel");
    p.add("statictext", undefined, "Copyright 2023 William Campbell");

    // DEFAULTS

    // Add Prefix/Suffix
    rbPrefix.value = true;
    rbSuffix.value = false;
    // Links All/Extensions
    rbLinksAll.value = false;
    rbLinksExtensions.value = true;
    // Extensions
    inpExtensions.text = "pdf";

    configureUi();

    // UI ELEMENT EVENT HANDLERS

    rbLinksAll.onClick = configureUi;
    rbLinksExtensions.onClick = configureUi;

    btnOk.onClick = function () {
        if (!confirm("WARNING!\nFiles on disk will be renamed and this cannot be undone. Make backup copies of files before proceeding. Are you sure you want to continue?", true, title)) {
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
            process();
            doneMessage = doneMessage || count + " links renamed";
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

    function configureUi() {
        grpExtensions.enabled = rbLinksExtensions.value;
    }

    function process() {
        var extension;
        var extensions;
        var file;
        var fileNew;
        var fileVersion;
        var i;
        var ii;
        var link;
        var links;
        var baseName;
        var nameChanged;
        var nameNew;
        var page;
        var parent;
        var pattern;
        var padZero = function (v) {
            if (v < 10) {
                return ("0" + v).slice(-2);
            }
            return v;
        };
        count = 0;
        progress.display("Initializing...");
        links = [];
        if (rbLinksExtensions.value) {
            extensions = inpExtensions.text.replace(/\./g, "");
            extensions = extensions.replace(/[^a-zA-Z]+/g, " ");
            extensions = extensions.replace(/ {2,}/g, " ");
            extensions = extensions.split(" ");
            extensions = extensions.join("|");
            pattern = new RegExp("\\." + extensions + "$", "i");
        }
        for (i = 0; i < doc.links.length; i++) {
            link = doc.links[i];
            if (rbLinksAll.value || pattern.test(link.filePath)) {
                links.push(link);
            }
        }
        if (!links.length) {
            doneMessage = "No links found";
            return;
        }
        progress.set(links.length);
        for (i = 0; i < links.length; i++) {
            link = links[i];
            progress.increment();
            progress.display(link.name);
            file = new File(link.filePath);
            parent = link.parent;
            try {
                page = parent.parent.parentPage;
            } catch (_) {
                // Doesn't have parentPage.
                continue;
            }
            // Split filename into base name and extension.
            baseName = link.name.replace(/\.[^\.]*$/, "");
            extension = String(String(link.name.match(/\..*$/) || "").match(/[^\.]*$/) || "");
            if (rbPrefix.value) {
                nameChanged = padZero(page.name) + "-" + baseName;
            } else if (rbSuffix.value) {
                nameChanged = baseName + "-" + padZero(page.name);
            }
            // Did name change?
            if (nameChanged != baseName) {
                // Test if file with name exists.
                nameNew = nameChanged;
                fileVersion = 0;
                fileNew = new File(file.path + "/" + nameNew + "." + extension);
                while (fileNew.exists) {
                    // File exists. Add version suffix.
                    fileVersion++;
                    nameNew = nameChanged + "~" + fileVersion;
                    fileNew = new File(file.path + "/" + nameNew + "." + extension);
                }
                // Rename and relink.
                if (file.exists) {
                    progress.display(link.name + " -> " + nameNew + "." + extension);
                    file.rename(nameNew + "." + extension);
                    // Loop through all graphics and relink.
                    // Graphic could be placed more than once.
                    for (ii = 0; ii < doc.links.length; ii++) {
                        if (doc.links[ii].name == link.name) {
                            doc.links[ii].relink(file);
                            doc.links[ii].update();
                            count++;
                        }
                    }
                }
            }
        }
    }

})();