var isLtR = app.documents[0].documentPreferences.pageBinding == PageBindingOptions.LEFT_TO_RIGHT;
var doc = app.activeDocument;

function myDisplayDialog() {
    if (!app.activeWindow || !app.activeWindow.activePage) {
        alert('No active page available.');
        return;
    }

    if (!doc || !doc.pages || doc.pages.length < 1) {
        alert('No pages in the active document.');
        return;
    }

    var bookSize = doc.pages.length;
    var currentPageNum = app.activeWindow.activePage.documentOffset + 1;
    var bookPageNum = isLtR ? bookSize - currentPageNum + 1 : currentPageNum;
    var selectedPageNumber = null;

    try {
        var dialogWindow = new Window('dialog', 'Go to Page');
        dialogWindow.alignChildren = 'left';
        dialogWindow.add('statictext', undefined, 'Page number:');
        var input = dialogWindow.add('edittext', undefined, bookPageNum.toString());
        input.characters = 6;
        input.active = true;

        var buttonGroup = dialogWindow.add('group');
        buttonGroup.alignment = 'right';
        buttonGroup.add('button', undefined, 'OK', { name: 'ok' });
        buttonGroup.add('button', undefined, 'Cancel', { name: 'cancel' });

        if (dialogWindow.show() == 1) {
            selectedPageNumber = parseInt(input.text, 10);
        }
    } catch (e) {
        try {
            var myDialog = app.dialogs.add({ name: 'Go to Page' });
            with (myDialog) {
                with (dialogColumns.add()) {
                    selectedPageNumber = integerEditboxes.add({
                        editValue: bookPageNum,
                        minimumValue: 1,
                        maximumValue: bookSize
                    });
                }
            }
            var myReturn = myDialog.show();
            if (myReturn == true) {
                selectedPageNumber = selectedPageNumber.editValue;
            }
            myDialog.destroy();
        } catch (err) {
            alert('Unable to create dialog: ' + err);
            return;
        }
    }

    if (selectedPageNumber == null || isNaN(selectedPageNumber)) {
        return;
    }

    if (selectedPageNumber < 1 || selectedPageNumber > bookSize) {
        alert('Please enter a valid page number');
        return;
    }

    var newPageNum = isLtR ? bookSize - selectedPageNumber + 1 : selectedPageNumber;
    var newPage = doc.pages.item(newPageNum - 1);

    app.activeWindow.activePage = newPage.isValid ?
        newPage :
        (isLtR ? doc.pages.firstItem() : doc.pages.lastItem());
}

try {
    myDisplayDialog();
} catch (err) { alert(err) }