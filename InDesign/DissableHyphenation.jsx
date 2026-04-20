// InDesign ExtendScript to disable hyphenation in all paragraph styles, with warning and done dialogs

if (app.documents.length === 0) {
    alert("No document is open.");
} else {
    var doc = app.activeDocument;

    // Confirm with the user
    var confirmAction = confirm("This will turn OFF hyphenation in ALL paragraph styles.\nDo you want to continue?");
    
    if (confirmAction) {
        var styles = doc.paragraphStyles;
        var countChanged = 0;

        for (var i = 0; i < styles.length; i++) {
            try {
                if (styles[i].name !== "[No Paragraph Style]") {
                    styles[i].hyphenation = false;
                    countChanged++;
                }
            } catch (e) {
                $.writeln("Could not update style: " + styles[i].name + " - " + e);
            }
        }

        alert("Hyphenation turned off in " + countChanged + " paragraph style(s).");
    } else {
        alert("Operation cancelled.");
    }
}
