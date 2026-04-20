#target indesign

// Function to trim white spaces
function trimString(str) {
    return str.replace(/^\s+|\s+$/g, '');
}

// Display an alert to the user before starting
var proceed = confirm("This script will trim leading and trailing spaces from all Script Labels in the document. Do you want to proceed?");

if (proceed) {
    var doc = app.activeDocument;
    var pageItems = doc.allPageItems;
    var labelsChecked = 0;
    var labelsFixed = 0;

    // Iterate over all page items
    for (var i = 0; i < pageItems.length; i++) {
        var item = pageItems[i];
        if (item.label != '') {
            labelsChecked++;
            var trimmedLabel = trimString(item.label);
            if (trimmedLabel !== item.label) {
                item.label = trimmedLabel;
                labelsFixed++;
            }
        }
    }

    // Display the results
    alert("Script Labels checked: " + labelsChecked + "\nScript Labels fixed: " + labelsFixed);
} else {
    alert("Operation canceled by user.");
}
