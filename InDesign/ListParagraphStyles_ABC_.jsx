/**
 * Script Name: List and Apply Paragraph Styles
 * Description:
 *    This script lists all custom paragraph styles (excluding [Basic Paragraph]) 
 *    in the selected text frame. Each style has "_ABC_" placed on a new line 
 *    with the corresponding paragraph style applied to it.
 *
 * Usage:
 *    1. Open an InDesign document.
 *    2. Select a single text frame where the style samples should be inserted.
 *    3. Run this script.
 *
 * Notes:
 *    - Only user-defined paragraph styles are included (excluding [Basic Paragraph]).
 *    - The script checks for proper selection and alerts the user if a valid text frame isn't selected.
 *    - Existing contents of the selected text frame will be cleared.
 *
 * Author: [Your Name or Organization]
 * Date: 2025-04-16
 */
// Main function
(function () {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }
    var doc = app.activeDocument;
    // Validate selection
    if (app.selection.length !== 1 || !(app.selection[0] instanceof TextFrame)) {
        alert("Please select a single text frame before running the script.");
        return;
    }
    var textFrame = app.selection[0];
    // Collect all paragraph styles excluding [Basic Paragraph]
    var styles = doc.paragraphStyles;
    var validStyles = [];
    for (var i = 0; i < styles.length; i++) {
        if (styles[i].name !== "[Basic Paragraph]") {
            validStyles.push(styles[i]);
        }
    }
    if (validStyles.length === 0) {
        alert("No custom paragraph styles found in the document.");
        return;
    }
    // Clear any existing text in the frame
    textFrame.contents = "";
    var insertionPoint = textFrame.insertionPoints[0];
    // Insert "_ABC_" for each style and apply the style
    for (var j = 0; j < validStyles.length; j++) {
        insertionPoint.contents = "_ABC_";
        // Apply style to the inserted line
        var styledText = insertionPoint.paragraphs[0];
        styledText.appliedParagraphStyle = validStyles[j];
        // Add line break if it's not the last one
        if (j < validStyles.length - 1) {
            insertionPoint = textFrame.insertionPoints[-1];
            insertionPoint.contents = "\r";
            insertionPoint = textFrame.insertionPoints[-1];
        }
    }
    alert("Paragraph styles have been listed and styled in the selected frame.");
})();