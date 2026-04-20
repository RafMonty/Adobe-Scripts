
/************************************************************************ 
 * Script Name: RafsMoveHideSHowSize.jsx                                *
 * Version: 1.0                                                          *
 * Date: 2025-01-08                                                      *
 * Description: This script processes selected text and picture frames   *
 * in Adobe InDesign documents, generating formatted commands based on   *
 * the frame's properties such as position, size, and label. Users can   *
 * select from various commands (Hide, Show, Set Position, etc.) and     *
 * copy the generated output for use in external applications.           *
 * Features include group handling, automatic labeling, and validation.  *
 ************************************************************************/
#target indesign

// Main function
function main() {
    if (app.documents.length === 0) {
        alert("No document open!");
        return;
    }
    if (app.selection.length === 0) {
        alert("Please select some text or picture frames.");
        return;
    }

    var selectedItems = app.selection;
    var output = [];
    var commands = [
        "Artwork.HidePO",
        "Artwork.ShowPO",
        "Artwork.SetPOSize",
        "Artwork.SetPOPos",
        "Artwork.SetPOPosSize"
    ];

    // Collect frame data, including group handling
    function processItem(item, index) {
        if (item.hasOwnProperty("geometricBounds") && item.hasOwnProperty("label")) {
            var bounds = item.geometricBounds; // [y1, x1, y2, x2]
            var label = item.label || "Frame" + (index + 1); // Use "FrameX" if no label is found
            var x = (bounds[1]).toFixed(2);
            var y = (bounds[0]).toFixed(2);
            var width = (bounds[3] - bounds[1]).toFixed(2);
            var height = (bounds[2] - bounds[0]).toFixed(2);

            output.push({
                label: label,
                x: x,
                y: y,
                width: width,
                height: height
            });
        } else if (item.constructor.name === "Group") {
            // Recursively process group items
            for (var i = 0; i < item.allPageItems.length; i++) {
                processItem(item.allPageItems[i], index + i);
            }
        }
    }

    // Process all selected items
    for (var i = 0; i < selectedItems.length; i++) {
        processItem(selectedItems[i], i);
    }

    if (output.length === 0) {
        alert("No valid frames selected.");
        return;
    }

    // Create the UI dialog
    var dialog = new Window("dialog", "Frame Information and Commands");
    dialog.alignChildren = "fill";

    // Dropdown to select the command
    var commandDropdown = dialog.add("dropdownlist", undefined, commands);
    commandDropdown.selection = 0;

    // Text field to display the generated output
    var textField = dialog.add("edittext", undefined, "", { multiline: true, scrolling: true });
    textField.minimumSize.width = 500;
    textField.minimumSize.height = 300;

    // Function to update the output based on the selected command
    function updateOutput() {
        var selectedCommand = commandDropdown.selection.text;
        var formattedOutput = "";

        for (var i = 0; i < output.length; i++) {
            var item = output[i];
            switch (selectedCommand) {
                case "Artwork.HidePO":
                case "Artwork.ShowPO":
                    formattedOutput += selectedCommand + '("' + item.label + '");\n';
                    break;
                case "Artwork.SetPOSize":
                    formattedOutput += selectedCommand + '("' + item.label + '", ' + item.width + ', ' + item.height + ');\n';
                    break;
                case "Artwork.SetPOPos":
                    formattedOutput += selectedCommand + '("' + item.label + '", ' + item.x + ', ' + item.y + ');\n';
                    break;
                case "Artwork.SetPOPosSize":
                    formattedOutput += selectedCommand + '("' + item.label + '", ' + item.x + ', ' + item.y + ', ' + item.width + ', ' + item.height + ');\n';
                    break;
            }
        }

        textField.text = formattedOutput;
    }

    // Update output when the dropdown changes
    commandDropdown.onChange = updateOutput;

    // Generate the initial output
    updateOutput();

    // Add an OK button
    dialog.add("button", undefined, "OK");

    // Show the dialog
    dialog.show();
}

// Run the main function
main();
