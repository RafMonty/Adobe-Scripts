// InDesign Script: Assign R&H Object Styles
// Applies object styles to page items based on their script labels
// Converts frames to correct type using InDesign's built-in content type

// Define the mapping: script label -> [content type, object style]
var styleMapping = {
    "AgentDetails": ["Text", "AlignBottom"],
    "AgentDetailsInConjunction": ["Text", "AlignBottom"],
    "bgGOLD": ["Graphic", "bgGOLD"],
    "bgGREY": ["Graphic", "bgGREY"],
    "Exclusive": ["Text", "Exclusive"],
    "PageBG": ["Graphic", "bgWhite"],
    "Pic99": ["Graphic", "FitToFrame"],
    "qrBG": ["Graphic", "bgWhite"],
    "RLA": ["Text", "AlignBottom"],
    "SaleMethod": ["Text", "AlignTop"],
    "CropMarks": ["Graphic", "FitToFrame"],
    "Logo": ["Graphic", "FitToFrame"],
    "Pic1": ["Graphic", "AutoClip"],
    "qrBox": ["Graphic", "FitToFrame"],
    "Tagline": ["Graphic", "FitToFrame"]
};

// Check if document is open
if (app.documents.length == 0) {
    alert("Please open a document first.");
    exit();
}

var doc = app.activeDocument;

// Function to recursively find an object style by name, even in folders
function findObjectStyle(styleName, parent) {
    if (!parent) {
        parent = doc;
    }
    
    // Check direct children
    for (var i = 0; i < parent.objectStyles.length; i++) {
        if (parent.objectStyles[i].name == styleName) {
            return parent.objectStyles[i];
        }
    }
    
    // Check in object style groups (folders)
    if (parent.objectStyleGroups) {
        for (var j = 0; j < parent.objectStyleGroups.length; j++) {
            var result = findObjectStyle(styleName, parent.objectStyleGroups[j]);
            if (result) {
                return result;
            }
        }
    }
    
    return null;
}

// Function to set content type (Graphic or Text)
function setContentType(item, targetType) {
    try {
        // Only set content type on page items that support it
        if (item.hasOwnProperty('contentType')) {
            var targetContentType;
            
            if (targetType == "Graphic") {
                targetContentType = ContentType.GRAPHIC_TYPE;
            } else if (targetType == "Text") {
                targetContentType = ContentType.TEXT_TYPE;
            } else {
                return false; // Unknown type
            }
            
            // Only convert if different
            if (item.contentType != targetContentType) {
                item.contentType = targetContentType;
                return true; // Conversion occurred
            }
        }
        return false; // No conversion needed
        
    } catch (e) {
        return false; // Conversion failed
    }
}

// Get all unique object style names needed
var requiredStyles = {};
for (var label in styleMapping) {
    requiredStyles[styleMapping[label][1]] = true;
}

// Check if all required object styles exist
var missingStyles = [];
var foundStyles = {};

for (var styleName in requiredStyles) {
    var style = findObjectStyle(styleName);
    if (style) {
        foundStyles[styleName] = style;
    } else {
        missingStyles.push(styleName);
    }
}

// Alert if styles are missing
if (missingStyles.length > 0) {
    alert("Missing Object Styles:\n\n" + missingStyles.join("\n") + 
          "\n\nPlease create these object styles before running this script.");
    exit();
}

// Process all pages in the document
var processedCount = 0;
var convertedCount = 0;
var notFoundLabels = [];

app.scriptPreferences.enableRedraw = false;

try {
    // Loop through each script label we're looking for
    for (var scriptLabel in styleMapping) {
        var contentType = styleMapping[scriptLabel][0];
        var objectStyleName = styleMapping[scriptLabel][1];
        var objectStyle = foundStyles[objectStyleName];
        var foundAny = false;
        
        // Search through all pages
        for (var i = 0; i < doc.pages.length; i++) {
            var page = doc.pages[i];
            var allItems = page.allPageItems;
            
            // Check each item on the page
            for (var j = 0; j < allItems.length; j++) {
                var item = allItems[j];
                
                // Check if this item has the script label we're looking for
                if (item.label == scriptLabel) {
                    // Convert to correct content type if needed
                    var wasConverted = setContentType(item, contentType);
                    if (wasConverted) {
                        convertedCount++;
                    }
                    
                    // Apply object style
                    item.appliedObjectStyle = objectStyle;
                    processedCount++;
                    foundAny = true;
                }
            }
        }
        
        // Track labels that weren't found
        if (!foundAny) {
            notFoundLabels.push(scriptLabel);
        }
    }
    
} finally {
    app.scriptPreferences.enableRedraw = true;
}

// Report results
var message = "Object Styles Applied Successfully!\n\n";
message += "Objects processed: " + processedCount;

if (convertedCount > 0) {
    message += "\nFrames converted: " + convertedCount;
}

if (notFoundLabels.length > 0) {
    message += "\n\nScript labels not found in document:\n" + notFoundLabels.join("\n");
}

alert(message);