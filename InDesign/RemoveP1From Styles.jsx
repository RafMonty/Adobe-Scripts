// InDesign script to:
// 1. Trim spaces from all style names (object, paragraph, character)
// 2. Remove "P#_" prefix from paragraph style names only

// Custom trim function since trim() isn't available in some InDesign versions
function trimString(str) {
    return str.replace(/^\s+|\s+$/g, '');
}

// Check if a style name exists in collection
function styleNameExists(collection, name) {
    try {
        var style = collection.itemByName(name);
        if (style.isValid) {
            return true;
        }
    } catch (e) {
        // Name doesn't exist
    }
    return false;
}

// Get a unique name by adding a suffix if needed
function getUniqueName(collection, baseName) {
    if (!styleNameExists(collection, baseName)) {
        return baseName;
    }
    
    var counter = 1;
    var newName = baseName + "_" + counter;
    
    while (styleNameExists(collection, newName)) {
        counter++;
        newName = baseName + "_" + counter;
    }
    
    return newName;
}

// Main function
function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var changes = {
        object: 0,
        paragraph: 0,
        character: 0
    };
    
    // Process all styles
    processObjectStyles(doc, changes);
    processParagraphStyles(doc, changes);
    processCharacterStyles(doc, changes);
    
    // Report results
    var totalChanges = changes.object + changes.paragraph + changes.character;
    alert("Processing complete:\n" + 
          "- Object styles: " + changes.object + " changed\n" +
          "- Paragraph styles: " + changes.paragraph + " changed\n" +
          "- Character styles: " + changes.character + " changed\n" +
          "Total changes: " + totalChanges);
}

// Process object styles (only trims spaces)
function processObjectStyles(doc, changes) {
    // Process root level styles
    for (var i = 0; i < doc.objectStyles.length; i++) {
        var style = doc.objectStyles[i];
        var oldName = style.name;
        var newName = trimString(oldName);
        
        if (oldName !== newName) {
            // Find a unique name if there's a conflict
            newName = getUniqueName(doc.objectStyles, newName);
            style.name = newName;
            changes.object++;
        }
    }
    
    // Process styles in groups
    for (var i = 0; i < doc.objectStyleGroups.length; i++) {
        processObjectStyleGroup(doc.objectStyleGroups[i], changes);
    }
}

// Process object style groups recursively
function processObjectStyleGroup(group, changes) {
    // Process styles in this group
    for (var i = 0; i < group.objectStyles.length; i++) {
        var style = group.objectStyles[i];
        var oldName = style.name;
        var newName = trimString(oldName);
        
        if (oldName !== newName) {
            // Find a unique name if there's a conflict
            newName = getUniqueName(group.objectStyles, newName);
            style.name = newName;
            changes.object++;
        }
    }
    
    // Process nested groups
    for (var i = 0; i < group.objectStyleGroups.length; i++) {
        processObjectStyleGroup(group.objectStyleGroups[i], changes);
    }
}

// Process paragraph styles (trims spaces and removes P#_ prefix)
function processParagraphStyles(doc, changes) {
    // Process root level styles
    for (var i = 0; i < doc.paragraphStyles.length; i++) {
        processParagraphStyle(doc.paragraphStyles[i], doc.paragraphStyles, changes);
    }
    
    // Process styles in groups
    for (var i = 0; i < doc.paragraphStyleGroups.length; i++) {
        processParagraphStyleGroup(doc.paragraphStyleGroups[i], changes);
    }
}

// Process a single paragraph style
function processParagraphStyle(style, collection, changes) {
    var oldName = style.name;
    var newName = trimString(oldName);
    
    // Remove P#_ prefix (P followed by any number, followed by underscore)
    if (/^P\d+_/.test(newName)) {
        newName = newName.substring(3);
    }
    
    if (oldName !== newName) {
        // Find a unique name if there's a conflict
        newName = getUniqueName(collection, newName);
        style.name = newName;
        changes.paragraph++;
    }
}

// Process paragraph style groups recursively
function processParagraphStyleGroup(group, changes) {
    // Process styles in this group
    for (var i = 0; i < group.paragraphStyles.length; i++) {
        processParagraphStyle(group.paragraphStyles[i], group.paragraphStyles, changes);
    }
    
    // Process nested groups
    for (var i = 0; i < group.paragraphStyleGroups.length; i++) {
        processParagraphStyleGroup(group.paragraphStyleGroups[i], changes);
    }
}

// Process character styles (only trims spaces)
function processCharacterStyles(doc, changes) {
    // Process root level styles
    for (var i = 0; i < doc.characterStyles.length; i++) {
        var style = doc.characterStyles[i];
        var oldName = style.name;
        var newName = trimString(oldName);
        
        if (oldName !== newName) {
            // Find a unique name if there's a conflict
            newName = getUniqueName(doc.characterStyles, newName);
            style.name = newName;
            changes.character++;
        }
    }
    
    // Process styles in groups
    for (var i = 0; i < doc.characterStyleGroups.length; i++) {
        processCharacterStyleGroup(doc.characterStyleGroups[i], changes);
    }
}

// Process character style groups recursively
function processCharacterStyleGroup(group, changes) {
    // Process styles in this group
    for (var i = 0; i < group.characterStyles.length; i++) {
        var style = group.characterStyles[i];
        var oldName = style.name;
        var newName = trimString(oldName);
        
        if (oldName !== newName) {
            // Find a unique name if there's a conflict
            newName = getUniqueName(group.characterStyles, newName);
            style.name = newName;
            changes.character++;
        }
    }
    
    // Process nested groups
    for (var i = 0; i < group.characterStyleGroups.length; i++) {
        processCharacterStyleGroup(group.characterStyleGroups[i], changes);
    }
}

// Run the script
main();