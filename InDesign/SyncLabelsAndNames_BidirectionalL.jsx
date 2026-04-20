/**
 * Sync Script Labels and Object Names - Bidirectional
 * 
 * This script allows you to:
 * 1. Rename objects in Layers panel FROM their Script Label values
 * 2. Set Script Labels FROM object names in Layers panel
 *
 * Features:
 * - User dialog to choose direction
 * - Advanced text cleaning and trimming
 * - Process selection or entire page
 * - Detailed reporting
 *
 * @author Adobe InDesign Script
 * @version 2.0
 * @compatibility InDesign CC 2024+
 */

(function() {
    
    // ==================== CONFIGURATION ====================
    var CONFIG = {
        defaultBlankName: "_BLANK",
        showDetailedReport: true,
        processTextFrames: true,
        processGraphics: true,
        processGroups: true,
        cleanText: true,              // Apply text cleaning
        trimWhitespace: true,         // Remove leading/trailing spaces
        removeInvalidChars: true,     // Remove problematic characters
        maxLabelLength: 100,          // Maximum Script Label length
        removeAllSpaces: false        // Remove ALL spaces (set by dialog)
    };
    
    // ==================== MAIN FUNCTION ====================
    function main() {
        // Check if InDesign is running and a document is open
        if (app.documents.length === 0) {
            alert("Please open a document before running this script.");
            return;
        }
        
        var doc = app.activeDocument;
        var selection = app.selection;
        
        // Show direction selection dialog
        var dialogResult = showDirectionDialog();
        
        if (dialogResult.direction === null) {
            return; // User cancelled
        }
        
        // Store the removeAllSpaces option in CONFIG for use by cleaning function
        CONFIG.removeAllSpaces = dialogResult.removeAllSpaces;
        
        // Check if anything is selected
        if (selection.length === 0) {
            var response = confirm(
                "No objects are currently selected.\n\n" +
                "Would you like to process all objects on the current page instead?\n\n" +
                "Click 'OK' to process current page, or 'Cancel' to exit."
            );
            
            if (response) {
                selection = getAllPageItems(doc.activeWindow.activePage);
            } else {
                return;
            }
        }
        
        // Process the objects based on direction
        var results;
        if (dialogResult.direction === "toName") {
            results = processLabelToName(selection);
        } else {
            results = processNameToLabel(selection);
        }
        
        // Display report
        showReport(results, dialogResult.direction);
    }
    
    // ==================== SHOW DIRECTION DIALOG ====================
    function showDirectionDialog() {
        var dialog = app.dialogs.add({name: "Sync Script Labels and Object Names"});
        
        with(dialog.dialogColumns.add()) {
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: "Choose synchronization direction:"});
            }
            
            with(borderPanels.add()) {
                with(dialogColumns.add()) {
                    var radioGroup = radiobuttonGroups.add();
                    with(radioGroup) {
                        var radio1 = radiobuttonControls.add({
                            staticLabel: "Script Label --> Object Name",
                            checkedState: true
                        });
                        var radio2 = radiobuttonControls.add({
                            staticLabel: "Object Name --> Script Label"
                        });
                    }
                }
            }
            
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: ""});
            }
            
            // Trim/Remove Spaces Option
            with(borderPanels.add()) {
                with(dialogColumns.add()) {
                    with(dialogRows.add()) {
                        staticTexts.add({staticLabel: "Text Processing Options:"});
                    }
                    with(dialogRows.add()) {
                        var trimSpacesCheckbox = checkboxControls.add({
                            staticLabel: "Remove ALL spaces and trim (no spaces in result)",
                            checkedState: false
                        });
                    }
                }
            }
            
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: ""});
            }
            
            with(dialogRows.add()) {
                staticTexts.add({
                    staticLabel: "Script Label --> Object Name:",
                    minWidth: 350
                });
            }
            with(dialogRows.add()) {
                staticTexts.add({
                    staticLabel: "  - Renames objects in Layers panel using Script Label values",
                    minWidth: 350
                });
            }
            with(dialogRows.add()) {
                staticTexts.add({
                    staticLabel: "  - Objects without labels are named '_BLANK'",
                    minWidth: 350
                });
            }
            
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: ""});
            }
            
            with(dialogRows.add()) {
                staticTexts.add({
                    staticLabel: "Object Name --> Script Label:",
                    minWidth: 350
                });
            }
            with(dialogRows.add()) {
                staticTexts.add({
                    staticLabel: "  - Copies object names to Script Label field",
                    minWidth: 350
                });
            }
            with(dialogRows.add()) {
                staticTexts.add({
                    staticLabel: "  - Cleans and validates text automatically",
                    minWidth: 350
                });
            }
        }
        
        var result = dialog.show();
        
        var dialogResult = {
            direction: null,
            removeAllSpaces: false
        };
        
        if (result === true) {
            dialogResult.direction = (radioGroup.selectedButton === 0) ? "toName" : "toLabel";
            dialogResult.removeAllSpaces = trimSpacesCheckbox.checkedState;
        }
        
        dialog.destroy();
        return dialogResult;
    }
    
    // ==================== PROCESS LABEL TO NAME ====================
    function processLabelToName(objects) {
        var results = {
            totalProcessed: 0,
            successCount: 0,
            blankCount: 0,
            skippedCount: 0,
            details: []
        };
        
        for (var i = 0; i < objects.length; i++) {
            var obj = objects[i];
            
            // Check if the object can be renamed
            if (!obj.hasOwnProperty("name") || !obj.hasOwnProperty("label")) {
                results.skippedCount++;
                continue;
            }
            
            results.totalProcessed++;
            var oldName = obj.name;
            var scriptLabel = obj.label;
            var newName;
            
            // Determine new name
            if (scriptLabel && scriptLabel.length > 0) {
                newName = scriptLabel;
                obj.name = newName;
                results.successCount++;
                
                if (CONFIG.showDetailedReport) {
                    results.details.push({
                        oldName: oldName,
                        newName: newName,
                        type: getObjectType(obj),
                        action: "Renamed from Label"
                    });
                }
            } else {
                newName = CONFIG.defaultBlankName;
                obj.name = newName;
                results.blankCount++;
                
                if (CONFIG.showDetailedReport) {
                    results.details.push({
                        oldName: oldName,
                        newName: newName,
                        type: getObjectType(obj),
                        action: "No label found"
                    });
                }
            }
        }
        
        return results;
    }
    
    // ==================== PROCESS NAME TO LABEL ====================
    function processNameToLabel(objects) {
        var results = {
            totalProcessed: 0,
            successCount: 0,
            cleanedCount: 0,
            skippedCount: 0,
            emptyCount: 0,
            details: []
        };
        
        for (var i = 0; i < objects.length; i++) {
            var obj = objects[i];
            
            // Check if the object has the required properties
            if (!obj.hasOwnProperty("name") || !obj.hasOwnProperty("label")) {
                results.skippedCount++;
                continue;
            }
            
            results.totalProcessed++;
            var objectName = obj.name;
            var oldLabel = obj.label;
            
            // Skip if object has no name or generic name
            if (!objectName || objectName.length === 0 || isGenericName(objectName)) {
                results.emptyCount++;
                
                if (CONFIG.showDetailedReport) {
                    results.details.push({
                        objectName: objectName || "(unnamed)",
                        oldLabel: oldLabel,
                        newLabel: "(skipped - generic name)",
                        type: getObjectType(obj),
                        action: "Skipped"
                    });
                }
                continue;
            }
            
            // Clean and process the name
            var cleanedName = objectName;
            var wasCleaned = false;
            
            if (CONFIG.cleanText) {
                cleanedName = cleanObjectName(objectName);
                wasCleaned = (cleanedName !== objectName);
            }
            
            // Apply to Script Label
            obj.label = cleanedName;
            results.successCount++;
            
            if (wasCleaned) {
                results.cleanedCount++;
            }
            
            if (CONFIG.showDetailedReport) {
                results.details.push({
                    objectName: objectName,
                    oldLabel: oldLabel,
                    newLabel: cleanedName,
                    type: getObjectType(obj),
                    action: wasCleaned ? "Applied & Cleaned" : "Applied"
                });
            }
        }
        
        return results;
    }
    
    // ==================== TEXT CLEANING FUNCTIONS ====================
    function cleanObjectName(text) {
        if (!text || text.length === 0) {
            return "";
        }
        
        var cleaned = text;
        
        // Remove ALL spaces if option is selected
        if (CONFIG.removeAllSpaces) {
            cleaned = cleaned.replace(/\s+/g, '');
        } else {
            // Trim whitespace
            if (CONFIG.trimWhitespace) {
                cleaned = trimString(cleaned);
            }
        }
        
        // Remove or replace invalid characters
        if (CONFIG.removeInvalidChars) {
            // Remove control characters and problematic Unicode
            cleaned = cleaned.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
            
            // Replace multiple spaces with single space (only if not removing all spaces)
            if (!CONFIG.removeAllSpaces) {
                cleaned = cleaned.replace(/\s+/g, ' ');
            }
            
            // Remove leading/trailing special characters
            cleaned = cleaned.replace(/^[^\w]+|[^\w]+$/g, '');
        }
        
        // Limit length
        if (CONFIG.maxLabelLength && cleaned.length > CONFIG.maxLabelLength) {
            cleaned = cleaned.substring(0, CONFIG.maxLabelLength);
        }
        
        // Final trim (unless we're removing all spaces)
        if (!CONFIG.removeAllSpaces) {
            cleaned = trimString(cleaned);
        }
        
        return cleaned;
    }
    
    function isGenericName(name) {
        // Check if name is a generic InDesign default name
        var genericPatterns = [
            /^<.*>$/,                    // <Rectangle>, <TextFrame>, etc.
            /^(Rectangle|Oval|Polygon|Group|TextFrame|GraphicLine)\s*\d*$/i,
            /^(Path|Frame|Image)\s*\d*$/i,
            /^Layer \d+$/i,
            /^_BLANK$/i
        ];
        
        for (var i = 0; i < genericPatterns.length; i++) {
            if (genericPatterns[i].test(name)) {
                return true;
            }
        }
        
        return false;
    }
    
    function trimString(str) {
        return str.replace(/^\s+|\s+$/g, '');
    }
    
    // ==================== GET ALL PAGE ITEMS ====================
    function getAllPageItems(page) {
        var items = [];
        
        if (CONFIG.processTextFrames) {
            items = items.concat(page.textFrames.everyItem().getElements());
        }
        if (CONFIG.processGraphics) {
            items = items.concat(page.rectangles.everyItem().getElements());
            items = items.concat(page.ovals.everyItem().getElements());
            items = items.concat(page.polygons.everyItem().getElements());
            items = items.concat(page.graphicLines.everyItem().getElements());
        }
        if (CONFIG.processGroups) {
            items = items.concat(page.groups.everyItem().getElements());
        }
        
        return items;
    }
    
    // ==================== GET OBJECT TYPE ====================
    function getObjectType(obj) {
        if (obj.constructor.name === "TextFrame") return "Text Frame";
        if (obj.constructor.name === "Rectangle") return "Rectangle";
        if (obj.constructor.name === "Oval") return "Oval";
        if (obj.constructor.name === "Polygon") return "Polygon";
        if (obj.constructor.name === "GraphicLine") return "Line";
        if (obj.constructor.name === "Group") return "Group";
        return "Object";
    }
    
    // ==================== SHOW REPORT ====================
    function showReport(results, direction) {
        var report = "Sync Script Labels and Object Names - Report\n";
        report += repeatString("=", 55) + "\n\n";
        
        // Direction indicator
        if (direction === "toName") {
            report += "DIRECTION: Script Label --> Object Name\n";
        } else {
            report += "DIRECTION: Object Name --> Script Label\n";
        }
        
        // Add info about space removal if enabled
        if (CONFIG.removeAllSpaces && direction === "toLabel") {
            report += "OPTIONS: Remove all spaces enabled\n";
        }
        
        report += repeatString("-", 55) + "\n\n";
        
        report += "SUMMARY:\n";
        report += repeatString("-", 55) + "\n";
        report += "Total objects processed: " + results.totalProcessed + "\n";
        
        if (direction === "toName") {
            report += "  - Renamed from Script Label: " + results.successCount + "\n";
            report += "  - Named as '" + CONFIG.defaultBlankName + "': " + results.blankCount + "\n";
        } else {
            report += "  - Script Labels applied: " + results.successCount + "\n";
            if (results.cleanedCount > 0) {
                report += "  - Text cleaned/trimmed: " + results.cleanedCount + "\n";
            }
            if (results.emptyCount > 0) {
                report += "  - Skipped (generic names): " + results.emptyCount + "\n";
            }
        }
        
        if (results.skippedCount > 0) {
            report += "  - Skipped (not compatible): " + results.skippedCount + "\n";
        }
        
        // Add detailed list if enabled and there are items
        if (CONFIG.showDetailedReport && results.details.length > 0) {
            report += "\n" + repeatString("-", 55) + "\n";
            report += "DETAILS:\n";
            report += repeatString("-", 55) + "\n";
            
            for (var i = 0; i < results.details.length; i++) {
                var detail = results.details[i];
                report += (i + 1) + ". " + detail.type + " - " + detail.action + "\n";
                
                if (direction === "toName") {
                    report += "   Old Name: " + (detail.oldName || "(unnamed)") + "\n";
                    report += "   New Name: " + detail.newName + "\n";
                } else {
                    report += "   Object Name: " + detail.objectName + "\n";
                    if (detail.oldLabel && detail.oldLabel.length > 0) {
                        report += "   Old Label: " + detail.oldLabel + "\n";
                    }
                    report += "   New Label: " + detail.newLabel + "\n";
                }
                report += "\n";
            }
        }
        
        report += repeatString("=", 55);
        
        // Create scrollable dialog for report
        showScrollableReport(report);
    }
    
    // ==================== SHOW SCROLLABLE REPORT ====================
    function showScrollableReport(reportText) {
        var dialog = app.dialogs.add({name: "Processing Report"});
        
        with(dialog.dialogColumns.add()) {
            with(dialogRows.add()) {
                var textBox = textEditboxes.add({
                    editContents: reportText,
                    minWidth: 600,
                    charactersAndLines: [80, 25]  // 80 chars wide, 25 lines tall
                });
            }
        }
        
        dialog.show();
        dialog.destroy();
    }
    
    // ==================== UTILITY FUNCTIONS ====================
    function repeatString(str, count) {
        var result = '';
        for (var i = 0; i < count; i++) {
            result += str;
        }
        return result;
    }
    
    // ==================== RUN SCRIPT ====================
    main();
    
})();
