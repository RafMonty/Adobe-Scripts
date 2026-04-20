/**
 * Create Character Styles from Text Frame Lines
 * 
 * This script:
 * 1. Requires a text frame to be selected
 * 2. Optionally deletes all existing character & paragraph styles
 * 3. Creates a new character style for each line of text
 * 4. Uses the line text as the style name (cleaned and trimmed)
 * 5. Applies all formatting from the original text (spacing, leading, scaling, etc.)
 * 6. Applies the new style back to its source line
 * 7. Reports number of styles created and any issues
 *
 * @author Adobe InDesign Script
 * @version 1.0
 * @compatibility InDesign CC 2024+
 */

(function() {
    
    // ==================== CONFIGURATION ====================
    var CONFIG = {
        cleanStyleNames: true,        // Remove/clean spaces and special chars
        removeAllSpaces: true,        // Remove all spaces from style names
        maxStyleNameLength: 100,      // Maximum style name length
        skipEmptyLines: true,         // Skip blank lines
        reportDetails: true           // Show detailed report
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
        
        // Validate selection - must be exactly one text frame
        if (selection.length === 0) {
            alert("Please select a text frame before running this script.");
            return;
        }
        
        if (selection.length > 1) {
            alert("Please select only ONE text frame.\n\nYou have " + selection.length + " objects selected.");
            return;
        }
        
        var textFrame = selection[0];
        
        // Verify it's a text frame
        if (textFrame.constructor.name !== "TextFrame") {
            alert("Please select a TEXT FRAME.\n\nYou have selected: " + textFrame.constructor.name);
            return;
        }
        
        // Check if text frame has content
        if (!textFrame.contents || textFrame.contents.length === 0) {
            alert("The selected text frame is empty.\n\nPlease select a text frame with content.");
            return;
        }
        
        // Show options dialog
        var options = showOptionsDialog();
        
        if (options === null) {
            return; // User cancelled
        }
        
        // Delete existing styles if requested
        var deletionReport = null;
        if (options.deleteExistingStyles) {
            deletionReport = deleteAllStyles(doc);
        }
        
        // Process the text frame and create styles
        var results = createStylesFromLines(doc, textFrame);
        
        // Show report
        showReport(results, deletionReport);
    }
    
    // ==================== SHOW OPTIONS DIALOG ====================
    function showOptionsDialog() {
        var dialog = app.dialogs.add({name: "Create Character Styles from Text Frame"});
        
        with(dialog.dialogColumns.add()) {
            with(dialogRows.add()) {
                staticTexts.add({
                    staticLabel: "This script will create character styles from each line",
                    minWidth: 400
                });
            }
            with(dialogRows.add()) {
                staticTexts.add({
                    staticLabel: "of text in the selected text frame.",
                    minWidth: 400
                });
            }
            
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: ""});
            }
            
            with(borderPanels.add()) {
                with(dialogColumns.add()) {
                    with(dialogRows.add()) {
                        staticTexts.add({staticLabel: "Options:"});
                    }
                    with(dialogRows.add()) {
                        var deleteStylesCheckbox = checkboxControls.add({
                            staticLabel: "Delete ALL existing Character & Paragraph Styles first",
                            checkedState: false
                        });
                    }
                    with(dialogRows.add()) {
                        staticTexts.add({
                            staticLabel: "  (Warning: This will remove all styles from document,",
                            minWidth: 400
                        });
                    }
                    with(dialogRows.add()) {
                        staticTexts.add({
                            staticLabel: "   converting styled text to 'No Style' with appearance preserved)",
                            minWidth: 400
                        });
                    }
                }
            }
            
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: ""});
            }
            
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: "Style Creation:"});
            }
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: "  - Each line will become a character style"});
            }
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: "  - Line text (cleaned) becomes style name"});
            }
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: "  - All formatting will be captured and applied"});
            }
            with(dialogRows.add()) {
                staticTexts.add({staticLabel: "  - Style will be applied back to source line"});
            }
        }
        
        var result = dialog.show();
        
        var options = null;
        if (result === true) {
            options = {
                deleteExistingStyles: deleteStylesCheckbox.checkedState
            };
        }
        
        dialog.destroy();
        return options;
    }
    
    // ==================== DELETE ALL STYLES ====================
    function deleteAllStyles(doc) {
        var report = {
            characterStylesDeleted: 0,
            paragraphStylesDeleted: 0,
            errors: []
        };
        
        try {
            // Delete all character styles (except [None])
            var charStyles = doc.characterStyles.everyItem().getElements();
            for (var i = charStyles.length - 1; i >= 0; i--) {
                if (charStyles[i].name !== "[None]") {
                    try {
                        charStyles[i].remove();
                        report.characterStylesDeleted++;
                    } catch (e) {
                        report.errors.push("Could not delete character style: " + charStyles[i].name);
                    }
                }
            }
            
            // Delete all paragraph styles (except [No Paragraph Style] and [Basic Paragraph])
            var paraStyles = doc.paragraphStyles.everyItem().getElements();
            for (var j = paraStyles.length - 1; j >= 0; j--) {
                if (paraStyles[j].name !== "[No Paragraph Style]" && 
                    paraStyles[j].name !== "[Basic Paragraph]") {
                    try {
                        paraStyles[j].remove();
                        report.paragraphStylesDeleted++;
                    } catch (e) {
                        report.errors.push("Could not delete paragraph style: " + paraStyles[j].name);
                    }
                }
            }
        } catch (e) {
            report.errors.push("Error during style deletion: " + e.message);
        }
        
        return report;
    }
    
    // ==================== CREATE STYLES FROM LINES ====================
    function createStylesFromLines(doc, textFrame) {
        var results = {
            totalLines: 0,
            stylesCreated: 0,
            linesSkipped: 0,
            errors: [],
            warnings: [],
            details: []
        };
        
        try {
            // Get all paragraphs (lines) from the text frame
            var paragraphs = textFrame.paragraphs;
            results.totalLines = paragraphs.length;
            
            // Process each paragraph/line
            for (var i = 0; i < paragraphs.length; i++) {
                var para = paragraphs[i];
                var lineText = para.contents;
                
                // Skip empty lines if configured
                if (CONFIG.skipEmptyLines && (!lineText || trimString(lineText).length === 0)) {
                    results.linesSkipped++;
                    if (CONFIG.reportDetails) {
                        results.details.push({
                            lineNumber: i + 1,
                            lineText: "(empty line)",
                            styleName: null,
                            status: "Skipped - empty"
                        });
                    }
                    continue;
                }
                
                // Clean the style name
                var styleName = lineText;
                if (CONFIG.cleanStyleNames) {
                    styleName = cleanStyleName(styleName);
                }
                
                // Check for empty style name after cleaning
                if (!styleName || styleName.length === 0) {
                    results.linesSkipped++;
                    results.warnings.push("Line " + (i + 1) + ": Style name empty after cleaning");
                    if (CONFIG.reportDetails) {
                        results.details.push({
                            lineNumber: i + 1,
                            lineText: lineText.substring(0, 30) + "...",
                            styleName: null,
                            status: "Skipped - no valid style name"
                        });
                    }
                    continue;
                }
                
                try {
                    // Check if style already exists
                    var existingStyle = null;
                    try {
                        existingStyle = doc.characterStyles.itemByName(styleName);
                        if (existingStyle.isValid) {
                            // Style exists, we'll update it
                            results.warnings.push("Line " + (i + 1) + ": Style '" + styleName + "' already exists - updating");
                        } else {
                            existingStyle = null;
                        }
                    } catch (e) {
                        existingStyle = null;
                    }
                    
                    // Create or get the character style
                    var charStyle;
                    if (existingStyle) {
                        charStyle = existingStyle;
                    } else {
                        charStyle = doc.characterStyles.add({name: styleName});
                        results.stylesCreated++;
                    }
                    
                    // Capture ALL formatting from the first character of this line
                    // We'll use the first character as the template for the style
                    var firstChar = para.characters[0];
                    
                    if (firstChar) {
                        // Apply text formatting properties
                        try {
                            charStyle.appliedFont = firstChar.appliedFont;
                            charStyle.fontStyle = firstChar.fontStyle;
                            charStyle.pointSize = firstChar.pointSize;
                            charStyle.leading = firstChar.leading;
                            charStyle.kerningMethod = firstChar.kerningMethod;
                            charStyle.tracking = firstChar.tracking;
                            charStyle.horizontalScale = firstChar.horizontalScale;
                            charStyle.verticalScale = firstChar.verticalScale;
                            charStyle.baselineShift = firstChar.baselineShift;
                            charStyle.skew = firstChar.skew;
                            charStyle.fillColor = firstChar.fillColor;
                            charStyle.strokeColor = firstChar.strokeColor;
                            charStyle.strokeWeight = firstChar.strokeWeight;
                            charStyle.underline = firstChar.underline;
                            charStyle.strikeThru = firstChar.strikeThru;
                            charStyle.capitalization = firstChar.capitalization;
                            charStyle.position = firstChar.position;
                            
                            // OpenType features if available
                            try {
                                charStyle.ligatures = firstChar.ligatures;
                                charStyle.otfContextualAlternate = firstChar.otfContextualAlternate;
                                charStyle.otfDiscretionaryLigature = firstChar.otfDiscretionaryLigature;
                            } catch (e) {
                                // OpenType features may not be available
                            }
                        } catch (e) {
                            results.warnings.push("Line " + (i + 1) + ": Could not capture all formatting - " + e.message);
                        }
                    }
                    
                    // Apply the style back to the paragraph
                    para.appliedCharacterStyle = charStyle;
                    
                    if (CONFIG.reportDetails) {
                        results.details.push({
                            lineNumber: i + 1,
                            lineText: lineText.substring(0, 30) + (lineText.length > 30 ? "..." : ""),
                            styleName: styleName,
                            status: existingStyle ? "Updated & Applied" : "Created & Applied"
                        });
                    }
                    
                } catch (e) {
                    results.errors.push("Line " + (i + 1) + ": " + e.message);
                    if (CONFIG.reportDetails) {
                        results.details.push({
                            lineNumber: i + 1,
                            lineText: lineText.substring(0, 30) + "...",
                            styleName: styleName,
                            status: "ERROR - " + e.message
                        });
                    }
                }
            }
            
        } catch (e) {
            results.errors.push("Critical error: " + e.message);
        }
        
        return results;
    }
    
    // ==================== CLEAN STYLE NAME ====================
    function cleanStyleName(text) {
        if (!text || text.length === 0) {
            return "";
        }
        
        var cleaned = text;
        
        // Trim whitespace
        cleaned = trimString(cleaned);
        
        // Remove all spaces if configured
        if (CONFIG.removeAllSpaces) {
            cleaned = cleaned.replace(/\s+/g, '');
        } else {
            // Replace multiple spaces with single space
            cleaned = cleaned.replace(/\s+/g, ' ');
        }
        
        // Remove control characters and problematic Unicode
        cleaned = cleaned.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
        
        // Remove characters that are problematic in style names
        // InDesign doesn't allow certain characters in style names
        cleaned = cleaned.replace(/[\/:*?"<>|]/g, '');
        
        // Remove leading/trailing special characters
        cleaned = cleaned.replace(/^[^\w]+|[^\w]+$/g, '');
        
        // Limit length
        if (CONFIG.maxStyleNameLength && cleaned.length > CONFIG.maxStyleNameLength) {
            cleaned = cleaned.substring(0, CONFIG.maxStyleNameLength);
        }
        
        // Final trim
        cleaned = trimString(cleaned);
        
        return cleaned;
    }
    
    // ==================== UTILITY FUNCTIONS ====================
    function trimString(str) {
        return str.replace(/^\s+|\s+$/g, '');
    }
    
    function repeatString(str, count) {
        var result = '';
        for (var i = 0; i < count; i++) {
            result += str;
        }
        return result;
    }
    
    // ==================== SHOW REPORT ====================
    function showReport(results, deletionReport) {
        var report = "Create Character Styles from Text Frame - Report\n";
        report += repeatString("=", 60) + "\n\n";
        
        // Deletion report if styles were deleted
        if (deletionReport) {
            report += "STYLE DELETION:\n";
            report += repeatString("-", 60) + "\n";
            report += "Character styles deleted: " + deletionReport.characterStylesDeleted + "\n";
            report += "Paragraph styles deleted: " + deletionReport.paragraphStylesDeleted + "\n";
            if (deletionReport.errors.length > 0) {
                report += "Deletion errors: " + deletionReport.errors.length + "\n";
            }
            report += "\n";
        }
        
        // Style creation summary
        report += "STYLE CREATION SUMMARY:\n";
        report += repeatString("-", 60) + "\n";
        report += "Total lines processed: " + results.totalLines + "\n";
        report += "Character styles created: " + results.stylesCreated + "\n";
        report += "Lines skipped: " + results.linesSkipped + "\n";
        
        if (results.warnings.length > 0) {
            report += "Warnings: " + results.warnings.length + "\n";
        }
        
        if (results.errors.length > 0) {
            report += "Errors: " + results.errors.length + "\n";
        }
        
        // Warnings section
        if (results.warnings.length > 0) {
            report += "\n" + repeatString("-", 60) + "\n";
            report += "WARNINGS:\n";
            report += repeatString("-", 60) + "\n";
            for (var w = 0; w < Math.min(results.warnings.length, 10); w++) {
                report += "  - " + results.warnings[w] + "\n";
            }
            if (results.warnings.length > 10) {
                report += "  ... and " + (results.warnings.length - 10) + " more warnings\n";
            }
        }
        
        // Errors section
        if (results.errors.length > 0) {
            report += "\n" + repeatString("-", 60) + "\n";
            report += "ERRORS:\n";
            report += repeatString("-", 60) + "\n";
            for (var e = 0; e < Math.min(results.errors.length, 10); e++) {
                report += "  - " + results.errors[e] + "\n";
            }
            if (results.errors.length > 10) {
                report += "  ... and " + (results.errors.length - 10) + " more errors\n";
            }
        }
        
        // Deletion errors
        if (deletionReport && deletionReport.errors.length > 0) {
            report += "\n" + repeatString("-", 60) + "\n";
            report += "DELETION ERRORS:\n";
            report += repeatString("-", 60) + "\n";
            for (var de = 0; de < Math.min(deletionReport.errors.length, 5); de++) {
                report += "  - " + deletionReport.errors[de] + "\n";
            }
            if (deletionReport.errors.length > 5) {
                report += "  ... and " + (deletionReport.errors.length - 5) + " more errors\n";
            }
        }
        
        // Details section
        if (CONFIG.reportDetails && results.details.length > 0) {
            report += "\n" + repeatString("-", 60) + "\n";
            report += "DETAILS:\n";
            report += repeatString("-", 60) + "\n";
            
            var maxDetails = Math.min(results.details.length, 50);
            for (var i = 0; i < maxDetails; i++) {
                var detail = results.details[i];
                report += "Line " + detail.lineNumber + ": " + detail.status + "\n";
                report += "  Text: " + detail.lineText + "\n";
                if (detail.styleName) {
                    report += "  Style: " + detail.styleName + "\n";
                }
                report += "\n";
            }
            
            if (results.details.length > 50) {
                report += "... and " + (results.details.length - 50) + " more lines\n";
            }
        }
        
        report += repeatString("=", 60);
        
        // Show scrollable report
        showScrollableReport(report);
    }
    
    // ==================== SHOW SCROLLABLE REPORT ====================
    function showScrollableReport(reportText) {
        var dialog = app.dialogs.add({name: "Processing Report"});
        
        with(dialog.dialogColumns.add()) {
            with(dialogRows.add()) {
                var textBox = textEditboxes.add({
                    editContents: reportText,
                    minWidth: 650,
                    charactersAndLines: [85, 30]
                });
            }
        }
        
        dialog.show();
        dialog.destroy();
    }
    
    // ==================== RUN SCRIPT ====================
    main();
    
})();
