/**
 * Create Paragraph Styles from Text Frames - FIXED & ENHANCED
 * 
 * Creates paragraph styles from all text in the document,
 * naming them by the first word of each paragraph.
 * Captures ALL formatting including alignment, indents, spacing, and character formatting.
 *
 * @author Adobe InDesign Script
 * @version 2.0
 * @compatibility InDesign CC 2019+
 */

#target 'indesign'

(function () {
    
    // ==================== CHECK DOCUMENT ====================
    if (!app.documents.length) {
        alert("Please open a document before running this script.");
        return;
    }

    var doc = app.activeDocument;
    
    // ==================== CONFIGURATION ====================
    var CONFIG = {
        maxStyleNameLength: 40,
        useFirstWord: true,          // Use first word of paragraph as style name
        skipEmptyParagraphs: true,   // Skip paragraphs with only whitespace
        showProgressBar: true,       // Show progress window
        applyStylesToSource: true,   // Apply created styles back to source paragraphs
        reportResults: true          // Show final report
    };
    
    // ==================== HELPER FUNCTIONS ====================
    
    // Execute with undo support
    function withUndo(name, fn) {
        app.doScript(fn, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, name);
    }
    
    // Extract first word from paragraph
    function getFirstWord(para) {
        var text = String(para.contents).replace(/[\r\n]+/g, " ");
        var match = text.match(/([^\s]+)/);
        return match ? match[1] : "Style";
    }
    
    // Clean and sanitize style name
    function sanitizeStyleName(name) {
        var clean = String(name)
            .replace(/[^\w\-\s]+/g, "")      // Remove special chars except word chars, dash, space
            .replace(/\s+/g, " ")             // Collapse multiple spaces
            .replace(/^ +| +$/g, "");         // Manual trim (no .trim() for old ES)
        
        if (clean.length > CONFIG.maxStyleNameLength) {
            clean = clean.substr(0, CONFIG.maxStyleNameLength);
        }
        
        return clean || "Style";
    }
    
    // Generate unique style name
    function getUniqueStyleName(baseName) {
        var name = baseName;
        var counter = 1;
        
        while (doc.paragraphStyles.itemByName(name).isValid) {
            counter++;
            name = baseName + "-" + counter;
        }
        
        return name;
    }
    
    // ==================== CREATE STYLE FROM PARAGRAPH ====================
    function createStyleFromParagraph(para, stats) {
        try {
            // Get style name from paragraph
            var baseName = sanitizeStyleName(
                CONFIG.useFirstWord ? getFirstWord(para) : String(para.contents).substr(0, 30)
            );
            
            // Check if style exists
            var styleName;
            if (doc.paragraphStyles.itemByName(baseName).isValid) {
                styleName = getUniqueStyleName(baseName);
                stats.duplicateNames++;
            } else {
                styleName = baseName;
            }
            
            // Create the paragraph style
            var style = doc.paragraphStyles.add({name: styleName});
            stats.stylesCreated++;
            
            // Get first character for character-level formatting
            var firstChar = para.characters[0];
            
            // ==================== CHARACTER FORMATTING ====================
            if (firstChar && firstChar.isValid) {
                try { style.appliedFont = firstChar.appliedFont; } catch (e) {}
                try { style.fontStyle = firstChar.fontStyle; } catch (e) {}
                try { style.pointSize = firstChar.pointSize; } catch (e) {}
                try { style.leading = firstChar.leading; } catch (e) {}
                try { style.kerningMethod = firstChar.kerningMethod; } catch (e) {}
                try { style.tracking = firstChar.tracking; } catch (e) {}
                try { style.horizontalScale = firstChar.horizontalScale; } catch (e) {}
                try { style.verticalScale = firstChar.verticalScale; } catch (e) {}
                try { style.baselineShift = firstChar.baselineShift; } catch (e) {}
                try { style.skew = firstChar.skew; } catch (e) {}
                try { style.fillColor = firstChar.fillColor; } catch (e) {}
                try { style.strokeColor = firstChar.strokeColor; } catch (e) {}
                try { style.strokeWeight = firstChar.strokeWeight; } catch (e) {}
                try { style.underline = firstChar.underline; } catch (e) {}
                try { style.strikeThru = firstChar.strikeThru; } catch (e) {}
                try { style.capitalization = firstChar.capitalization; } catch (e) {}
                try { style.position = firstChar.position; } catch (e) {}
                
                // OpenType features
                try { style.ligatures = firstChar.ligatures; } catch (e) {}
                try { style.otfContextualAlternate = firstChar.otfContextualAlternate; } catch (e) {}
                try { style.otfDiscretionaryLigature = firstChar.otfDiscretionaryLigature; } catch (e) {}
            }
            
            // ==================== PARAGRAPH FORMATTING ====================
            
            // Alignment and justification
            try { style.justification = para.justification; } catch (e) {}
            
            // Indents
            try { style.leftIndent = para.leftIndent; } catch (e) {}
            try { style.rightIndent = para.rightIndent; } catch (e) {}
            try { style.firstLineIndent = para.firstLineIndent; } catch (e) {}
            
            // Spacing
            try { style.spaceBefore = para.spaceBefore; } catch (e) {}
            try { style.spaceAfter = para.spaceAfter; } catch (e) {}
            
            // Tab stops
            try { style.tabList = para.tabList; } catch (e) {}
            
            // Hyphenation
            try { style.hyphenation = para.hyphenation; } catch (e) {}
            try { style.hyphenateCapitalizedWords = para.hyphenateCapitalizedWords; } catch (e) {}
            
            // Keep options
            try { style.keepAllLinesTogether = para.keepAllLinesTogether; } catch (e) {}
            try { style.keepLinesTogether = para.keepLinesTogether; } catch (e) {}
            try { style.keepFirstLines = para.keepFirstLines; } catch (e) {}
            try { style.keepLastLines = para.keepLastLines; } catch (e) {}
            try { style.keepWithNext = para.keepWithNext; } catch (e) {}
            
            // Paragraph rules (lines above/below)
            try {
                style.ruleAbove = para.ruleAbove;
                if (para.ruleAbove) {
                    try { style.ruleAboveColor = para.ruleAboveColor; } catch (e) {}
                    try { style.ruleAboveLineWeight = para.ruleAboveLineWeight; } catch (e) {}
                    try { style.ruleAboveOffset = para.ruleAboveOffset; } catch (e) {}
                    try { style.ruleAboveWidth = para.ruleAboveWidth; } catch (e) {}
                    try { style.ruleAboveType = para.ruleAboveType; } catch (e) {}
                }
            } catch (e) {}
            
            try {
                style.ruleBelow = para.ruleBelow;
                if (para.ruleBelow) {
                    try { style.ruleBelowColor = para.ruleBelowColor; } catch (e) {}
                    try { style.ruleBelowLineWeight = para.ruleBelowLineWeight; } catch (e) {}
                    try { style.ruleBelowOffset = para.ruleBelowOffset; } catch (e) {}
                    try { style.ruleBelowWidth = para.ruleBelowWidth; } catch (e) {}
                    try { style.ruleBelowType = para.ruleBelowType; } catch (e) {}
                }
            } catch (e) {}
            
            // Drop caps
            try { style.dropCapCharacters = para.dropCapCharacters; } catch (e) {}
            try { style.dropCapLines = para.dropCapLines; } catch (e) {}
            
            // Bullets and numbering
            try { style.bulletsAndNumberingListType = para.bulletsAndNumberingListType; } catch (e) {}
            
            // Apply the style back to the paragraph if configured
            if (CONFIG.applyStylesToSource) {
                try {
                    para.appliedParagraphStyle = style;
                    stats.stylesApplied++;
                } catch (e) {
                    stats.errors.push("Could not apply style '" + styleName + "' to paragraph: " + e.message);
                }
            }
            
            // Record detail
            stats.details.push({
                styleName: styleName,
                text: String(para.contents).substr(0, 50) + (para.contents.length > 50 ? "..." : ""),
                status: "Created & Applied"
            });
            
            return style;
            
        } catch (e) {
            stats.errors.push("Error creating style from paragraph: " + e.message);
            return null;
        }
    }
    
    // ==================== MAIN PROCESSING ====================
    function processDocument() {
        var stats = {
            totalFrames: 0,
            totalParagraphs: 0,
            stylesCreated: 0,
            stylesApplied: 0,
            paragraphsSkipped: 0,
            duplicateNames: 0,
            errors: [],
            details: []
        };
        
        var frames = doc.textFrames.everyItem().getElements();
        stats.totalFrames = frames.length;
        
        // Progress window
        var progressWin = null;
        var progressBar = null;
        var progressLabel = null;
        
        if (CONFIG.showProgressBar) {
            progressWin = new Window("palette", "Creating Paragraph Styles", undefined, {closeButton: false});
            progressBar = progressWin.add("progressbar", undefined, 0, frames.length);
            progressBar.preferredSize = [300, 20];
            progressLabel = progressWin.add("statictext", undefined, "Starting...");
            progressLabel.preferredSize = [300, 20];
            progressWin.show();
        }
        
        // Process each text frame
        for (var i = 0; i < frames.length; i++) {
            var frame = frames[i];
            
            try {
                var paragraphs = frame.paragraphs.everyItem().getElements();
                
                for (var j = 0; j < paragraphs.length; j++) {
                    var para = paragraphs[j];
                    stats.totalParagraphs++;
                    
                    // Skip empty paragraphs
                    var content = String(para.contents).replace(/[\r\n\s]+/g, "");
                    if (CONFIG.skipEmptyParagraphs && content.length === 0) {
                        stats.paragraphsSkipped++;
                        continue;
                    }
                    
                    // Create style from paragraph
                    createStyleFromParagraph(para, stats);
                }
                
            } catch (e) {
                stats.errors.push("Error processing frame " + (i + 1) + ": " + e.message);
            }
            
            // Update progress
            if (CONFIG.showProgressBar && progressBar && progressLabel) {
                progressBar.value = i + 1;
                progressLabel.text = "Processed frame " + (i + 1) + " of " + frames.length;
                progressWin.update();
            }
        }
        
        // Close progress window
        if (progressWin) {
            progressWin.close();
        }
        
        return stats;
    }
    
    // ==================== SHOW REPORT ====================
    function showReport(stats) {
        var report = "Create Paragraph Styles - Report\n";
        report += repeatString("=", 60) + "\n\n";
        
        report += "SUMMARY:\n";
        report += repeatString("-", 60) + "\n";
        report += "Text frames processed: " + stats.totalFrames + "\n";
        report += "Paragraphs processed: " + stats.totalParagraphs + "\n";
        report += "Paragraph styles created: " + stats.stylesCreated + "\n";
        
        if (CONFIG.applyStylesToSource) {
            report += "Styles applied to source: " + stats.stylesApplied + "\n";
        }
        
        if (stats.paragraphsSkipped > 0) {
            report += "Paragraphs skipped (empty): " + stats.paragraphsSkipped + "\n";
        }
        
        if (stats.duplicateNames > 0) {
            report += "Duplicate names handled: " + stats.duplicateNames + "\n";
        }
        
        if (stats.errors.length > 0) {
            report += "Errors encountered: " + stats.errors.length + "\n";
        }
        
        // Errors section
        if (stats.errors.length > 0) {
            report += "\n" + repeatString("-", 60) + "\n";
            report += "ERRORS:\n";
            report += repeatString("-", 60) + "\n";
            var errorLimit = Math.min(stats.errors.length, 10);
            for (var i = 0; i < errorLimit; i++) {
                report += (i + 1) + ". " + stats.errors[i] + "\n";
            }
            if (stats.errors.length > 10) {
                report += "... and " + (stats.errors.length - 10) + " more errors\n";
            }
        }
        
        // Details section (first 20 styles)
        if (stats.details.length > 0) {
            report += "\n" + repeatString("-", 60) + "\n";
            report += "STYLES CREATED (first 20):\n";
            report += repeatString("-", 60) + "\n";
            var detailLimit = Math.min(stats.details.length, 20);
            for (var j = 0; j < detailLimit; j++) {
                var detail = stats.details[j];
                report += (j + 1) + ". " + detail.styleName + "\n";
                report += "   Text: " + detail.text + "\n";
                report += "   Status: " + detail.status + "\n\n";
            }
            if (stats.details.length > 20) {
                report += "... and " + (stats.details.length - 20) + " more styles\n";
            }
        }
        
        report += "\n" + repeatString("=", 60) + "\n";
        report += "COMPLETE\n";
        report += "All formatting captured including:\n";
        report += "- Font, size, leading, tracking, scaling\n";
        report += "- Alignment and indents\n";
        report += "- Space before/after\n";
        report += "- Tab stops\n";
        report += "- Paragraph rules (lines above/below)\n";
        report += "- Keep options and hyphenation\n";
        report += repeatString("=", 60);
        
        // Show in scrollable dialog
        showScrollableReport(report);
    }
    
    function showScrollableReport(reportText) {
        var dialog = app.dialogs.add({name: "Style Creation Report"});
        
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
    
    function repeatString(str, count) {
        var result = '';
        for (var i = 0; i < count; i++) {
            result += str;
        }
        return result;
    }
    
    // ==================== RUN SCRIPT ====================
    withUndo("Create Paragraph Styles", function() {
        var stats = processDocument();
        
        if (CONFIG.reportResults) {
            showReport(stats);
        } else {
            if (stats.errors.length > 0) {
                alert("Completed with " + stats.errors.length + " errors.\n\n" +
                      "Styles created: " + stats.stylesCreated);
            } else {
                alert("Style creation complete!\n\n" +
                      "Styles created: " + stats.stylesCreated + "\n" +
                      "Paragraphs processed: " + stats.totalParagraphs);
            }
        }
    });
    
})();
