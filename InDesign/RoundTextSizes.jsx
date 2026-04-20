// InDesign Text Formatting Cleanup Script
// Rounds font sizes to nearest 0.5pt, leading to nearest 0.5pt, spacing to nearest 0.25pt

function main() {
    // Check if InDesign is running and has a document
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = app.selection;
    
    // Validate selection
    if (selection.length === 0) {
        alert("Please select some text first.");
        return;
    }
    
    // Check for text selection - more comprehensive approach
    var textSelection = null;
    
    // Method 1: Direct text selection
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].hasOwnProperty('contents') && selection[i].constructor.name === 'Text') {
            textSelection = selection[i];
            break;
        }
    }
    
    // Method 2: Check if we have a text frame with insertion point or text selection
    if (!textSelection) {
        try {
            // Check if there's an active text selection via the selection object
            if (app.selection.length > 0 && app.selection[0].constructor.name === 'TextFrame') {
                var textFrame = app.selection[0];
                if (textFrame.parentStory.texts.length > 0) {
                    // Check if there's a text selection within the frame
                    var story = textFrame.parentStory;
                    if (story.texts.length > 0) {
                        textSelection = story.texts[0];
                    }
                }
            }
        } catch (e) {
            // Continue to next method
        }
    }
    
    // Method 3: Try to get text selection from document
    if (!textSelection) {
        try {
            // Access the current text selection directly
            var currentSelection = doc.selection;
            if (currentSelection.length > 0) {
                for (var j = 0; j < currentSelection.length; j++) {
                    if (currentSelection[j].constructor.name === 'Text' || 
                        currentSelection[j].constructor.name === 'InsertionPoint' ||
                        currentSelection[j].hasOwnProperty('contents')) {
                        textSelection = currentSelection[j];
                        break;
                    }
                }
            }
        } catch (e) {
            // Continue
        }
    }
    
    // Method 4: Last resort - check if there's any text selection at all
    if (!textSelection) {
        try {
            if (app.selection.length > 0) {
                // Try to work with whatever is selected
                var sel = app.selection[0];
                if (sel.hasOwnProperty('parentStory') || sel.hasOwnProperty('contents')) {
                    textSelection = sel;
                }
            }
        } catch (e) {
            // Final fallback
        }
    }
    
    if (!textSelection) {
        alert("Please select text within a text frame.\n\nMake sure you have:\n1. Selected actual text (not just clicked in the frame)\n2. The text cursor is active in the text frame");
        return;
    }
    
    // Process the selection
    try {
        processTextSelection(textSelection);
        alert("Text formatting cleanup completed successfully!");
    } catch (error) {
        alert("Error processing text: " + error.message);
    }
}

function processTextSelection(textSelection) {
    // Get the parent story to work with paragraphs
    var parentStory = textSelection.parentStory;
    var selectionStart = textSelection.index;
    var selectionEnd = selectionStart + textSelection.length - 1;
    
    // Process character formatting (font sizes)
    processCharacterFormatting(textSelection);
    
    // Find and process affected paragraphs
    var processedParagraphs = [];
    
    for (var i = 0; i < parentStory.paragraphs.length; i++) {
        var para = parentStory.paragraphs[i];
        var paraStart = para.index;
        var paraEnd = paraStart + para.length - 1;
        
        // Check if this paragraph overlaps with selection
        if (paraStart <= selectionEnd && paraEnd >= selectionStart) {
            processParagraphFormatting(para);
            processedParagraphs.push(i + 1);
        }
    }
}

function processCharacterFormatting(textSelection) {
    // Process each character in the selection
    for (var i = 0; i < textSelection.characters.length; i++) {
        var character = textSelection.characters[i];
        var currentSize = character.pointSize;
        
        // Skip if point size is not set (inherited)
        if (currentSize !== NothingEnum.NOTHING) {
            var roundedSize = roundToNearestHalf(currentSize);
            if (roundedSize !== currentSize) {
                character.pointSize = roundedSize;
            }
        }
    }
}

function processParagraphFormatting(paragraph) {
    // Process space before
    if (paragraph.spaceBefore !== NothingEnum.NOTHING) {
        var roundedSpaceBefore = roundToNearestQuarter(paragraph.spaceBefore);
        if (roundedSpaceBefore !== paragraph.spaceBefore) {
            paragraph.spaceBefore = roundedSpaceBefore;
        }
    }
    
    // Process space after
    if (paragraph.spaceAfter !== NothingEnum.NOTHING) {
        var roundedSpaceAfter = roundToNearestQuarter(paragraph.spaceAfter);
        if (roundedSpaceAfter !== paragraph.spaceAfter) {
            paragraph.spaceAfter = roundedSpaceAfter;
        }
    }
    
    // Process leading - more robust detection
    try {
        var currentLeading = paragraph.leading;
        var isAutoLeading = false;
        
        // Multiple ways to detect auto leading
        if (currentLeading === Leading.AUTO) {
            isAutoLeading = true;
        } else if (currentLeading.toString().indexOf("Leading.AUTO") !== -1) {
            isAutoLeading = true;
        } else {
            // Check if it's a numeric value that might be auto-calculated
            // We'll convert auto to fixed if it's currently showing as auto in UI
            try {
                var testValue = parseFloat(currentLeading);
                if (isNaN(testValue)) {
                    isAutoLeading = true;
                }
            } catch (e) {
                isAutoLeading = true;
            }
        }
        
        if (isAutoLeading) {
            // Leading is auto - calculate actual value and set as fixed
            var actualLeading = getActualLeading(paragraph);
            if (actualLeading !== null && actualLeading > 0) {
                var roundedLeading = roundToNearestHalf(actualLeading);
                paragraph.leading = roundedLeading;
            }
        } else if (currentLeading !== NothingEnum.NOTHING && typeof currentLeading === 'number') {
            // Leading is set to a specific value - just round it
            var roundedLeading = roundToNearestHalf(currentLeading);
            if (Math.abs(roundedLeading - currentLeading) > 0.001) {
                paragraph.leading = roundedLeading;
            }
        }
    } catch (leadingError) {
        // If leading processing fails, continue with other formatting
        // Don't let leading errors stop the entire script
    }
}

function getActualLeading(paragraph) {
    // Get the actual computed leading value when it's set to auto
    try {
        // For auto leading, we need to get the computed value
        // This is tricky - we'll use a temporary override method
        var originalLeading = paragraph.leading;
        
        // Create a temporary text frame to measure the leading
        var tempFrame = paragraph.parentStory.parentTextFrames[0];
        if (tempFrame) {
            // Get the baseline positions to calculate actual leading
            var lines = tempFrame.lines;
            if (lines.length > 0) {
                // Try to get leading from the first line that contains our paragraph
                for (var i = 0; i < lines.length; i++) {
                    var line = lines[i];
                    if (line.index >= paragraph.index && line.index < paragraph.index + paragraph.length) {
                        // Get the largest font size in the line and multiply by auto leading factor
                        var maxFontSize = 0;
                        for (var j = 0; j < line.characters.length; j++) {
                            var character = line.characters[j];
                            if (character.pointSize > maxFontSize) {
                                maxFontSize = character.pointSize;
                            }
                        }
                        // InDesign's default auto leading is 120% of font size
                        return maxFontSize * 1.2;
                    }
                }
            }
        }
        
        // Fallback: estimate based on font size
        var maxFontSize = 0;
        for (var k = 0; k < paragraph.characters.length; k++) {
            var character = paragraph.characters[k];
            if (character.pointSize > maxFontSize) {
                maxFontSize = character.pointSize;
            }
        }
        return maxFontSize * 1.2; // Default auto leading multiplier
        
    } catch (e) {
        return null;
    }
}

function roundToNearestHalf(value) {
    // Round to nearest 0.5
    return Math.round(value * 2) / 2;
}

function roundToNearestQuarter(value) {
    // Round to nearest 0.25
    return Math.round(value * 4) / 4;
}

// Run the script
main();