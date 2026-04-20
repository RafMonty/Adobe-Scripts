﻿////////////////////////////////////////////////////
// Raf's Style Maker — UPDATED (no .trim dependency)
// Creates Paragraph Styles from text samples,
// naming them by the first word (no page prefix),
// and capturing alignment & indents + spacing.
//
// Target: InDesign 2023+ (ExtendScript)
// V1.3.1 20/08/2025
////////////////////////////////////////////////////

#target 'indesign'

(function () {
    if (!app.documents.length) {
        alert("Open a document before running this script.");
        return;
    }

    var doc = app.activeDocument;

    // ---------- Helpers ----------
    function withUndo(name, fn) {
        app.doScript(fn, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, name);
    }

    // Safe first word extractor (no .trim needed)
    function firstWordOfParagraph(para) {
        // Convert to string, normalise newlines to spaces, then pick first non‑whitespace run
        var s = String(para.contents).replace(/[\r\n]+/g, " ");
        var m = s.match(/([^\s]+)/); // first token of non‑whitespace
        return m ? m[1] : "Style";
    }

    function sanitiseName(name) {
        // Keep letters, numbers, spaces, underscore, dash; collapse spaces; cap length
        var clean = String(name)
            .replace(/[^\w\-\s]+/g, "")
            .replace(/\s+/g, " ")
            .replace(/^ +| +$/g, "") // manual trim
            .substr(0, 40);
        return clean || "Style";
    }

    function uniqueStyleName(base) {
        var name = base, i = 1;
        while (doc.paragraphStyles.itemByName(name).isValid) {
            i++;
            name = base + "-" + i;
        }
        return name;
    }

    function createOrGetStyleFromParagraph(para) {
        var baseName = sanitiseName(firstWordOfParagraph(para));
        var styleName = doc.paragraphStyles.itemByName(baseName).isValid ? uniqueStyleName(baseName) : baseName;

        var style;
        if (!doc.paragraphStyles.itemByName(styleName).isValid) {
            style = doc.paragraphStyles.add({ name: styleName });
        } else {
            style = doc.paragraphStyles.itemByName(styleName);
        }

        // ---- Copy paragraph-level formatting safely ----
        // Basic character metrics that map at paragraph-style level
        try { style.appliedFont   = para.appliedFont; } catch (_) {}
        try { style.fontStyle     = para.fontStyle; } catch (_) {}
        try { style.pointSize     = para.pointSize; } catch (_) {}
        try { style.leading       = para.leading; } catch (_) {}

        // Colour & stroke (harmless if unused)
        try { style.fillColor     = para.fillColor; } catch (_) {}
        try { style.strokeColor   = para.strokeColor; } catch (_) {}
        try { style.strokeWeight  = para.strokeWeight; } catch (_) {}

        // Spacing
        try { style.spaceBefore   = para.spaceBefore; } catch (_) {}
        try { style.spaceAfter    = para.spaceAfter; } catch (_) {}

        // Alignment & indents (your request)
        try { style.justification   = para.justification; } catch (_) {}
        try { style.leftIndent      = para.leftIndent; } catch (_) {}
        try { style.rightIndent     = para.rightIndent; } catch (_) {}
        try { style.firstLineIndent = para.firstLineIndent; } catch (_) {}

        // Horizontal scale
        try { style.horizontalScale = para.horizontalScale; } catch (_) {}

        // Tabs
        try { style.tabList = para.tabList; } catch (_) {}

        // Useful toggles
        try { style.position   = para.position; } catch (_) {}
        try { style.underline  = para.underline; } catch (_) {}
        try { style.strikeThru = para.strikeThru; } catch (_) {}

        // Paragraph rules
        try {
            style.ruleAbove           = para.ruleAbove;
            style.ruleAboveColor      = para.ruleAboveColor;
            style.ruleAboveLineWeight = para.ruleAboveLineWeight;
            style.ruleAboveOffset     = para.ruleAboveOffset;

            style.ruleBelow           = para.ruleBelow;
            style.ruleBelowColor      = para.ruleBelowColor;
            style.ruleBelowLineWeight = para.ruleBelowLineWeight;
            style.ruleBelowOffset     = para.ruleBelowOffset;
        } catch (_) {}

        return style;
    }

    // ---------- Progress UI ----------
    var frames = doc.textFrames;
    var w = new Window("palette", "Creating Paragraph Styles", undefined, { closeButton: false });
    var bar = w.add("progressbar", undefined, 0, frames.length);
    var label = w.add("statictext", undefined, "Starting…");
    w.show();

    var errors = [];

    withUndo("Create Styles", function () {
        for (var i = 0; i < frames.length; i++) {
            var tf = frames[i];
            try {
                // Skip pasteboard/master frames gracefully (no parentPage), but they still have paragraphs — that’s OK.
                var paras = tf.paragraphs;
                for (var j = 0; j < paras.length; j++) {
                    var p = paras[j];
                    // Ignore completely empty paragraphs (single return)
                    if (!String(p.contents).replace(/[\r\n\s]+/g, "").length) continue;

                    var style = createOrGetStyleFromParagraph(p);
                    try { p.appliedParagraphStyle = style; } catch (eAssign) {
                        errors.push("Could not apply style to a paragraph in frame #" + (i + 1) + ": " + eAssign);
                    }
                }
            } catch (e) {
                errors.push("Error in text frame #" + (i + 1) + ": " + e);
            }
            bar.value = i + 1;
            label.text = "Processed frame " + (i + 1) + " of " + frames.length;
            w.update();
        }
    });

    w.close();

    if (errors.length) {
        alert("Completed with some issues:\n\n" + errors.join("\n"));
    } else {
        alert("Styles creation complete.\n\n• Names no longer include page numbers.\n• Alignment & indents + spacing captured.");
    }
})();
