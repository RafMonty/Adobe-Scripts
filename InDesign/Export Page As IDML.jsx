﻿﻿/**
 * Export Active Page to IDML — fixed duplicate() target
 * InDesign 2023+ (Mac/Win)
 */

(function () {
    if (app.documents.length === 0) {
        alert("No document is open.\nOpen a document and select/activate the page you want to export.");
        return;
    }

    var doc = app.activeDocument;
    var win = app.layoutWindows.length ? app.activeWindow : null;
    if (!win) {
        alert("No layout window is active.\nMake sure a document window is frontmost and a page is visible.");
        return;
    }

    var srcPage;
    try { srcPage = win.activePage; } catch (e) { if (doc.pages.length > 0) srcPage = doc.pages[0]; }
    if (!srcPage) {
        alert("Could not determine the active page.\nClick the page in the Pages panel and try again.");
        return;
    }

    var pageName = srcPage.name;

    var confirmMsg = "This will export the ACTIVE page only:\n\n" +
        "Document: " + doc.name + "\n" +
        "Page: " + pageName + "\n\nProceed?";
    if (!confirm(confirmMsg)) return;

    var b = srcPage.bounds;
    var pageWidthPt  = b[3] - b[1];
    var pageHeightPt = b[2] - b[0];

    var dp = doc.documentPreferences;
    var bleedTop     = dp.documentBleedTopOffset;
    var bleedBot     = dp.documentBleedBottomOffset;
    var bleedInside  = dp.documentBleedInsideOrLeftOffset;
    var bleedOutside = dp.documentBleedOutsideOrRightOffset;

    var tempDoc;
    app.doScript(function () {
        tempDoc = app.documents.add();
        var tdp = tempDoc.documentPreferences;
        tdp.facingPages = false;
        tdp.pagesPerDocument = 1;
        tdp.pageWidth  = pageWidthPt;
        tdp.pageHeight = pageHeightPt;
        tdp.documentBleedTopOffset            = bleedTop;
        tdp.documentBleedBottomOffset         = bleedBot;
        tdp.documentBleedInsideOrLeftOffset   = bleedInside;
        tdp.documentBleedOutsideOrRightOffset = bleedOutside;

        // >>> FIX: duplicate before the temp doc's first page (a Page), then remove the blank
        srcPage.duplicate(LocationOptions.BEFORE, tempDoc.pages[0]);
        tempDoc.pages[-1].remove(); // remove the original blank last page
        // <<<

        try {
            var mp = tempDoc.pages[0].marginPreferences;
            mp.top = mp.bottom = mp.left = mp.right = 0;
        } catch (ignore) {}
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Export Active Page to IDML");

    var baseName = doc.name.replace(/\.[^\.]+$/, "");
    var suggested = baseName + "_p" + pageName + ".idml";

    var outFile = File.saveDialog("Save selected page as IDML", "IDML:*.idml");
    if (!outFile) { if (tempDoc && tempDoc.isValid) tempDoc.close(SaveOptions.NO); return; }
    if (!/\.idml$/i.test(outFile.name)) outFile = File(outFile.fsName + ".idml");
    try {
        if (outFile.parent && !outFile.exists && outFile.displayName.match(/^untitled/i)) {
            outFile = File(outFile.parent.fsName + "/" + suggested);
        }
    } catch (ignore2) {}

    var ok = false, errMsg = "";
    try {
        tempDoc.exportFile(ExportFormat.INDESIGN_MARKUP, outFile);
        ok = true;
    } catch (err) { errMsg = err && err.message ? err.message : String(err); }
    finally { if (tempDoc && tempDoc.isValid) try { tempDoc.close(SaveOptions.NO); } catch (ignore3) {} }

    if (ok) alert("Export complete.\n\n" + outFile.fsName);
    else alert("Export failed.\n\n" + errMsg);
})();
