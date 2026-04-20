// ============================================================
// InDesign - DEBUG: List all page item names
// Run this first to see the EXACT names InDesign uses
// so the rename script can be matched correctly.
// ============================================================

#target indesign

(function () {

    if (app.documents.length === 0) {
        alert("No document open.");
        return;
    }

    var doc   = app.activeDocument;
    var found = [];

    // Recursive function to walk into groups and nested items
    function collectItems(itemCollection) {
        for (var i = 0; i < itemCollection.length; i++) {
            var item = itemCollection[i];
            found.push("\"" + item.name + "\"  [" + item.constructor.name + "]");
            // Walk into groups
            if (item.hasOwnProperty("pageItems") || item.groups && item.groups.length > 0) {
                try { collectItems(item.pageItems); } catch(e){}
            }
        }
    }

    // Walk every spread > page > pageItems
    for (var s = 0; s < doc.spreads.length; s++) {
        var spread = doc.spreads[s];
        for (var p = 0; p < spread.pages.length; p++) {
            collectItems(spread.pages[p].pageItems);
        }
        // Also check spread-level items (items spanning both pages)
        collectItems(spread.pageItems);
    }

    if (found.length === 0) {
        alert("No named page items found in this document.\n\nMake sure the document has content on its pages.");
        return;
    }

    var report = [];
    report.push("FOUND " + found.length + " PAGE ITEMS:\n");
    for (var i = 0; i < found.length; i++) {
        report.push("  " + (i + 1) + ".  " + found[i]);
    }

    alert(report.join("\n"));

})();
