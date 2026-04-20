#target indesign

(function () {

    if (app.documents.length === 0) { alert("No document open."); return; }
    var doc = app.activeDocument;

    // ─────────────────────────────
    // Regex: matches filenames like pic1.pdf, pic12.pdf (case-insensitive)
    // ─────────────────────────────
    var pattern = /^pic(\d+)\.pdf$/i;

    var renamed  = [];
    var skipped  = [];

    // ─────────────────────────────
    // Walk every page, every item
    // ─────────────────────────────
    for (var p = 0; p < doc.pages.length; p++) {
        var page  = doc.pages[p];
        var items = page.allPageItems;

        for (var i = 0; i < items.length; i++) {
            var item = items[i];

            // Only interested in frames that have a placed graphic
            if (!item.allGraphics || item.allGraphics.length === 0) continue;

            try {
                var link = item.allGraphics[0].itemLink;
                if (!link) continue;

                // Extract just the filename from the full path
                var fullPath = link.filePath || link.name || "";
                var parts    = fullPath.replace(/\\/g, "/").split("/");
                var fileName = parts[parts.length - 1];  // e.g. "pic3.pdf"

                var match = fileName.match(pattern);
                if (!match) continue;

                // Build the new name: Pic1, Pic2, etc.
                var newName = "Pic" + match[1];
                var oldName = item.name || "(unnamed)";

                // Skip if it's already correctly named
                if (oldName === newName) {
                    skipped.push({ page: p + 1, name: newName, reason: "already correct" });
                    continue;
                }

                item.name = newName;
                renamed.push({ page: p + 1, from: oldName, to: newName, file: fileName });

            } catch (e) {
                skipped.push({ page: p + 1, name: "(error)", reason: e.message });
            }
        }
    }

    // ─────────────────────────────
    // Summary report
    // ─────────────────────────────
    var msg = "── Rename Complete ──────────────────\n\n";

    if (renamed.length === 0) {
        msg += "No matching elements found.\n";
    } else {
        msg += "✅ Renamed: " + renamed.length + " element(s)\n\n";
        for (var r = 0; r < renamed.length; r++) {
            var entry = renamed[r];
            msg += "  Pg " + entry.page + "  [" + entry.from + "]  →  [" + entry.to + "]"
                 + "  (" + entry.file + ")\n";
        }
    }

    if (skipped.length > 0) {
        msg += "\n⚠ Skipped: " + skipped.length + " element(s)\n";
        for (var s = 0; s < skipped.length; s++) {
            var sk = skipped[s];
            msg += "  Pg " + sk.page + "  [" + sk.name + "]  — " + sk.reason + "\n";
        }
    }

    alert(msg);

})();