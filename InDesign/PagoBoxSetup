#target indesign

// ══════════════════════════════════════════════════════
//  Module Frame Renamer
//  Renames frames by size to: Module001_FPFP_001 format
//  Runs on selection, or falls back to active page
// ══════════════════════════════════════════════════════

(function () {

    if (app.documents.length === 0) { alert("No document open."); return; }
    var doc = app.activeDocument;

    // ─────────────────────────────
    // Helpers
    // ─────────────────────────────
    function pad(n, len) {
        var s = String(Math.round(n));
        while (s.length < len) s = '0' + s;
        return s;
    }

    var docUnit    = doc.viewPreferences.horizontalMeasurementUnits;
    var unitStrMap = {};
    unitStrMap[MeasurementUnits.MILLIMETERS]    = 'mm';
    unitStrMap[MeasurementUnits.POINTS]         = 'pt';
    unitStrMap[MeasurementUnits.INCHES]         = 'in';
    unitStrMap[MeasurementUnits.INCHES_DECIMAL] = 'in';
    unitStrMap[MeasurementUnits.CENTIMETERS]    = 'cm';
    unitStrMap[MeasurementUnits.PICAS]          = 'pc';
    unitStrMap[MeasurementUnits.CICEROS]        = 'ci';
    var docUnitStr = unitStrMap[docUnit] || 'mm';

    function toMM(val) {
        if (docUnitStr === 'mm') return val;
        return new UnitValue(val, docUnitStr).as('mm');
    }

    // ─────────────────────────────
    // Config Dialog
    // ─────────────────────────────
    var dlg = new Window("dialog", "Module Frame Renamer");
    dlg.alignChildren = 'left';
    dlg.spacing = 10;
    dlg.margins = 16;

    // Page code row
    var pageRow = dlg.add("group");
    pageRow.add("statictext", undefined, "Page size code:");
    var pageCodeInput = pageRow.add("edittext", undefined, "");
    pageCodeInput.preferredSize.width = 60;

    // Tolerance row
    var tolRow = dlg.add("group");
    tolRow.add("statictext", undefined, "Match tolerance (mm):");
    var tolInput = tolRow.add("edittext", undefined, "0.5");
    tolInput.preferredSize.width = 40;

    // Script label option
    var labelRow = dlg.add("group");
    var applyLabelChk = labelRow.add("checkbox", undefined, "Also apply name to Script Label");
    applyLabelChk.value = true;

    // Module size definitions panel
    var panel = dlg.add("panel", undefined, "Module Sizes  (Code | Width mm | Height mm)");
    panel.alignChildren = 'left';
    panel.spacing = 5;
    panel.margins = 12;

    // Column headers
    var hdr = panel.add("group");
    hdr.spacing = 6;
    var hChk  = hdr.add("statictext", undefined, "✓");
    hChk.preferredSize.width = 20;
    var hCode = hdr.add("statictext", undefined, "Code");
    hCode.preferredSize.width = 55;
    var hW    = hdr.add("statictext", undefined, "Width (mm)");
    hW.preferredSize.width = 85;
    var hH    = hdr.add("statictext", undefined, "Height (mm)");
    hH.preferredSize.width = 85;

    // 8 blank rows — no defaults
    var rows = [];
    for (var r = 0; r < 8; r++) {
        var row  = panel.add("group");
        row.spacing = 6;
        var chk  = row.add("checkbox", undefined, "");
        chk.value = false;
        chk.preferredSize.width = 20;
        var code = row.add("edittext", undefined, "");
        code.preferredSize.width = 55;
        var wIn  = row.add("edittext", undefined, "");
        wIn.preferredSize.width  = 85;
        var hIn  = row.add("edittext", undefined, "");
        hIn.preferredSize.width  = 85;
        rows.push({ chk:chk, code:code, w:wIn, h:hIn });
    }

    // Buttons
    var btnGroup = dlg.add("group");
    btnGroup.alignment = 'right';
    btnGroup.add("button", undefined, "Cancel", { name:"cancel" });
    btnGroup.add("button", undefined, "Run",    { name:"ok" });

    if (dlg.show() !== 1) return;

    // ─────────────────────────────
    // Parse user config
    // ─────────────────────────────
    var pageCode     = pageCodeInput.text.replace(/\s/g, "").toUpperCase();
    var tolerance    = parseFloat(tolInput.text) || 0.5;
    var applyLabel   = applyLabelChk.value;

    if (!pageCode) { alert("Page size code cannot be empty."); return; }

    var moduleDefs = [];
    for (var r = 0; r < rows.length; r++) {
        var row  = rows[r];
        if (!row.chk.value) continue;
        var code = row.code.text.replace(/\s/g, "").toUpperCase();
        var w    = parseFloat(row.w.text);
        var h    = parseFloat(row.h.text);
        if (!code || isNaN(w) || isNaN(h)) continue;
        moduleDefs.push({ code:code, w:w, h:h });
    }

    if (moduleDefs.length === 0) { alert("No valid module sizes defined."); return; }

    // ─────────────────────────────
    // Collect items
    // Selection takes priority — falls back to active page
    // ─────────────────────────────
    var items  = [];
    var source = "";

    if (app.selection.length > 0) {
        source = "selection (" + app.selection.length + " item(s))";
        for (var s = 0; s < app.selection.length; s++) {
            var sel = app.selection[s];
            if (sel instanceof Group) {
                var gi = sel.allPageItems;
                for (var g = 0; g < gi.length; g++) items.push(gi[g]);
            } else {
                items.push(sel);
            }
        }
    } else {
        var page = app.activeWindow.activePage;
        source   = "active page (" + (page.name || String(page.documentOffset + 1)) + ")";
        var pi   = page.allPageItems;
        for (var i = 0; i < pi.length; i++) items.push(pi[i]);
    }

    if (items.length === 0) { alert("No items found in " + source + "."); return; }

    // ─────────────────────────────
    // Sort: top → bottom, left → right
    // ─────────────────────────────
    items.sort(function (a, b) {
        var ay = 0, by = 0, ax = 0, bx = 0;
        try { ay = toMM(a.geometricBounds[0]); } catch(e) {}
        try { by = toMM(b.geometricBounds[0]); } catch(e) {}
        try { ax = toMM(a.geometricBounds[1]); } catch(e) {}
        try { bx = toMM(b.geometricBounds[1]); } catch(e) {}
        if (Math.abs(ay - by) > tolerance) return ay - by;
        return ax - bx;
    });

    // ─────────────────────────────
    // Match sizes and rename
    // ─────────────────────────────
    var moduleCounter = 1;
    var sizeCounters  = {};
    var renamed       = [];
    var skipped       = [];

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var bounds;

        try { bounds = item.geometricBounds; }
        catch (e) {
            skipped.push("item " + (i + 1) + ": cannot read bounds");
            continue;
        }

        var itemW = toMM(bounds[3] - bounds[1]);
        var itemH = toMM(bounds[2] - bounds[0]);

        // Find matching size definition
        var matched = null;
        for (var m = 0; m < moduleDefs.length; m++) {
            var def = moduleDefs[m];
            if (Math.abs(itemW - def.w) <= tolerance &&
                Math.abs(itemH - def.h) <= tolerance) {
                matched = def;
                break;
            }
        }

        if (!matched) {
            skipped.push('"' + (item.name || "unnamed") + '"  '
                + itemW.toFixed(1) + ' × ' + itemH.toFixed(1) + ' mm  — no size match');
            continue;
        }

        // Per-size counter
        if (!sizeCounters[matched.code]) sizeCounters[matched.code] = 1;
        var sizeNum = sizeCounters[matched.code]++;

        var newName = "Module" + pad(moduleCounter, 3)
                    + "_" + pageCode + matched.code
                    + "_" + pad(sizeNum, 3);

        var oldName = item.name || "(unnamed)";

        // Apply element name
        item.name = newName;

        // Apply script label if option is ticked
        if (applyLabel) {
            try { item.label = newName; } catch(e) {}
        }

        moduleCounter++;

        renamed.push({
            from:       oldName,
            to:         newName,
            size:       itemW.toFixed(1) + " × " + itemH.toFixed(1) + " mm",
            labelSet:   applyLabel
        });
    }

    // ─────────────────────────────
    // Summary report
    // ─────────────────────────────
    var msg = "── Module Rename Complete ──────────────────\n";
    msg    += "Source: " + source + "\n";
    msg    += "Script label: " + (applyLabel ? "✅ applied" : "— skipped") + "\n\n";

    if (renamed.length === 0) {
        msg += "No matching frames found.\n";
    } else {
        msg += "✅  Renamed: " + renamed.length + " element(s)\n\n";
        for (var r = 0; r < renamed.length; r++) {
            var entry = renamed[r];
            msg += "  " + entry.to
                 + "   (" + entry.size + ")"
                 + (entry.from !== "(unnamed)" && entry.from.indexOf("Module") === -1
                     ? "   was: " + entry.from : "")
                 + "\n";
        }
    }

    if (skipped.length > 0) {
        msg += "\n⚠  Skipped: " + skipped.length + " element(s)\n";
        for (var s = 0; s < skipped.length; s++) {
            msg += "  " + skipped[s] + "\n";
        }
    }

    msg += "\n── Size breakdown ──\n";
    for (var code in sizeCounters) {
        if (Object.prototype.hasOwnProperty.call(sizeCounters, code)) {
            msg += "  " + pageCode + code + ":  " + (sizeCounters[code] - 1) + " frame(s)\n";
        }
    }

    alert(msg);

})();
