﻿(function () {
    // ---- Guards ----
    if (app.documents.length === 0) { alert("Open a document and select objects."); return; }
    var srcDoc = app.activeDocument;
    if (!app.selection || app.selection.length === 0) { alert("Select one or more objects and run again."); return; }

    var items = getSelectablePageItems(app.selection);
    if (items.length === 0) { alert("Selection has no movable page items."); return; }

    // ---- Snapshot + normalise for measurement ----
    var originalRulerOrigin = srcDoc.viewPreferences.rulerOrigin;
    var originalZeroPoint   = srcDoc.zeroPoint.slice(0);

    srcDoc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
    srcDoc.zeroPoint = [0, 0];

    // ---- UI (mm only) ----
    var lastKey = "ExportSelToNewDoc_mm_withStrokes";
    var last = readAppLabelJSON(app, lastKey) || { paddingMm: 0, groupOnPaste: false, includeStrokes: true };

    var dlg = new Window("dialog", "Export Selection — mm (no bleed)");
    dlg.alignChildren = "fill";

    var g1 = dlg.add("group");
    g1.add("statictext", undefined, "Padding (mm):");
    var padEdit = g1.add("edittext", undefined, String(last.paddingMm || 0));
    padEdit.characters = 8;

    var g2 = dlg.add("group");
    var chkGroup = g2.add("checkbox", undefined, "Group items on paste");
    chkGroup.value = !!last.groupOnPaste;

    var g3 = dlg.add("group");
    var chkStrokes = g3.add("checkbox", undefined, "Include strokes in measurement (visible bounds)");
    chkStrokes.value = last.includeStrokes !== false; // default true

    var gBtns = dlg.add("group"); gBtns.alignment = "right";
    var okBtn = gBtns.add("button", undefined, "OK");
    gBtns.add("button", undefined, "Cancel", { name: "cancel" });

    okBtn.onClick = function () {
        var s = (padEdit.text || "").trim(); if (s === "") s = "0";
        if (!/^[-+]?\d+(\.\d+)?$/.test(s)) { alert("Padding must be numeric (mm)."); return; }
        dlg.close(1);
    };
    if (dlg.show() !== 1) { restoreRulers(); return; }

    var paddingMm = parseFloat(padEdit.text || "0");
    var groupOnPaste = !!chkGroup.value;
    var includeStrokes = !!chkStrokes.value;
    writeAppLabelJSON(app, lastKey, { paddingMm: paddingMm, groupOnPaste: groupOnPaste, includeStrokes: includeStrokes });

    if (paddingMm < 0) paddingMm = 0;

    // ---- Measure selection (using chosen bounds mode) ----
    var unionPt = getUnionBounds(items, includeStrokes); // [t,l,b,r] in points
    var selWpt = unionPt[3] - unionPt[1];
    var selHpt = unionPt[2] - unionPt[0];
    if (selWpt <= 0 || selHpt <= 0) {
        restoreRulers(); alert("Measured selection size is zero. Please check your selection."); return;
    }

    // Convert to mm for maths & reporting
    var selWmm = ptToMm(selWpt);
    var selHmm = ptToMm(selHpt);

    // Target page size (mm)
    var pageWmm = selWmm + (2 * paddingMm);
    var pageHmm = selHmm + (2 * paddingMm);

    // Copy selection before we leave the source doc
    try { app.select(NothingEnum.NOTHING); app.select(items); app.copy(); }
    catch (e) { restoreRulers(); alert("Copy failed: " + e); return; }

    // Restore source ruler state
    restoreRulers();

    // ---- Create new doc in mm, no bleed/margins, paste+position ----
    app.doScript(function () {
        var newDoc = app.documents.add();
        try { newDoc.documentColorSpace = srcDoc.documentColorSpace; } catch (_e) {}
        try { newDoc.documentPreferences.intent = srcDoc.documentPreferences.intent; } catch (_e2) {}

        // Page size (mm -> pt for API)
        var dp = newDoc.documentPreferences;
        dp.facingPages = false;
        dp.pagesPerDocument = 1;
        dp.pageWidth  = mmToPt(pageWmm);
        dp.pageHeight = mmToPt(pageHmm);

        // No margins (doc + first page)
        try {
            newDoc.marginPreferences.top = 0;
            newDoc.marginPreferences.bottom = 0;
            newDoc.marginPreferences.left = 0;
            newDoc.marginPreferences.right = 0;
        } catch (_m) {}
        try {
            var mp0 = newDoc.pages[0].marginPreferences;
            mp0.top = mp0.bottom = mp0.left = mp0.right = 0;
        } catch (_m0) {}

        // No bleed
        try {
            newDoc.documentPreferences.documentBleedTopOffset =
            newDoc.documentPreferences.documentBleedBottomOffset =
            newDoc.documentPreferences.documentBleedInsideOrLeftOffset =
            newDoc.documentPreferences.documentBleedOutsideOrRightOffset = 0;
        } catch (_b) {}

        // Make the new doc display in millimetres
        newDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
        newDoc.viewPreferences.verticalMeasurementUnits   = MeasurementUnits.MILLIMETERS;
        newDoc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
        newDoc.zeroPoint = [0, 0];

        // Paste
        app.paste();
        var pasted = app.selection;
        if (!pasted || pasted.length === 0) throw new Error("Paste failed.");

        // Optionally group
        var pastedGroup = null;
        if (groupOnPaste && pasted.length > 1) {
            try { pastedGroup = newDoc.pages[0].groups.add(pasted); } catch (_ge) { pastedGroup = null; }
        }

        // Shift so top-left = (paddingMm, paddingMm); measure pasted with SAME mode
        var paddingPt = mmToPt(paddingMm);
        var tgt = pastedGroup ? getItemBounds(pastedGroup, includeStrokes) : getUnionBounds(pasted, includeStrokes); // points
        var shiftX = paddingPt - tgt[1];
        var shiftY = paddingPt - tgt[0];

        if (pastedGroup) pastedGroup.move(undefined, [shiftX, shiftY]);
        else {
            for (var i = 0; i < pasted.length; i++) {
                try { pasted[i].move(undefined, [shiftX, shiftY]); } catch (_m) {}
            }
        }

        // Report (mm)
        alert(
            "Done.\n\n" +
            "Padding: " + nice3(paddingMm) + " mm\n" +
            "Page size: " + nice3(pageWmm) + " mm × " + nice3(pageHmm) + " mm\n" +
            "Measured with: " + (includeStrokes ? "Visible bounds (includes strokes/effects)" : "Geometric bounds (ignores stroke)") + "\n" +
            "Content placed at (" + nice3(paddingMm) + ", " + nice3(paddingMm) + ") mm from top-left.\n" +
            "Margins: 0 mm, Bleed: 0 mm."
        );
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Export Selection To New Document — mm");

    // ---- Helpers ----
    function restoreRulers() {
        try { srcDoc.viewPreferences.rulerOrigin = originalRulerOrigin; } catch (_e) {}
        try { srcDoc.zeroPoint = originalZeroPoint; } catch (_e2) {}
    }

    function getSelectablePageItems(selArray) {
        var out = [];
        for (var i = 0; i < selArray.length; i++) {
            var it = selArray[i];
            if (it && it.hasOwnProperty("geometricBounds") && it.hasOwnProperty("move")) {
                try {
                    if (it.locked || (it.itemLayer && it.itemLayer.locked)) continue;
                    if (isGuide(it)) continue;
                    out.push(it);
                } catch (_e) {}
            }
        }
        return out;
    }

    function isGuide(item) {
        try { return item.constructor && (item.constructor.name === "Guide"); }
        catch (_e) { return false; }
    }

    // Get bounds of one item in points using chosen mode
    function getItemBounds(item, useVisible) {
        try {
            if (useVisible && item.hasOwnProperty("visibleBounds")) return item.visibleBounds;
            return item.geometricBounds;
        } catch (_e) {
            return null;
        }
    }

    // union bounds [t,l,b,r] in points; useVisible => visibleBounds, else geometricBounds
    function getUnionBounds(items, useVisible) {
        var t =  1e12, l =  1e12, b = -1e12, r = -1e12;
        for (var i = 0; i < items.length; i++) {
            var gb = getItemBounds(items[i], useVisible);
            if (!gb || gb.length !== 4) continue;
            t = Math.min(t, gb[0]); l = Math.min(l, gb[1]);
            b = Math.max(b, gb[2]); r = Math.max(r, gb[3]);
        }
        return [t, l, b, r];
    }

    // ---- Conversions (mm <-> pt) ----
    function mmToPt(mm) { return Number(mm) * 2.8346456693; }
    function ptToMm(pt) { return Number(pt) / 2.8346456693; }

    // display rounding helpers
    function nice3(n) {
        var v = Number(n);
        if (Math.abs(v) >= 100) return v.toFixed(0);
        return v.toFixed(3);
    }

    // labels
    function readAppLabelJSON(appRef, key) {
        try { var raw = appRef.extractLabel(key); return raw ? JSON.parse(raw) : null; }
        catch (_e) { return null; }
    }
    function writeAppLabelJSON(appRef, key, obj) {
        try { appRef.insertLabel(key, JSON.stringify(obj)); } catch (_e) {}
    }
})();
