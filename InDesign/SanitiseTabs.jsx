/**
 * SanitiseTabs.jsx — v1.0
 * Find and remove DUPLICATE tab stops in supplied InDesign artwork, in both
 * paragraph style definitions and local text — a sanitisation pass before
 * uploading to an InDesign-server based system.
 * InDesign ExtendScript (2023+ Mac/Win)
 *
 * WHAT IT DOES
 *   1. SCAN (no changes): clusters each tab list's stops by position within
 *      a tolerance (default 0.05 mm — catches both real duplicates and the
 *      float-dust kind left by scaling). Within a cluster:
 *        - same alignment + leader (+ align-on char)  -> DUPLICATES
 *        - different types at the same position       -> CONFLICTS
 *      Styles are analysed first (all groups, incl. [Basic Paragraph]);
 *      then every paragraph in every story — table cells, footnotes,
 *      type-on-path and parent pages included. A paragraph whose tab list
 *      simply inherits its style is attributed to the STYLE, not counted
 *      again; only genuine local overrides land in the "unstyled/local
 *      text" bucket.
 *   2. REPORT: counts per style + unstyled/local total, with page numbers
 *      and text snippets for local findings. Conflicts listed separately.
 *   3. REMOVE ALL DUPLICATES (optional, one undo step): rewrites cleaned
 *      tab lists — first stop of each duplicate group survives at its exact
 *      original position. Conflicts are left alone unless "treat conflicts
 *      as duplicates (keep first stop)" is ticked. Locked layers/items are
 *      unlocked and restored. A verification rescan runs afterwards and is
 *      appended to the final report — it should read zero.
 *
 * KNOWN LIMIT
 *   InDesign has no per-attribute override clear, so a local fix leaves a
 *   value-identical tab override on the paragraph (the override "+" may
 *   remain even though the tabs now match the style). Composition is
 *   unaffected.
 */

#target "InDesign"

(function () {
    if (!app.documents.length) { alert("Open a document first."); return; }

    var MM2PT = 72 / 25.4;
    var doc = app.activeDocument;

    // ---------------- options ----------------

    var cfg = optionsDialog();
    if (!cfg) return;
    var TOL = cfg.tolMM * MM2PT;

    var oldUnit = app.scriptPreferences.measurementUnit;
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

    var scan;
    try {
        scan = doScan();
    } finally {
        app.scriptPreferences.measurementUnit = oldUnit;
    }

    var choice = reportDialog(buildReport(scan, false), scan.totalDup + (cfg.conflictsToo ? scan.totalConflictStops : 0) > 0);
    if (choice !== 1) return;

    // ---------------- remove + verify ----------------

    var lockedLayers = [], lockedItems = [];
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
    try {
        unlockAll(lockedLayers, lockedItems);
        app.doScript(function () { applyFixes(scan); },
            ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Sanitise duplicate tabs");
        var verify = doScan();
        showReport("Sanitise Tabs \u2014 done",
            buildReport(scan, true) +
            "\n\nVERIFICATION RESCAN:\n" +
            "  duplicate tab stops remaining: " + verify.totalDup +
            "\n  conflicts remaining: " + verify.totalConflictStops +
            (verify.totalDup === 0 ? "\n\nClean. One undo step reverts everything." :
                "\n\nSomething survived \u2014 send me this report."));
    } finally {
        restoreLocks(lockedLayers, lockedItems);
        app.scriptPreferences.measurementUnit = oldUnit;
    }
    return;

    // =====================================================================
    // Tab list analysis
    // =====================================================================

    function sigOf(stop) {
        var al = "", ld = "", ch = "";
        try { al = String(stop.alignment); } catch (_) {}
        try { ld = String(stop.leader || ""); } catch (_) {}
        try {
            if (String(stop.alignment) === String(TabStopAlignment.CHARACTER_ALIGN)) {
                ch = String(stop.alignmentCharacter || "");
            }
        } catch (_) {}
        return al + "|" + ld + "|" + ch;
    }

    function plainStop(stop) {
        var o = { position: Number(stop.position) };
        try { o.alignment = stop.alignment; } catch (_) {}
        try { o.leader = stop.leader; } catch (_) {}
        try { o.alignmentCharacter = stop.alignmentCharacter; } catch (_) {}
        return o;
    }

    // Analyse one tab list.
    // Returns { dup: n removable duplicates, conflictStops: n stops in
    //           conflicted clusters beyond the first, cleaned: [stops],
    //           cleanedIfConflicts: [stops], changed: bool }
    function analyse(tabs) {
        var stops = [];
        for (var i = 0; i < tabs.length; i++) {
            try { stops.push(plainStop(tabs[i])); } catch (_) {}
        }
        stops.sort(function (a, b) { return a.position - b.position; });

        var clusters = [], cur = null;
        for (var s = 0; s < stops.length; s++) {
            if (cur === null || (stops[s].position - cur.anchor) > TOL) {
                cur = { anchor: stops[s].position, items: [] };
                clusters.push(cur);
            }
            cur.items.push(stops[s]);
        }

        var dup = 0, conflictStops = 0;
        var cleaned = [], cleanedC = [];
        for (var c = 0; c < clusters.length; c++) {
            var items = clusters[c].items;
            if (items.length === 1) { cleaned.push(items[0]); cleanedC.push(items[0]); continue; }
            var seen = {}, kept = [], sigCount = 0;
            for (var k = 0; k < items.length; k++) {
                var g = sigOf(items[k]);
                if (!seen[g]) { seen[g] = true; kept.push(items[k]); sigCount++; }
                else dup++;
            }
            if (sigCount > 1) conflictStops += (sigCount - 1);
            for (var m = 0; m < kept.length; m++) cleaned.push(kept[m]);
            cleanedC.push(kept[0]); // conflicts-as-duplicates: first stop only
        }
        return {
            dup: dup,
            conflictStops: conflictStops,
            cleaned: cleaned,
            cleanedIfConflicts: cleanedC,
            changed: dup > 0,
            changedIfConflicts: (dup + conflictStops) > 0
        };
    }

    function sameList(a, b) {
        if (a.length !== b.length) return false;
        for (var i = 0; i < a.length; i++) {
            if (Math.abs(Number(a[i].position) - Number(b[i].position)) > 1e-4) return false;
            if (sigOf(a[i]) !== sigOf(b[i])) return false;
        }
        return true;
    }

    // =====================================================================
    // Scan
    // =====================================================================

    function doScan() {
        var res = {
            styles: [],          // {style, name, dup, conflictStops, ana}
            styleById: {},
            local: [],           // {para, page, snippet, dup, conflictStops, ana}
            totalDup: 0,
            totalConflictStops: 0,
            localDup: 0,
            localConflicts: 0,
            paraChecked: 0
        };

        // ---- styles ----
        var all = doc.allParagraphStyles;
        for (var i = 0; i < all.length; i++) {
            if (i === 0) continue; // [No Paragraph Style]
            var st = all[i];
            var tl = null;
            try { tl = st.tabList; } catch (_) { continue; }
            if (!tl || tl.length < 2) { res.styleById[String(st.id)] = null; continue; }
            var ana = analyse(tl);
            res.styleById[String(st.id)] = ana;
            if (ana.dup || ana.conflictStops) {
                res.styles.push({ style: st, name: stylePath(st), dup: ana.dup, conflictStops: ana.conflictStops, ana: ana });
                res.totalDup += ana.dup;
                res.totalConflictStops += ana.conflictStops;
            }
        }

        // ---- local text: every story, tables, footnotes ----
        var stories = doc.stories;
        for (var sI = 0; sI < stories.length; sI++) {
            var story = stories[sI];
            scanParas(story, res);
            try {
                var tabs = story.tables;
                for (var t = 0; t < tabs.length; t++) {
                    var cells = tabs[t].cells;
                    for (var c = 0; c < cells.length; c++) {
                        try { scanParas(cells[c].texts[0], res); } catch (_) {}
                    }
                }
            } catch (_) {}
            try {
                var fns = story.footnotes;
                for (var f = 0; f < fns.length; f++) {
                    try { scanParas(fns[f].texts[0], res); } catch (_) {}
                }
            } catch (_) {}
        }
        return res;
    }

    function scanParas(txt, res) {
        var paras;
        try { paras = txt.paragraphs; } catch (_) { return; }
        for (var p = 0; p < paras.length; p++) {
            var para = paras[p];
            res.paraChecked++;
            var tl = null;
            try { tl = para.tabList; } catch (_) { continue; }
            if (!tl || tl.length < 2) continue;

            // inherited-only? then it's the style's problem, already counted
            var styleAna = null, styleTabs = null;
            try {
                var st = para.appliedParagraphStyle;
                styleTabs = st.tabList;
            } catch (_) {}
            var effective = [];
            for (var i = 0; i < tl.length; i++) { try { effective.push(plainStop(tl[i])); } catch (_) {} }
            effective.sort(function (a, b) { return a.position - b.position; });
            if (styleTabs) {
                var styleList = [];
                for (var j = 0; j < styleTabs.length; j++) { try { styleList.push(plainStop(styleTabs[j])); } catch (_) {} }
                styleList.sort(function (a, b) { return a.position - b.position; });
                if (sameList(effective, styleList)) continue; // pure inheritance
            }

            var ana = analyse(tl);
            if (!ana.dup && !ana.conflictStops) continue;

            var page = "?", snippet = "";
            try {
                var fr = para.insertionPoints[0].parentTextFrames;
                if (fr && fr.length && fr[0].parentPage && fr[0].parentPage.isValid) page = fr[0].parentPage.name;
            } catch (_) {}
            try {
                snippet = String(para.contents).replace(/[\r\n\u2028\u2029\t]+/g, " ").replace(/^\s+|\s+$/g, "");
                if (snippet.length > 34) snippet = snippet.substring(0, 34) + "\u2026";
            } catch (_) {}

            res.local.push({ para: para, page: page, snippet: snippet, dup: ana.dup, conflictStops: ana.conflictStops, ana: ana });
            res.localDup += ana.dup;
            res.localConflicts += ana.conflictStops;
            res.totalDup += ana.dup;
            res.totalConflictStops += ana.conflictStops;
        }
    }

    function stylePath(style) {
        var parts = [style.name];
        try {
            var p = style.parent, guard = 0;
            while (p && p.isValid && (p.constructor ? p.constructor.name : "") === "ParagraphStyleGroup" && guard++ < 16) {
                parts.unshift(p.name);
                p = p.parent;
            }
        } catch (_) {}
        return parts.join("/");
    }

    // =====================================================================
    // Fixes
    // =====================================================================

    function applyFixes(scanRes) {
        var i;
        for (i = 0; i < scanRes.styles.length; i++) {
            var se = scanRes.styles[i];
            var list = cfg.conflictsToo ? se.ana.cleanedIfConflicts : se.ana.cleaned;
            var changed = cfg.conflictsToo ? se.ana.changedIfConflicts : se.ana.changed;
            if (!changed) continue;
            try { se.style.tabList = list; } catch (_) {}
        }
        for (i = 0; i < scanRes.local.length; i++) {
            var le = scanRes.local[i];
            try {
                if (!le.para.isValid) continue;
                // re-analyse fresh: the style fix above may have changed the
                // effective list under this paragraph
                var ana2 = analyse(le.para.tabList);
                var list2 = cfg.conflictsToo ? ana2.cleanedIfConflicts : ana2.cleaned;
                var changed2 = cfg.conflictsToo ? ana2.changedIfConflicts : ana2.changed;
                if (changed2) le.para.tabList = list2;
            } catch (_) {}
        }
    }

    function unlockAll(lockedLayers, lockedItems) {
        var i;
        try {
            for (i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].locked) { lockedLayers.push(doc.layers[i]); doc.layers[i].locked = false; }
            }
        } catch (_) {}
        try {
            var all = doc.allPageItems;
            for (i = 0; i < all.length; i++) {
                try { if (all[i].locked) { lockedItems.push(all[i]); all[i].locked = false; } } catch (_) {}
            }
        } catch (_) {}
    }

    function restoreLocks(lockedLayers, lockedItems) {
        var i;
        for (i = 0; i < lockedItems.length; i++) { try { lockedItems[i].locked = true; } catch (_) {} }
        for (i = 0; i < lockedLayers.length; i++) { try { lockedLayers[i].locked = true; } catch (_) {} }
    }

    // =====================================================================
    // Report text
    // =====================================================================

    function buildReport(s, afterFix) {
        var L = [];
        L.push((afterFix ? "REMOVED" : "FOUND") + " \u2014 duplicate tab stops (tolerance " + cfg.tolMM + " mm):");
        L.push("  Unstyled / local text: " + s.localDup + " duplicate tab" + (s.localDup === 1 ? "" : "s") +
               (s.local.length ? "  (in " + s.local.length + " paragraph" + (s.local.length === 1 ? "" : "s") + ")" : ""));
        if (!s.styles.length) L.push("  (no styles affected)");
        for (var i = 0; i < s.styles.length; i++) {
            L.push("  " + s.styles[i].name + ": " + s.styles[i].dup + " duplicate tab" + (s.styles[i].dup === 1 ? "" : "s") +
                   (s.styles[i].conflictStops ? "   [+" + s.styles[i].conflictStops + " conflict]" : ""));
        }
        L.push("");
        L.push("Total duplicates: " + s.totalDup + "    Conflicts (same position, different type): " + s.totalConflictStops +
               (s.totalConflictStops ? (cfg.conflictsToo ? "  \u2014 WILL be removed (keep first)" : "  \u2014 report only") : ""));
        L.push("Paragraphs checked: " + s.paraChecked);

        if (s.local.length) {
            L.push("");
            L.push("LOCAL FINDINGS:");
            for (var j = 0; j < s.local.length; j++) {
                var e = s.local[j];
                L.push("  p." + e.page + "  \"" + e.snippet + "\"  \u2014 " + e.dup + " dup" +
                       (e.conflictStops ? ", " + e.conflictStops + " conflict" : ""));
            }
        }
        return L.join("\n");
    }

    // =====================================================================
    // UI
    // =====================================================================

    function optionsDialog() {
        var w = new Window("dialog", "Sanitise Tabs \u2014 v1.0");
        w.orientation = "column"; w.alignChildren = "left";
        var g1 = w.add("group");
        g1.add("statictext", undefined, "Treat tab stops within");
        var et = g1.add("edittext", [0, 0, 60, 24], "0.05");
        g1.add("statictext", undefined, "mm of each other as the same position.");
        var cbConf = w.add("checkbox", undefined, "Also remove CONFLICTS (same position, different tab type) \u2014 keeps the first stop");
        cbConf.value = false;
        w.add("statictext", undefined, "Scope: entire document \u2014 styles, all stories, tables, footnotes, type-on-path, parent pages.");
        var gB = w.add("group"); gB.alignment = "right";
        gB.add("button", undefined, "Cancel", { name: "cancel" });
        gB.add("button", undefined, "Scan", { name: "ok" });
        if (w.show() !== 1) return null;
        var tol = parseFloat(et.text);
        if (isNaN(tol) || tol < 0) { alert("Tolerance must be zero or a positive number of millimetres."); return null; }
        return { tolMM: tol, conflictsToo: cbConf.value };
    }

    function reportDialog(text, anythingToRemove) {
        var w = new Window("dialog", "Sanitise Tabs \u2014 scan results (no changes made yet)");
        w.orientation = "column"; w.alignChildren = "fill";
        w.add("edittext", [0, 0, 700, 380], text, { multiline: true, readonly: true, scrolling: true });
        w.add("statictext", undefined, "Select text + Ctrl/Cmd+C to copy the report.");
        var g = w.add("group"); g.alignment = "right";
        g.add("button", undefined, "Close", { name: "cancel" });
        var go = g.add("button", undefined, "Remove All Duplicates", { name: "ok" });
        go.enabled = anythingToRemove === true;
        return w.show(); // 1 = remove
    }

    function showReport(title, text) {
        var w = new Window("dialog", title);
        w.orientation = "column"; w.alignChildren = "fill";
        w.add("edittext", [0, 0, 700, 400], text, { multiline: true, readonly: true, scrolling: true });
        var g = w.add("group"); g.alignment = "right";
        g.add("button", undefined, "Close", { name: "ok" });
        w.show();
    }

})();
