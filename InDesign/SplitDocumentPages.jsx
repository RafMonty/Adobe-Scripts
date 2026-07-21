/**
 * SplitDocumentPages.jsx — v1.2
 * Export every page of the active document as its own INDD or IDML file.
 * InDesign ExtendScript (2023+ Mac/Win)
 *
 * v1.2 CHANGES — purge scope choice
 *   - "Keep ONLY this page's styles" (default, upload mode): the keep-set is
 *     exactly the audit-frame list + the two built-ins. Styles used only by
 *     parent-page or pasteboard text ARE deleted; InDesign preserves that
 *     text's appearance as local overrides (delete-and-preserve), so parents
 *     look the same but the style panel is page-pure.
 *   - "Also keep parent/pasteboard usage + based-on ancestors": the v1.1
 *     conservative behaviour, for non-upload uses of the splitter.
 *   - Deletion of a still-referenced style falls back to replace-with-
 *     [No Paragraph Style] instead of silently surviving.
 *
 * v1.1 CHANGES — per-page cleanup (both optional, off by default)
 *   - "Styles" audit frame: creates a text frame on each split page (top-
 *     left, 5 mm inset), applies the object style named in the dialog
 *     (default "AllignTop"), names AND script-labels it "Styles", and lists
 *     the paragraph styles in use ON THAT PAGE — one line per style, each
 *     line set in itself. If the object style isn't found in the document
 *     the frame is still created (auto-sizing) and the file is flagged.
 *   - Purge unused paragraph styles: deletes every paragraph style not in
 *     use anywhere in the split file. "In use" is computed from ALL stories
 *     (page, parent pages, pasteboard, tables, footnotes), plus BASED-ON
 *     ancestors of used styles, the document text default, and paragraph
 *     styles referenced by object styles — so nothing that still matters
 *     can be deleted. Styles are resolved by ID at removal time (no stale
 *     index references). Emptied style groups are pruned. [Basic Paragraph]
 *     and [No Paragraph Style] are never touched.
 *   The audit frame is created BEFORE the purge, so its own styled lines
 *   keep their styles alive; the frame lists page-scope usage, the purge
 *   keeps whole-file usage — parents and pasteboard are protected without
 *   being listed.
 *
 * HOW IT WORKS
 *   For each page: saveACopy of the full document to a temp file, open it
 *   hidden, delete every other page, then Save As (.indd) or export (.idml)
 *   to the final name, close without saving, remove the temp. Full fidelity:
 *   styles, swatches, layers and parent pages all survive (each split file
 *   carries the complete set — heavier than strictly needed, but safe).
 *   The ORIGINAL document is never modified or saved.
 *
 * NAMING
 *   Sequential:  Base_1, Base_2, ...          (unpadded up to 9 pages)
 *                Base_01 ... Base_12          (padded once the count is
 *                                              double digits; 3 digits at
 *                                              100+, same rule)
 *   Cover mode:  Base_Cover, Base_1..X, Base_BackCover
 *                where X = pages - 2, padded by the same rule. The Cover /
 *                BackCover labels are editable. Needs at least 2 pages.
 *
 * SAFETY / BEHAVIOUR
 *   - Preflight warns (with a count) if any story threads across pages:
 *     split pages will REFLOW in that case, because the other frames of the
 *     thread no longer exist. Proceed knowingly.
 *   - User interaction is suppressed during the batch, so supplied files
 *     with missing links/fonts won't stall on warning dialogs. The final
 *     report lists every file written and any failures.
 *   - Existing outputs: overwrite (optional) or auto " (2)" suffix.
 */

#target "InDesign"

(function () {
    if (!app.documents.length) { alert("Open a document first."); return; }
    var doc = app.activeDocument;
    var N = doc.pages.length;
    if (N < 1) { alert("Document has no pages."); return; }

    // ---------------- preflight: cross-page text threads ----------------

    var threaded = 0;
    try {
        for (var s = 0; s < doc.stories.length; s++) {
            var tcs = doc.stories[s].textContainers;
            var seen = {}, pagesHit = 0;
            for (var t = 0; t < tcs.length; t++) {
                try {
                    var pp = tcs[t].parentPage;
                    if (pp && pp.isValid && !seen[pp.documentOffset]) { seen[pp.documentOffset] = true; pagesHit++; }
                } catch (_) {}
            }
            if (pagesHit > 1) threaded++;
        }
    } catch (_) {}

    // ---------------- dialog ----------------

    var cfg = showDialog();
    if (!cfg) return;

    if (threaded > 0) {
        if (!confirm("WARNING: " + threaded + " stor" + (threaded === 1 ? "y" : "ies") +
                     " thread(s) across multiple pages.\n\nSplit pages will REFLOW \u2014 the other frames of " +
                     "each thread no longer exist in the single-page files.\n\nContinue anyway?")) return;
    }

    // ---------------- build the job list ----------------

    var jobs = [];   // {offset, fileBase}
    var padSeq = (N >= 10) ? String(N).length : 1;
    if (cfg.coverMode) {
        var X = N - 2;
        var padIn = (X >= 10) ? String(X).length : 1;
        for (var p = 0; p < N; p++) {
            var suffix;
            if (p === 0) suffix = cfg.coverLabel;
            else if (p === N - 1) suffix = cfg.backLabel;
            else suffix = pad(p, padIn);
            jobs.push({ offset: p, fileBase: cfg.base + "_" + suffix });
        }
    } else {
        for (var q = 0; q < N; q++) {
            jobs.push({ offset: q, fileBase: cfg.base + "_" + pad(q + 1, padSeq) });
        }
    }

    // ---------------- run ----------------

    var ext = cfg.idml ? ".idml" : ".indd";
    var lines = [], fails = [], nOk = 0;
    var prog = progressWin("Splitting pages\u2026", jobs.length);

    var oldInteract = app.scriptPreferences.userInteractionLevel;
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    try {
        for (var j = 0; j < jobs.length; j++) {
            var job = jobs[j];
            prog.set(j, "Page " + (job.offset + 1) + "  \u2192  " + job.fileBase + ext + "   \u2014  " + (j + 1) + " of " + jobs.length);

            var outFile = resolveOut(cfg.folder, job.fileBase, ext, cfg.overwrite);
            var tempFile = new File(Folder.temp.fsName + "/idsplit_" + (new Date().getTime()) + "_" + j + ".indd");
            var d2 = null;
            try {
                doc.saveACopy(tempFile);
                d2 = app.open(tempFile, false); // no window
                try { d2.documentPreferences.allowPageShuffle = true; } catch (_) {}

                for (var r = d2.pages.length - 1; r >= 0; r--) {
                    if (r === job.offset) continue;
                    try { d2.pages[r].remove(); } catch (eDel) { throw ("could not remove page " + (r + 1) + ": " + eDel); }
                }
                if (d2.pages.length !== 1) throw ("expected 1 page after deletion, found " + d2.pages.length);

                var cleanNote = "";
                if (cfg.addFrame || cfg.purge) {
                    cleanNote = perPageCleanup(d2, cfg);
                }

                if (cfg.idml) {
                    d2.exportFile(ExportFormat.INDESIGN_MARKUP, outFile);
                } else {
                    d2.save(outFile);
                }
                nOk++;
                lines.push("OK    page " + (job.offset + 1) + "  \u2192  " + outFile.name + cleanNote);
            } catch (eJob) {
                fails.push("Page " + (job.offset + 1) + " ('" + job.fileBase + "'): " + eJob);
            } finally {
                try { if (d2 && d2.isValid) d2.close(SaveOptions.NO); } catch (_) {}
                try { if (tempFile.exists) tempFile.remove(); } catch (_) {}
            }
        }
        prog.set(jobs.length, "Done.");
    } finally {
        app.scriptPreferences.userInteractionLevel = oldInteract;
        prog.close();
    }

    // ---------------- report ----------------

    var rep = [];
    rep.push("Folder: " + cfg.folder.fsName);
    rep.push("Format: " + (cfg.idml ? "IDML" : "INDD") + "    Naming: " + (cfg.coverMode ? "Cover / inner / BackCover" : "Sequential"));
    if (threaded > 0) rep.push("REMINDER: " + threaded + " cross-page thread(s) \u2014 check reflow on the split files.");
    rep.push("");
    for (var L = 0; L < lines.length; L++) rep.push(lines[L]);
    if (fails.length) {
        rep.push("");
        rep.push("ISSUES (" + fails.length + "):");
        for (var F = 0; F < fails.length; F++) rep.push("  " + fails[F]);
    }
    rep.push("");
    rep.push("Summary: " + nOk + " of " + jobs.length + " page file(s) written. Original document untouched.");
    showReport("Split Document Pages \u2014 Report", rep.join("\n"));
    return;

    // =====================================================================

    function pad(n, width) {
        var s = String(n);
        while (s.length < width) s = "0" + s;
        return s;
    }

    function sanitizeName(s) {
        var out = String(s);
        out = out.replace(/[\r\n\u2028\u2029]+/g, " ");
        out = out.replace(/[\\\/:\*\?"<>\|]/g, "");
        out = out.replace(/\s+/g, " ");
        out = out.replace(/^\s+|\s+$/g, "");
        out = out.replace(/[\. ]+$/g, "");
        return out;
    }

    function resolveOut(folder, base, extension, overwrite) {
        var f = new File(folder.fsName + "/" + base + extension);
        if (f.exists && !overwrite) {
            var n = 2;
            while (true) {
                var alt = new File(folder.fsName + "/" + base + " (" + n + ")" + extension);
                if (!alt.exists) return alt;
                n++;
            }
        }
        return f;
    }

    // =====================================================================
    // Per-page cleanup: "Styles" audit frame + unused-style purge
    // =====================================================================

    function perPageCleanup(d2, cfg) {
        var oldU = app.scriptPreferences.measurementUnit;
        app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
        try {
            var pg = d2.pages[0];
            var notes = [];

            var pageStyles = pageStyleList(d2, pg);   // [{style, path}] page scope

            if (cfg.addFrame) {
                var missing = buildStylesFrame(d2, pg, pageStyles, cfg);
                notes.push("styles listed: " + pageStyles.length + (missing ? " (object style '" + cfg.objStyleName + "' NOT FOUND)" : ""));
            }

            if (cfg.purge) {
                var keepIds = null;
                if (cfg.purgePageOnly) {
                    keepIds = {};
                    for (var kI = 0; kI < pageStyles.length; kI++) {
                        try { keepIds[String(pageStyles[kI].style.id)] = true; } catch (_) {}
                    }
                }
                var removed = purgeUnusedParaStyles(d2, keepIds);
                notes.push("styles removed: " + removed + (cfg.purgePageOnly ? " (page-only)" : ""));
            }
            return notes.length ? "   [" + notes.join(", ") + "]" : "";
        } catch (eC) {
            return "   [cleanup failed: " + eC + "]";
        } finally {
            app.scriptPreferences.measurementUnit = oldU;
        }
    }

    // ---- page-scoped in-use paragraph styles (frames, tables, type-on-path)

    function pageStyleList(d2, pg) {
        var noParaId = null;
        try { noParaId = d2.paragraphStyles[0].id; } catch (_) {}
        var seen = {}, out = [];
        function note(st) {
            try {
                if (!st || !st.isValid) return;
                if (noParaId !== null && st.id === noParaId) return;
                var k = String(st.id);
                if (seen[k]) return;
                seen[k] = true;
                out.push({ style: st, path: paraStylePath(st) });
            } catch (_) {}
        }
        function harvest(txt) {
            try {
                var stys = txt.paragraphs.everyItem().appliedParagraphStyle.getElements();
                for (var i = 0; i < stys.length; i++) note(stys[i]);
            } catch (_) {
                try { for (var j = 0; j < txt.paragraphs.length; j++) note(txt.paragraphs[j].appliedParagraphStyle); } catch (__) {}
            }
            try {
                var tabs = txt.tables;
                for (var t = 0; t < tabs.length; t++) {
                    var cells = tabs[t].cells;
                    for (var c = 0; c < cells.length; c++) {
                        try {
                            var ps = cells[c].paragraphs;
                            for (var q = 0; q < ps.length; q++) note(ps[q].appliedParagraphStyle);
                        } catch (_) {}
                    }
                }
            } catch (_) {}
        }
        function onPage(item) {
            try {
                var pp = item.parentPage;
                return pp && pp.isValid && pp.documentOffset === pg.documentOffset;
            } catch (_) { return false; }
        }
        try {
            var pool = pg.parent.allPageItems;
            for (var i = 0; i < pool.length; i++) {
                var it = pool[i];
                try {
                    if (!it.isValid) continue;
                    var cn = it.constructor ? it.constructor.name : "";
                    if (cn === "TextFrame" && onPage(it)) { try { harvest(it.texts[0]); } catch (_) {} }
                    try {
                        if (it.textPaths && it.textPaths.length && onPage(it)) {
                            for (var tp = 0; tp < it.textPaths.length; tp++) {
                                try { harvest(it.textPaths[tp].texts[0]); } catch (_) {}
                            }
                        }
                    } catch (_) {}
                } catch (_) {}
            }
        } catch (_) {}
        out.sort(function (a, b) {
            var A = a.path.toLowerCase(), B = b.path.toLowerCase();
            return A < B ? -1 : (A > B ? 1 : 0);
        });
        return out;
    }

    function paraStylePath(style) {
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

    // ---- audit frame

    function buildStylesFrame(d2, pg, list, cfg) {
        var INSET = 14, WIDTH = 255;
        var pb = pg.bounds;
        var tf = pg.parent.textFrames.add();
        tf.geometricBounds = [pb[0] + INSET, pb[1] + INSET, pb[0] + INSET + 60, pb[1] + INSET + WIDTH];

        var os = findObjectStyleByName(d2, cfg.objStyleName);
        var missing = true;
        if (os) {
            try { tf.appliedObjectStyle = os; missing = false; } catch (_) {}
        }
        if (missing) {
            try {
                tf.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
                tf.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
            } catch (_) {}
        }
        try { tf.name = cfg.frameName; } catch (_) {}
        try { tf.label = cfg.frameName; } catch (_) {}

        if (!list.length) {
            tf.contents = "(no paragraph styles in use on this page)";
            return missing;
        }
        var linesArr = [];
        for (var i = 0; i < list.length; i++) linesArr.push(list[i].path);
        tf.contents = linesArr.join("\r");
        for (var j = 0; j < list.length && j < tf.paragraphs.length; j++) {
            try { tf.paragraphs[j].appliedParagraphStyle = list[j].style; } catch (_) {}
        }
        return missing;
    }

    function findObjectStyleByName(d2, name) {
        try { var os = d2.objectStyles.itemByName(name); if (os && os.isValid) return os; } catch (_) {}
        function walk(container) {
            var i, s;
            try {
                for (i = 0; i < container.objectStyles.length; i++) { s = container.objectStyles[i]; if (s.name === name) return s; }
                for (i = 0; i < container.objectStyleGroups.length; i++) { s = walk(container.objectStyleGroups[i]); if (s) return s; }
            } catch (_) {}
            return null;
        }
        return walk(d2);
    }

    // ---- purge

    function purgeUnusedParaStyles(d2, pageKeep) {
        var keep = {}; // id -> true
        function keepStyle(st) {
            try {
                if (!st || !st.isValid) return;
                if ((st.constructor ? st.constructor.name : "") !== "ParagraphStyle") return;
                var k = String(st.id);
                if (keep[k]) return;
                keep[k] = true;
                // based-on ancestors ride along
                var p = st, guard = 0;
                while (guard++ < 16) {
                    var b = null;
                    try { b = p.basedOn; } catch (_) { break; }
                    if (!b || !b.isValid || (b.constructor ? b.constructor.name : "") !== "ParagraphStyle") break;
                    var kb = String(b.id);
                    if (keep[kb]) break;
                    keep[kb] = true;
                    p = b;
                }
            } catch (_) {}
        }
        function harvestUsage(txt) {
            try {
                var stys = txt.paragraphs.everyItem().appliedParagraphStyle.getElements();
                for (var i = 0; i < stys.length; i++) keepStyle(stys[i]);
            } catch (_) {
                try { for (var j = 0; j < txt.paragraphs.length; j++) keepStyle(txt.paragraphs[j].appliedParagraphStyle); } catch (__) {}
            }
        }

        if (pageKeep) {
            // upload mode: the keep-set is exactly the page's styles
            for (var pk in pageKeep) { if (pageKeep.hasOwnProperty(pk)) keep[pk] = true; }
        } else {

        // whole file: every story (page, parents, pasteboard), tables, footnotes
        try {
            for (var s = 0; s < d2.stories.length; s++) {
                var story = d2.stories[s];
                harvestUsage(story);
                try {
                    var tabs = story.tables;
                    for (var t = 0; t < tabs.length; t++) {
                        var cells = tabs[t].cells;
                        for (var c = 0; c < cells.length; c++) { try { harvestUsage(cells[c].texts[0]); } catch (_) {} }
                    }
                } catch (_) {}
                try {
                    var fns = story.footnotes;
                    for (var f = 0; f < fns.length; f++) { try { harvestUsage(fns[f].texts[0]); } catch (_) {} }
                } catch (_) {}
            }
        } catch (_) {}

        // document default + object-style references
        try { keepStyle(d2.textDefaults.appliedParagraphStyle); } catch (_) {}
        try {
            var allOS = d2.allObjectStyles;
            for (var o = 0; o < allOS.length; o++) { try { keepStyle(allOS[o].appliedParagraphStyle); } catch (_) {} }
        } catch (_) {}

        } // end whole-file keep-set

        // built-ins are never candidates
        try { keep[String(d2.paragraphStyles[0].id)] = true; } catch (_) {}
        try { var bp = d2.paragraphStyles.itemByName("[Basic Paragraph]"); if (bp.isValid) keep[String(bp.id)] = true; } catch (_) {}

        // snapshot candidate IDs, then resolve fresh by ID at removal time
        var candidates = [];
        try {
            var all = d2.allParagraphStyles;
            for (var a = 0; a < all.length; a++) {
                try {
                    var id = all[a].id;
                    if (!keep[String(id)]) candidates.push(id);
                } catch (_) {}
            }
        } catch (_) {}

        var removed = 0;
        for (var r = 0; r < candidates.length; r++) {
            var st2 = resolveParaStyleById(d2, candidates[r]);
            if (!st2) continue;
            try { st2.remove(); removed++; }
            catch (_) {
                // still referenced (parents/pasteboard in page-only mode):
                // replace with [No Paragraph Style] — appearance is preserved
                // as local overrides on the affected text
                try { st2.remove(d2.paragraphStyles[0]); removed++; } catch (__) { /* leave it */ }
            }
        }

        pruneEmptyParaGroups(d2);
        return removed;
    }

    function resolveParaStyleById(d2, id) {
        try {
            var all = d2.allParagraphStyles;
            for (var i = 0; i < all.length; i++) {
                try { if (all[i].id === id) return all[i]; } catch (_) {}
            }
        } catch (_) {}
        return null;
    }

    function pruneEmptyParaGroups(container) {
        try {
            var groups = container.paragraphStyleGroups;
            for (var i = groups.length - 1; i >= 0; i--) {
                var g = groups[i];
                try {
                    pruneEmptyParaGroups(g);
                    if (g.paragraphStyles.length === 0 && g.paragraphStyleGroups.length === 0) g.remove();
                } catch (_) {}
            }
        } catch (_) {}
    }

    // ---------------- UI ----------------

    function showDialog() {
        var w = new Window("dialog", "Split Document Pages \u2014 v1.0");
        w.orientation = "column"; w.alignChildren = "fill";

        var defBase = sanitizeName(doc.name.replace(/\.indd$/i, ""));
        var defFolder = doc.saved ? doc.filePath.fsName : Folder.desktop.fsName;

        var g1 = w.add("group");
        g1.add("statictext", [0, 0, 90, 20], "Base name:");
        var etBase = g1.add("edittext", [0, 0, 340, 24], defBase);
        g1.add("statictext", undefined, "(page suffix is appended: base_1 / base_Cover)");

        var g2 = w.add("group");
        g2.add("statictext", [0, 0, 90, 20], "Output folder:");
        var etFolder = g2.add("edittext", [0, 0, 440, 24], defFolder);
        var btnBr = g2.add("button", undefined, "Browse\u2026");
        btnBr.onClick = function () {
            var f = Folder.selectDialog("Choose the output folder", new Folder(etFolder.text));
            if (f) etFolder.text = f.fsName;
        };

        var gF = w.add("panel", undefined, "Format");
        gF.orientation = "row"; gF.margins = 12;
        var rbIndd = gF.add("radiobutton", undefined, "INDD");
        var rbIdml = gF.add("radiobutton", undefined, "IDML");
        rbIndd.value = true;

        var gN = w.add("panel", undefined, "Naming  \u2014  " + N + " page" + (N === 1 ? "" : "s") + " in this document");
        gN.orientation = "column"; gN.alignChildren = "left"; gN.margins = 12;
        var rbSeq = gN.add("radiobutton", undefined, "Sequential: 1, 2, 3\u2026  (leading zeros once the count is double digits)");
        var rbCov = gN.add("radiobutton", undefined, "First page = cover, last page = back cover, inner pages 1\u2026" + Math.max(0, N - 2));
        rbSeq.value = true;
        var gLab = gN.add("group");
        gLab.add("statictext", undefined, "Cover label:");
        var etCov = gLab.add("edittext", [0, 0, 110, 24], "Cover");
        gLab.add("statictext", undefined, "Back cover label:");
        var etBack = gLab.add("edittext", [0, 0, 110, 24], "BackCover");
        etCov.enabled = etBack.enabled = false;
        rbSeq.onClick = function () { etCov.enabled = etBack.enabled = false; };
        rbCov.onClick = function () { etCov.enabled = etBack.enabled = true; };

        var gO = w.add("group");
        var cbOver = gO.add("checkbox", undefined, "Overwrite existing files (off = auto-suffix)");
        cbOver.value = false;

        var gC = w.add("panel", undefined, "Per-page cleanup (optional)");
        gC.orientation = "column"; gC.alignChildren = "left"; gC.margins = 12;
        var gC1 = gC.add("group");
        var cbFrame = gC1.add("checkbox", undefined, "Add 'Styles' audit frame listing this page's paragraph styles \u2014 object style:");
        var etOS = gC1.add("edittext", [0, 0, 120, 24], "AllignTop");
        var gC2 = gC.add("group");
        gC2.add("statictext", undefined, "Frame name && script label:");
        var etFN = gC2.add("edittext", [0, 0, 120, 24], "Styles");
        var cbPurge = gC.add("checkbox", undefined, "Delete unused paragraph styles from each page file");
        var gC3 = gC.add("group");
        var rbPageOnly = gC3.add("radiobutton", undefined, "Keep ONLY this page's styles (upload mode \u2014 parent/pasteboard text keeps its look, loses style links)");
        var rbSafe = gC3.add("radiobutton", undefined, "Also keep parent/pasteboard usage + based-on ancestors");
        rbPageOnly.value = true;
        cbFrame.value = false; cbPurge.value = false;
        etOS.enabled = etFN.enabled = false;
        rbPageOnly.enabled = rbSafe.enabled = false;
        cbFrame.onClick = function () { etOS.enabled = etFN.enabled = cbFrame.value; };
        cbPurge.onClick = function () { rbPageOnly.enabled = rbSafe.enabled = cbPurge.value; };

        var gB = w.add("group"); gB.alignment = "right";
        gB.add("button", undefined, "Cancel", { name: "cancel" });
        var ok = gB.add("button", undefined, "Split", { name: "ok" });

        var result = null;
        ok.onClick = function () {
            var base = sanitizeName(etBase.text);
            if (base === "") { alert("Enter a base name."); return; }
            if (rbCov.value && N < 2) { alert("Cover mode needs at least 2 pages."); return; }
            var covL = sanitizeName(etCov.text), backL = sanitizeName(etBack.text);
            if (rbCov.value && (covL === "" || backL === "")) { alert("Cover / back cover labels can't be empty."); return; }
            var folder = new Folder(etFolder.text);
            if (!folder.exists) {
                if (!confirm("Folder does not exist:\n" + etFolder.text + "\n\nCreate it?")) return;
                if (!folder.create()) { alert("Could not create the folder."); return; }
            }
            var osName = sanitizeName(etOS.text), frName = sanitizeName(etFN.text);
            if (cbFrame.value && (osName === "" || frName === "")) { alert("Object style and frame name can't be empty when the audit frame is on."); return; }
            result = {
                base: base,
                folder: folder,
                idml: rbIdml.value,
                coverMode: rbCov.value,
                coverLabel: covL,
                backLabel: backL,
                overwrite: cbOver.value,
                addFrame: cbFrame.value,
                objStyleName: osName,
                frameName: frName,
                purge: cbPurge.value,
                purgePageOnly: rbPageOnly.value
            };
            w.close(1);
        };

        return (w.show() === 1) ? result : null;
    }

    function progressWin(title, max) {
        var w = new Window("palette", title);
        w.orientation = "column"; w.alignChildren = "fill";
        var t = w.add("statictext", [0, 0, 420, 18], "");
        var pb = w.add("progressbar", [0, 0, 420, 12], 0, Math.max(1, max));
        w.show();
        return {
            set: function (v, msg) { try { pb.value = v; t.text = msg; w.update(); } catch (_) {} },
            close: function () { try { w.close(); } catch (_) {} }
        };
    }

    function showReport(title, text) {
        var w = new Window("dialog", title);
        w.orientation = "column"; w.alignChildren = "fill";
        w.add("edittext", [0, 0, 720, 400], text, { multiline: true, readonly: true, scrolling: true });
        var g = w.add("group"); g.alignment = "right";
        g.add("button", undefined, "Close", { name: "ok" });
        w.show();
    }

})();
