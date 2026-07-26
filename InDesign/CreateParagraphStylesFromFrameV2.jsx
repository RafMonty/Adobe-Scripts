/**
 * CreateParagraphStylesFromFrame.jsx — v2.0
 * Create PARAGRAPH (and/or character) styles from the lines of a selected
 * text frame — each line's formatting becomes a style named after its text.
 * InDesign ExtendScript (2023+ Mac/Win)
 *
 * WHY v2 (review of the old CreateCharStylesFromTextFrame.jsx)
 *   - The old script created CHARACTER styles, which cannot hold paragraph
 *     formatting at all (alignment, indents, tabs, spacing, rules, keeps,
 *     hyphenation) — structural, not a bug.
 *   - It copied a ~20-property whitelist; everything else was dropped.
 *   - One try{} wrapped all assignments: the first throwing property
 *     silently aborted the rest of the formatting.
 *
 * HOW v2 CAPTURES
 *   The line's first character's FULL .properties snapshot (character AND
 *   paragraph attributes) is copied into the style key-by-key, each key in
 *   its own try/catch — everything writable lands, inapplicable keys skip
 *   individually. tabList is rebuilt as plain stops (live references don't
 *   assign). The style is made fully explicit: basedOn [No Paragraph
 *   Style], nextStyle itself. The report shows how many attributes each
 *   style captured.
 *
 * FEATURES
 *   - Create Paragraph styles, Character styles, or Both (same engine —
 *     char-style targets simply reject paragraph-level keys).
 *   - '/' in line text = style GROUP path: "H/Header" creates group H.
 *     Round-trips with the audit tools' path listing.
 *   - Name collisions: Update existing / Skip / Create with suffix.
 *   - Optional apply-back to the source line (clean, override-free state).
 *   - Mixed-formatting warning when a line's first and last characters
 *     differ (first character wins).
 *   - Optional "delete ALL existing styles first" reset — now group-aware
 *     (allParagraphStyles/allCharacterStyles), ID-resolved, built-ins kept,
 *     with confirmation.
 *   - One undo step. ScriptUI dialogs.
 */

#target "InDesign"

(function () {
    if (!app.documents.length) { alert("Open a document first."); return; }
    var doc = app.activeDocument;

    // ---------------- selection ----------------

    var frame = resolveSelectedFrame();
    if (!frame) { alert("Select a text frame (or click into one) first."); return; }
    if (frame.paragraphs.length === 0) { alert("The selected text frame is empty."); return; }

    // ---------------- options ----------------

    var cfg = showDialog();
    if (!cfg) return;
    if (cfg.nuke) {
        if (!confirm("Really delete ALL existing paragraph and character styles first?\n\n" +
                     "Styled text keeps its appearance as local formatting. Built-ins are kept. One undo step reverts.")) return;
    }

    var R = { created: 0, updated: 0, skipped: 0, nukedP: 0, nukedC: 0, lines: [], warnings: [], groupsMade: 0 };

    app.doScript(run, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Create Styles From Frame");

    showReport("Create Styles From Frame \u2014 Report", buildReport());
    return;

    // =====================================================================

    function run() {
        if (cfg.nuke) nukeStyles();

        var paras = frame.paragraphs;
        for (var i = 0; i < paras.length; i++) {
            var para = paras[i];
            var raw = "";
            try { raw = String(para.contents); } catch (_) {}
            var text = raw.replace(/[\r\n\u2028\u2029]+/g, "").replace(/^\s+|\s+$/g, "");
            if (text === "") { R.skipped++; R.lines.push("Line " + (i + 1) + ": skipped (empty)"); continue; }
            if (para.characters.length === 0) { R.skipped++; R.lines.push("Line " + (i + 1) + ": skipped (no characters)"); continue; }

            // group path + name
            var segs = cfg.groups ? text.split("/") : [text];
            var cleanSegs = [];
            for (var g = 0; g < segs.length; g++) {
                var cs = cleanName(segs[g]);
                if (cs !== "") cleanSegs.push(cs);
            }
            if (!cleanSegs.length) { R.skipped++; R.warnings.push("Line " + (i + 1) + ": name empty after cleaning \u2014 skipped."); continue; }
            var leafName = cleanSegs.pop();
            var groupPath = cleanSegs; // remaining = groups

            // mixed formatting check (first vs last character, light)
            if (para.characters.length > 1) {
                try {
                    var a = para.characters[0], b = para.characters[para.characters.length - 1];
                    var mixed = false;
                    try { if (Number(a.pointSize) !== Number(b.pointSize)) mixed = true; } catch (_) {}
                    try { if (String(a.appliedFont.name) !== String(b.appliedFont.name)) mixed = true; } catch (_) {}
                    if (mixed) R.warnings.push("Line " + (i + 1) + " ('" + leafName + "'): mixed formatting \u2014 first character used.");
                } catch (_) {}
            }

            var snap = null;
            try { snap = para.characters[0].properties; } catch (_) {}
            if (!snap) { R.skipped++; R.warnings.push("Line " + (i + 1) + ": could not read formatting \u2014 skipped."); continue; }

            var madeP = null;
            if (cfg.doPara) madeP = buildStyle(true, groupPath, leafName, snap, para, i);
            if (cfg.doChar) buildStyle(false, groupPath, leafName, snap, para, i);

            if (cfg.applyBack && madeP) {
                try { para.appliedParagraphStyle = madeP; } catch (_) {}
            }
        }
    }

    // ---------------- style building ----------------

    function buildStyle(isPara, groupPath, leafName, snap, para, lineIdx) {
        var container = doc, g;
        for (g = 0; g < groupPath.length; g++) {
            container = getOrCreateGroup(container, groupPath[g], isPara);
            if (!container) { R.warnings.push("Line " + (lineIdx + 1) + ": could not create group '" + groupPath[g] + "'."); return null; }
        }
        var coll = isPara ? container.paragraphStyles : container.characterStyles;

        var existing = null;
        try { var hit = coll.itemByName(leafName); if (hit.isValid) existing = hit; } catch (_) {}

        var style = null, verb = "";
        if (existing) {
            if (cfg.collide === "skip") {
                R.skipped++;
                R.lines.push("Line " + (lineIdx + 1) + ": '" + fullPath(groupPath, leafName) + "' exists \u2014 skipped.");
                return existing; // still usable for apply-back
            }
            if (cfg.collide === "suffix") {
                var nm = leafName, n = 2;
                while (true) {
                    var probe = null;
                    try { probe = coll.itemByName(nm); } catch (_) {}
                    if (!probe || !probe.isValid) break;
                    nm = leafName + " " + (n++);
                }
                leafName = nm;
                existing = null;
            } else {
                style = existing; verb = "Updated";
            }
        }
        if (!style) {
            try { style = coll.add({ name: leafName }); verb = "Created"; }
            catch (eA) { R.warnings.push("Line " + (lineIdx + 1) + ": could not create '" + leafName + "' (" + eA + ")."); return null; }
        }

        var applied = pourProperties(style, snap);

        // tabs rebuilt as plain stops (live refs don't assign)
        if (isPara) {
            try {
                var tl = snap.tabList;
                if (tl && tl.length !== undefined) {
                    var plain = [];
                    for (var t = 0; t < tl.length; t++) {
                        try {
                            var o = { position: Number(tl[t].position) };
                            try { o.alignment = tl[t].alignment; } catch (_) {}
                            try { o.leader = tl[t].leader; } catch (_) {}
                            try { o.alignmentCharacter = tl[t].alignmentCharacter; } catch (_) {}
                            plain.push(o);
                        } catch (_) {}
                    }
                    style.tabList = plain;
                }
            } catch (_) {}
            try { style.basedOn = doc.paragraphStyles[0]; } catch (_) {}
            try { style.nextStyle = style; } catch (_) {}
        }

        if (verb === "Created") R.created++; else R.updated++;
        R.lines.push("Line " + (lineIdx + 1) + ": " + verb + " " + (isPara ? "\u00b6" : "A") + " '" +
                     fullPath(groupPath, leafName) + "'  (" + applied + " attributes)");
        return style;
    }

    // Every writable key from the snapshot, each in its own try — one bad
    // key can never drop the rest (the old script's core failure).
    function pourProperties(style, snap) {
        var BLOCK = {
            appliedParagraphStyle: 1, appliedCharacterStyle: 1, appliedLanguage: 0, // language IS wanted
            contents: 1, parent: 1, parentStory: 1, parentTextFrames: 1,
            index: 1, id: 1, label: 1, name: 1, isValid: 1, properties: 1,
            associatedXMLElements: 1, storyOffset: 1, tabList: 1, // handled separately
            appliedConditions: 1, characters: 1, words: 1, lines: 1,
            paragraphs: 1, texts: 1, insertionPoints: 1, textStyleRanges: 1,
            textColumns: 1, tables: 1, footnotes: 1, notes: 1, hyperlinks: 1,
            bulletsAndNumberingListType: 0
        };
        var n = 0;
        for (var k in snap) {
            if (!snap.hasOwnProperty(k)) continue;
            if (BLOCK[k] === 1) continue;
            try {
                style[k] = snap[k];
                n++;
            } catch (_) { /* not writable / not applicable on this style type */ }
        }
        return n;
    }

    function getOrCreateGroup(container, name, isPara) {
        var groups = isPara ? container.paragraphStyleGroups : container.characterStyleGroups;
        try { var hit = groups.itemByName(name); if (hit.isValid) return hit; } catch (_) {}
        try { var made = groups.add({ name: name }); R.groupsMade++; return made; } catch (_) {}
        return null;
    }

    function fullPath(groupPath, leaf) {
        return groupPath.length ? groupPath.join("/") + "/" + leaf : leaf;
    }

    // ---------------- nuke (optional reset, group-aware, ID-resolved) ----

    function nukeStyles() {
        var keepP = {}, keepC = {};
        try { keepP[String(doc.paragraphStyles[0].id)] = 1; } catch (_) {}
        try { var bp = doc.paragraphStyles.itemByName("[Basic Paragraph]"); if (bp.isValid) keepP[String(bp.id)] = 1; } catch (_) {}
        try { keepC[String(doc.characterStyles[0].id)] = 1; } catch (_) {}

        var ids = [], i;
        try { var ap = doc.allParagraphStyles; for (i = 0; i < ap.length; i++) { try { if (!keepP[String(ap[i].id)]) ids.push({ p: true, id: ap[i].id }); } catch (_) {} } } catch (_) {}
        try { var ac = doc.allCharacterStyles; for (i = 0; i < ac.length; i++) { try { if (!keepC[String(ac[i].id)]) ids.push({ p: false, id: ac[i].id }); } catch (_) {} } } catch (_) {}

        for (i = 0; i < ids.length; i++) {
            var st = resolveStyleById(ids[i].p, ids[i].id);
            if (!st) continue;
            try { st.remove(); if (ids[i].p) R.nukedP++; else R.nukedC++; }
            catch (_) {
                try { st.remove(ids[i].p ? doc.paragraphStyles[0] : doc.characterStyles[0]); if (ids[i].p) R.nukedP++; else R.nukedC++; }
                catch (__) {}
            }
        }
        pruneGroups(doc, true);
        pruneGroups(doc, false);
    }

    function resolveStyleById(isPara, id) {
        try {
            var all = isPara ? doc.allParagraphStyles : doc.allCharacterStyles;
            for (var i = 0; i < all.length; i++) { try { if (all[i].id === id) return all[i]; } catch (_) {} }
        } catch (_) {}
        return null;
    }

    function pruneGroups(container, isPara) {
        try {
            var groups = isPara ? container.paragraphStyleGroups : container.characterStyleGroups;
            for (var i = groups.length - 1; i >= 0; i--) {
                var g = groups[i];
                try {
                    pruneGroups(g, isPara);
                    var empty = isPara
                        ? (g.paragraphStyles.length === 0 && g.paragraphStyleGroups.length === 0)
                        : (g.characterStyles.length === 0 && g.characterStyleGroups.length === 0);
                    if (empty) g.remove();
                } catch (_) {}
            }
        } catch (_) {}
    }

    // ---------------- helpers ----------------

    function resolveSelectedFrame() {
        try {
            if (app.selection.length !== 1) return null;
            var s = app.selection[0];
            var cn = s.constructor ? s.constructor.name : "";
            if (cn === "TextFrame") return s;
            if (cn === "InsertionPoint" || cn === "Text" || cn === "Character" ||
                cn === "Word" || cn === "Line" || cn === "Paragraph" || cn === "TextStyleRange") {
                if (s.parentTextFrames && s.parentTextFrames.length) return s.parentTextFrames[0];
            }
        } catch (_) {}
        return null;
    }

    function cleanName(s) {
        var out = String(s);
        out = out.replace(/[\x00-\x1F\x7F-\x9F]+/g, "");
        out = cfg && cfg.noSpaces ? out.replace(/\s+/g, "") : out.replace(/\s+/g, " ");
        out = out.replace(/[\\:\*\?"<>\|]/g, "");   // '/' already consumed as group separator
        out = out.replace(/^\s+|\s+$/g, "");
        if (out.length > 100) out = out.substring(0, 100);
        return out;
    }

    // ---------------- UI ----------------

    function showDialog() {
        var w = new Window("dialog", "Create Styles From Frame \u2014 v2.0");
        w.orientation = "column"; w.alignChildren = "left";

        var gT = w.add("panel", undefined, "Create");
        gT.orientation = "row"; gT.margins = 12;
        var rbP = gT.add("radiobutton", undefined, "Paragraph styles");
        var rbC = gT.add("radiobutton", undefined, "Character styles");
        var rbB = gT.add("radiobutton", undefined, "Both");
        rbP.value = true;

        var gN = w.add("panel", undefined, "Naming");
        gN.orientation = "column"; gN.alignChildren = "left"; gN.margins = 12;
        var cbGroups = gN.add("checkbox", undefined, "Treat '/' in line text as a style-group path  (H/Header \u2192 group H)");
        var cbNoSp = gN.add("checkbox", undefined, "Remove spaces from style names");
        cbGroups.value = true; cbNoSp.value = true;

        var gX = w.add("panel", undefined, "If a style of that name already exists");
        gX.orientation = "row"; gX.margins = 12;
        var rbUpd = gX.add("radiobutton", undefined, "Update its definition");
        var rbSkip = gX.add("radiobutton", undefined, "Skip");
        var rbSuf = gX.add("radiobutton", undefined, "Create with suffix");
        rbUpd.value = true;

        var cbApply = w.add("checkbox", undefined, "Apply each created paragraph style back to its line (clean, override-free)");
        cbApply.value = true;
        var cbNuke = w.add("checkbox", undefined, "DELETE ALL existing paragraph && character styles first (group-aware; built-ins kept)");
        cbNuke.value = false;

        var gB = w.add("group"); gB.alignment = "right";
        gB.add("button", undefined, "Cancel", { name: "cancel" });
        gB.add("button", undefined, "Run", { name: "ok" });

        if (w.show() !== 1) return null;
        return {
            doPara: rbP.value || rbB.value,
            doChar: rbC.value || rbB.value,
            groups: cbGroups.value,
            noSpaces: cbNoSp.value,
            collide: rbUpd.value ? "update" : (rbSkip.value ? "skip" : "suffix"),
            applyBack: cbApply.value,
            nuke: cbNuke.value
        };
    }

    function buildReport() {
        var L = [];
        if (cfg.nuke) L.push("Reset: deleted " + R.nukedP + " paragraph + " + R.nukedC + " character style(s) (built-ins kept, empty groups pruned).");
        L.push("Created: " + R.created + "    Updated: " + R.updated + "    Skipped: " + R.skipped +
               (R.groupsMade ? "    Groups made: " + R.groupsMade : ""));
        L.push("");
        for (var i = 0; i < R.lines.length; i++) L.push(R.lines[i]);
        if (R.warnings.length) {
            L.push("");
            L.push("WARNINGS (" + R.warnings.length + "):");
            for (var wI = 0; wI < R.warnings.length; wI++) L.push("  " + R.warnings[wI]);
        }
        L.push("");
        L.push("One undo step reverts the whole run.");
        return L.join("\n");
    }

    function showReport(title, text) {
        var w = new Window("dialog", title);
        w.orientation = "column"; w.alignChildren = "fill";
        w.add("edittext", [0, 0, 720, 420], text, { multiline: true, readonly: true, scrolling: true });
        var g = w.add("group"); g.alignment = "right";
        g.add("button", undefined, "Close", { name: "ok" });
        w.show();
    }

})();
