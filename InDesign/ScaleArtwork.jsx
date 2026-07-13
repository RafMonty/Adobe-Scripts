/* =========================================================================
   ScaleArtwork.jsx  —  v1.1
   Scale a large-format InDesign document to a new size, with optional
   bleed added by ENLARGING THE PAGE (trim + bleed becomes the page size).

   v1.1 — Rebuilt the placement engine. v1 scaled items around a fixed
          pasteboard anchor and assumed the page stayed put through
          page.resize(); InDesign renormalises page/pasteboard geometry
          on resize, which displaced artwork off the page on large
          documents. v1.1 records each item's offset from the trim
          top-left, scales items in place, resizes the pages, then
          MEASURES where the page actually landed and places every item
          deterministically at trimTL + (offset x scale). Placement is
          verified and the max drift is reported in the summary.

   WHAT IT DOES
   ------------
   1. Scales every page and every page item (incl. pasteboard items) by a
      user-entered percentage, anchored at each page's top-left.
   2. Uses "Apply to Content" scaling so point sizes, leading, indents,
      tabs, stroke weights and effects become REAL new values — no
      "48pt (12pt)" bracketed values.
   3. Optionally extends each page outward by a bleed amount. This does
      NOT touch InDesign's document bleed settings — the page itself is
      enlarged, so the delivered page = trim + bleed. Bleed can be
      entered at original artwork scale (scales down with the doc) or as
      a final value.
   4. Scales STYLE DEFINITIONS (paragraph, character, object, table,
      cell) so styled text/objects don't end up as overrides against
      unscaled styles. Without this, "Clear Overrides" would blow text
      back up to original size, and GREP styles would re-apply unscaled
      point sizes.
   5. Runs redefineScaling() on all items afterwards so every frame
      reads 100% / 100% (placed graphics are skipped so their scale %
      stays meaningful in the Links panel).
   6. Optionally scales guides, margins/columns, and the baseline /
      document grids; optionally zeroes the document bleed values;
      optionally reports overset stories at the end.

   INSTALL
   -------
   Window > Utilities > Scripts > right-click "User" > Reveal in
   Explorer/Finder > drop this file in "Scripts Panel". Double-click to
   run with a document open.

   NOTES / ASSUMPTIONS
   -------------------
   - Designed for single-page-per-spread artwork documents. Facing-page
     spreads with 2+ pages will trigger a warning (spine-glue behaviour
     of page resize makes multi-page spreads unreliable for this).
   - Style-definition scaling writes computed values back to each style,
     which makes those attributes explicit (flattens based-on
     inheritance for the scaled attributes only). Values stay correct.
   - The script cannot invent artwork: objects must already extend past
     the trim for the new bleed area to be covered. Anything that stops
     dead on the trim edge will show white in the bleed zone.
   - Document Setup may still show the old page size after running
     (pages are resized individually, like the Page tool). The pages
     themselves are authoritative.
   - Undo: one Ctrl/Cmd-Z reverts the whole run.

   Tested targets: InDesign CC 2019+ (ExtendScript). Effects scaling
   requires transformPreferences.adjustEffectsWhenScaling (recent CC);
   the script falls back to manually scaling effect values if absent.
   ========================================================================= */

(function () {

    if (app.documents.length === 0) {
        alert("Open a document first.");
        return;
    }

    var MM2PT = 72 / 25.4;
    var EPS = 0.001; // drift tolerance in pt

    // ---------------------------------------------------------------------
    // Dialog
    // ---------------------------------------------------------------------
    function showDialog() {
        var w = new Window("dialog", "Scale Artwork Document");
        w.orientation = "column";
        w.alignChildren = "fill";

        // --- Scale ---
        var pScale = w.add("panel", undefined, "Scale");
        pScale.orientation = "row";
        pScale.alignChildren = "center";
        pScale.margins = 15;
        pScale.add("statictext", undefined, "New size (% of original):");
        var etScale = pScale.add("edittext", undefined, "50");
        etScale.characters = 7;
        pScale.add("statictext", undefined, "%   (e.g. 25 = quarter size)");

        // --- Bleed ---
        var pBleed = w.add("panel", undefined, "Bleed (extends the page — document bleed settings are NOT used)");
        pBleed.orientation = "column";
        pBleed.alignChildren = "left";
        pBleed.margins = 15;
        var gB = pBleed.add("group");
        gB.add("statictext", undefined, "Bleed each side:");
        var etBleed = gB.add("edittext", undefined, "0");
        etBleed.characters = 7;
        gB.add("statictext", undefined, "mm");
        var rbOrig = pBleed.add("radiobutton", undefined, "Value is at ORIGINAL size (scales down with the artwork)");
        var rbFinal = pBleed.add("radiobutton", undefined, "Value is FINAL size (applied as-is after scaling)");
        rbOrig.value = true;

        // --- Options ---
        var pOpt = w.add("panel", undefined, "Options");
        pOpt.orientation = "column";
        pOpt.alignChildren = "left";
        pOpt.margins = 15;
        var cbStyles = pOpt.add("checkbox", undefined, "Scale style definitions (paragraph / character / object / table / cell)");
        var cbParents = pOpt.add("checkbox", undefined, "Include parent (master) pages");
        var cbGuides = pOpt.add("checkbox", undefined, "Scale ruler guides");
        var cbMargins = pOpt.add("checkbox", undefined, "Scale margins && columns");
        var cbGrids = pOpt.add("checkbox", undefined, "Scale baseline && document grids");
        var cbRedefine = pOpt.add("checkbox", undefined, "Bake all scaling to 100% when done (redefine scaling)");
        var cbZeroBleed = pOpt.add("checkbox", undefined, "Zero out document bleed values (avoid double bleed on export)");
        var cbOverset = pOpt.add("checkbox", undefined, "Check for overset text afterwards");
        cbStyles.value = true;
        cbParents.value = true;
        cbGuides.value = true;
        cbMargins.value = true;
        cbGrids.value = true;
        cbRedefine.value = true;
        cbZeroBleed.value = true;
        cbOverset.value = true;

        var cbCopy = pOpt.add("checkbox", undefined, "Run on a copy (Save a Copy next to the original, then open it)");
        cbCopy.value = false;

        // --- Buttons ---
        var gBtn = w.add("group");
        gBtn.alignment = "right";
        gBtn.add("button", undefined, "Cancel", { name: "cancel" });
        gBtn.add("button", undefined, "OK", { name: "ok" });

        if (w.show() !== 1) return null;

        var pct = parseFloat(etScale.text);
        var bleedMM = parseFloat(etBleed.text);
        if (isNaN(pct) || pct <= 0) { alert("Scale must be a positive number."); return null; }
        if (isNaN(bleedMM) || bleedMM < 0) { alert("Bleed must be zero or a positive number."); return null; }

        return {
            scale: pct / 100,
            pct: pct,
            bleedMM: bleedMM,
            bleedIsOriginal: rbOrig.value,
            doStyles: cbStyles.value,
            doParents: cbParents.value,
            doGuides: cbGuides.value,
            doMargins: cbMargins.value,
            doGrids: cbGrids.value,
            doRedefine: cbRedefine.value,
            doZeroBleed: cbZeroBleed.value,
            doOverset: cbOverset.value,
            onCopy: cbCopy.value
        };
    }

    // ---------------------------------------------------------------------
    // Helpers
    // ---------------------------------------------------------------------
    function isNum(v) { return typeof v === "number" && isFinite(v); }

    function scaleNumProp(obj, prop, s) {
        // Scales obj[prop] only if it currently holds a plain finite number.
        // Enum values (e.g. Leading.AUTO) and unset character-style
        // attributes are objects in ExtendScript, so they're skipped.
        try {
            var v = obj[prop];
            if (isNum(v)) obj[prop] = v * s;
        } catch (e) { /* property absent in this version — ignore */ }
    }

    function isGraphicClass(item) {
        var n = "";
        try { n = String(item.constructor.name); } catch (e) { return false; }
        return n === "Image" || n === "EPS" || n === "PDF" || n === "WMF" ||
               n === "PICT" || n === "ImportedPage" || n === "Graphic" ||
               n === "EPSText" || n === "Movie" || n === "Sound";
    }

    // ---------------------------------------------------------------------
    // Effects scaling (used for object-style definitions always, and as a
    // fallback for live items if adjustEffectsWhenScaling is unavailable)
    // ---------------------------------------------------------------------
    function scaleEffectsGroup(ts, s) {
        if (!ts) return;
        try { scaleNumProp(ts.dropShadowSettings, "distance", s); } catch (e) {}
        try { scaleNumProp(ts.dropShadowSettings, "size", s); } catch (e) {}
        try { scaleNumProp(ts.innerShadowSettings, "distance", s); } catch (e) {}
        try { scaleNumProp(ts.innerShadowSettings, "size", s); } catch (e) {}
        try { scaleNumProp(ts.outerGlowSettings, "size", s); } catch (e) {}
        try { scaleNumProp(ts.innerGlowSettings, "size", s); } catch (e) {}
        try { scaleNumProp(ts.bevelAndEmbossSettings, "size", s); } catch (e) {}
        try { scaleNumProp(ts.bevelAndEmbossSettings, "softness", s); } catch (e) {}
        try { scaleNumProp(ts.featherSettings, "width", s); } catch (e) {}
        try { scaleNumProp(ts.directionalFeatherSettings, "leftWidth", s); } catch (e) {}
        try { scaleNumProp(ts.directionalFeatherSettings, "rightWidth", s); } catch (e) {}
        try { scaleNumProp(ts.directionalFeatherSettings, "topWidth", s); } catch (e) {}
        try { scaleNumProp(ts.directionalFeatherSettings, "bottomWidth", s); } catch (e) {}
        try { scaleNumProp(ts.gradientFeatherSettings, "length", s); } catch (e) {}
    }

    function scaleEffectsOn(obj, s) {
        try { scaleEffectsGroup(obj.transparencySettings, s); } catch (e) {}
        try { scaleEffectsGroup(obj.fillTransparencySettings, s); } catch (e) {}
        try { scaleEffectsGroup(obj.strokeTransparencySettings, s); } catch (e) {}
        try { scaleEffectsGroup(obj.contentTransparencySettings, s); } catch (e) {}
    }

    // ---------------------------------------------------------------------
    // Style-definition scaling
    // Read all computed values first, THEN write, so based-on chains don't
    // get scaled twice (writing to a parent changes what a child reads).
    // ---------------------------------------------------------------------
    var TEXT_PROPS = [
        "pointSize", "leading", "baselineShift",
        "underlineOffset", "underlineWeight",
        "strikeThroughOffset", "strikeThroughWeight"
    ];
    var PARA_PROPS = TEXT_PROPS.concat([
        "spaceBefore", "spaceAfter",
        "firstLineIndent", "leftIndent", "rightIndent", "lastLineIndent",
        "ruleAboveOffset", "ruleAboveWeight", "ruleAboveLeftIndent", "ruleAboveRightIndent",
        "ruleBelowOffset", "ruleBelowWeight", "ruleBelowLeftIndent", "ruleBelowRightIndent",
        // Newer paragraph border/shading offsets — silently skipped on older versions
        "paragraphBorderTopOffset", "paragraphBorderBottomOffset",
        "paragraphBorderLeftOffset", "paragraphBorderRightOffset",
        "paragraphShadingTopOffset", "paragraphShadingBottomOffset",
        "paragraphShadingLeftOffset", "paragraphShadingRightOffset"
    ]);

    function snapshotProps(style, propList) {
        var snap = {};
        for (var i = 0; i < propList.length; i++) {
            try {
                var v = style[propList[i]];
                if (isNum(v)) snap[propList[i]] = v;
            } catch (e) {}
        }
        // Tab stops (paragraph styles only)
        try {
            var tl = style.tabList;
            if (tl && tl.length !== undefined && tl.length > 0) snap.__tabs = tl;
        } catch (e) {}
        return snap;
    }

    function applyProps(style, snap, s) {
        for (var k in snap) {
            if (k === "__tabs") continue;
            try { style[k] = snap[k] * s; } catch (e) {}
        }
        if (snap.__tabs) {
            try {
                var tabs = snap.__tabs;
                for (var t = 0; t < tabs.length; t++) {
                    if (isNum(tabs[t].position)) tabs[t].position = tabs[t].position * s;
                }
                style.tabList = tabs;
            } catch (e) {}
        }
    }

    function scaleStyleCollection(styles, propList, s, skipFirst) {
        var arr = [], snaps = [], i;
        for (i = 0; i < styles.length; i++) arr.push(styles[i]);
        // Pass 1: read everything (computed values)
        for (i = 0; i < arr.length; i++) {
            if (skipFirst && i === 0) { snaps.push(null); continue; } // "[No Paragraph Style]" / "[None]"
            snaps.push(snapshotProps(arr[i], propList));
        }
        // Pass 2: write scaled values
        var count = 0;
        for (i = 0; i < arr.length; i++) {
            if (!snaps[i]) continue;
            applyProps(arr[i], snaps[i], s);
            count++;
        }
        return count;
    }

    function scaleObjectStyles(doc, s) {
        var styles = doc.allObjectStyles, count = 0;
        for (var i = 0; i < styles.length; i++) {
            var os = styles[i];
            try {
                if (os.name === "[None]") continue;
            } catch (e) {}
            scaleNumProp(os, "strokeWeight", s);
            scaleNumProp(os, "cornerRadius", s);
            scaleNumProp(os, "topLeftCornerRadius", s);
            scaleNumProp(os, "topRightCornerRadius", s);
            scaleNumProp(os, "bottomLeftCornerRadius", s);
            scaleNumProp(os, "bottomRightCornerRadius", s);
            // Text frame insets held in the style definition
            try {
                var ins = os.insetSpacing;
                if (isNum(ins)) { os.insetSpacing = ins * s; }
                else if (ins && ins.length === 4) {
                    os.insetSpacing = [ins[0] * s, ins[1] * s, ins[2] * s, ins[3] * s];
                }
            } catch (e) {}
            // Text wrap offsets held in the style definition
            try {
                var two = os.textWrapPreferences.textWrapOffset;
                if (isNum(two)) { os.textWrapPreferences.textWrapOffset = two * s; }
                else if (two && two.length === 4) {
                    os.textWrapPreferences.textWrapOffset =
                        [two[0] * s, two[1] * s, two[2] * s, two[3] * s];
                }
            } catch (e) {}
            scaleEffectsOn(os, s);
            count++;
        }
        return count;
    }

    var TABLE_PROPS = [
        "spaceBefore", "spaceAfter",
        "topBorderStrokeWeight", "bottomBorderStrokeWeight",
        "leftBorderStrokeWeight", "rightBorderStrokeWeight"
    ];
    var CELL_PROPS = [
        "topInset", "bottomInset", "leftInset", "rightInset",
        "topEdgeStrokeWeight", "bottomEdgeStrokeWeight",
        "leftEdgeStrokeWeight", "rightEdgeStrokeWeight"
    ];

    // ---------------------------------------------------------------------
    // Spread engine (v1.1): measure-and-place.
    //   0. Record each top-level item's offset from the page's top-left.
    //   1. Scale each item IN PLACE about its own centre (Apply to Content
    //      handles text, strokes, effects).
    //   2. Resize the pages, then extend them for bleed.
    //   3. Measure where the page ACTUALLY ended up and translate every
    //      item to trimTL + (offset x scale). No assumptions about how
    //      InDesign repositions pages or renormalises coordinate spaces.
    //      Each placement is re-measured; a second corrective pass runs if
    //      needed, and the worst residual is reported.
    // ---------------------------------------------------------------------
    function pageTL(page) {
        return page.resolve(AnchorPoint.TOP_LEFT_ANCHOR,
                            CoordinateSpaces.PASTEBOARD_COORDINATES)[0];
    }

    function itemTL(item) {
        return item.resolve(AnchorPoint.TOP_LEFT_ANCHOR,
                            CoordinateSpaces.PASTEBOARD_COORDINATES)[0];
    }

    function processSpread(spread, s, bleedPt, cfg, report) {
        var pages = spread.pages, i, k;
        var oldRules = [];

        // Liquid Layout off so nothing "adjusts" behind our back
        for (i = 0; i < pages.length; i++) {
            try { oldRules.push(pages[i].layoutRule); pages[i].layoutRule = LayoutRuleOptions.OFF; }
            catch (e) { oldRules.push(null); }
        }

        var A = pageTL(pages[0]); // trim top-left BEFORE anything

        // ---- 0. Snapshot top-level items (incl. pasteboard items) and
        //         their offsets from the trim top-left ----
        var items = [];
        if (spread.pageItems.length > 0) {
            items = spread.pageItems.everyItem().getElements();
        }
        var offs = [];
        for (i = 0; i < items.length; i++) {
            try {
                var t = itemTL(items[i]);
                offs.push([t[0] - A[0], t[1] - A[1]]);
            } catch (e) { offs.push(null); }
        }

        // ---- 1. Scale each item in place, about its own centre ----
        if (s !== 1) {
            for (i = 0; i < items.length; i++) {
                try {
                    items[i].transform(
                        CoordinateSpaces.PASTEBOARD_COORDINATES,
                        AnchorPoint.CENTER_ANCHOR,
                        [s, 0, 0, s, 0, 0]
                    );
                } catch (e) {}
            }
        }

        // ---- 2. Scale the pages themselves ----
        if (s !== 1) {
            for (i = 0; i < pages.length; i++) {
                pages[i].resize(
                    CoordinateSpaces.INNER_COORDINATES,
                    AnchorPoint.TOP_LEFT_ANCHOR,
                    ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY,
                    [s, s]
                );
            }
        }

        // ---- 3. Scale margins / columns ----
        if (cfg.doMargins && s !== 1) {
            for (i = 0; i < pages.length; i++) {
                var mp = pages[i].marginPreferences;
                try {
                    mp.top *= s; mp.bottom *= s; mp.left *= s; mp.right *= s;
                    if (isNum(mp.columnGutter)) mp.columnGutter *= s;
                } catch (e) {}
            }
        }

        // ---- 4. Scale guides ----
        if (cfg.doGuides && s !== 1) {
            try {
                var gds = spread.guides;
                for (i = 0; i < gds.length; i++) {
                    try { if (isNum(gds[i].location)) gds[i].location = gds[i].location * s; }
                    catch (e) {}
                }
            } catch (e) {}
        }

        // ---- 5. Extend pages for bleed ----
        if (bleedPt > 0) {
            for (i = 0; i < pages.length; i++) {
                pages[i].resize(
                    CoordinateSpaces.INNER_COORDINATES,
                    AnchorPoint.CENTER_ANCHOR,
                    ResizeMethods.ADDING_CURRENT_DIMENSIONS_TO,
                    [2 * bleedPt, 2 * bleedPt]
                );
            }
            // Margins & guides are measured from the (new, larger) page
            // edge; push them in by the bleed so they stay on the trim.
            if (cfg.doMargins) {
                for (i = 0; i < pages.length; i++) {
                    var mp2 = pages[i].marginPreferences;
                    try {
                        mp2.top += bleedPt; mp2.bottom += bleedPt;
                        mp2.left += bleedPt; mp2.right += bleedPt;
                    } catch (e) {}
                }
            }
            if (cfg.doGuides) {
                try {
                    var gds2 = spread.guides;
                    for (i = 0; i < gds2.length; i++) {
                        try { if (isNum(gds2[i].location)) gds2[i].location = gds2[i].location + bleedPt; }
                        catch (e) {}
                    }
                } catch (e) {}
            }
        }

        // ---- 6. Place every item relative to where the page ACTUALLY is ----
        var N = pageTL(pages[0]);              // page (= bleed box) top-left now
        var trimX = N[0] + bleedPt;            // trim top-left now
        var trimY = N[1] + bleedPt;

        for (i = 0; i < items.length; i++) {
            if (!offs[i]) continue;
            var tgtX = trimX + s * offs[i][0];
            var tgtY = trimY + s * offs[i][1];
            try {
                // Up to two measure-translate passes, then verify
                for (k = 0; k < 2; k++) {
                    var c = itemTL(items[i]);
                    var dx = tgtX - c[0];
                    var dy = tgtY - c[1];
                    if (Math.abs(dx) <= EPS && Math.abs(dy) <= EPS) break;
                    items[i].transform(
                        CoordinateSpaces.PASTEBOARD_COORDINATES,
                        AnchorPoint.CENTER_ANCHOR,
                        [1, 0, 0, 1, dx, dy]
                    );
                }
                var v = itemTL(items[i]);
                var rx = Math.abs(v[0] - tgtX);
                var ry = Math.abs(v[1] - tgtY);
                var r = rx > ry ? rx : ry;
                if (r > report.maxResid) report.maxResid = r;
            } catch (e) {}
        }

        // Restore Liquid Layout rules
        for (i = 0; i < pages.length; i++) {
            try { if (oldRules[i] !== null) pages[i].layoutRule = oldRules[i]; } catch (e) {}
        }
    }

    // ---------------------------------------------------------------------
    // Main
    // ---------------------------------------------------------------------
    function main(doc, cfg) {

        var s = cfg.scale;
        var bleedPt = cfg.bleedMM * MM2PT;
        if (cfg.bleedIsOriginal) bleedPt = bleedPt * s; // entered at original size → scales down

        var summary = [];
        var savedPrefs = {};
        var lockedLayers = [];
        var lockedItems = [];
        var i, j;
        var b0;

        try {
            // ------- save & set app/doc prefs -------
            savedPrefs.measurementUnit = app.scriptPreferences.measurementUnit;
            app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

            // ------- record starting size (page 1), now guaranteed in points -------
            b0 = doc.pages[0].bounds; // [y1, x1, y2, x2]

            savedPrefs.enableRedraw = app.scriptPreferences.enableRedraw;
            app.scriptPreferences.enableRedraw = false;

            savedPrefs.whenScaling = app.transformPreferences.whenScaling;
            app.transformPreferences.whenScaling = WhenScalingOptions.APPLY_TO_CONTENT;

            savedPrefs.adjustStroke = app.transformPreferences.adjustStrokeWeightWhenScaling;
            app.transformPreferences.adjustStrokeWeightWhenScaling = true;

            var effectsPrefOK = true;
            try {
                savedPrefs.adjustEffects = app.transformPreferences.adjustEffectsWhenScaling;
                app.transformPreferences.adjustEffectsWhenScaling = true;
            } catch (e) { effectsPrefOK = false; }

            savedPrefs.rulerOrigin = doc.viewPreferences.rulerOrigin;
            doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
            savedPrefs.zeroPoint = doc.zeroPoint;
            doc.zeroPoint = [0, 0];

            // ------- unlock layers & items (remember what we changed) -------
            for (i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].locked) { lockedLayers.push(doc.layers[i]); doc.layers[i].locked = false; }
            }
            var unlockAll = function (items) {
                for (var k = 0; k < items.length; k++) {
                    try { if (items[k].locked) { lockedItems.push(items[k]); items[k].locked = false; } }
                    catch (e) {}
                }
            };
            unlockAll(doc.allPageItems);
            if (cfg.doParents) {
                for (i = 0; i < doc.masterSpreads.length; i++) {
                    try { unlockAll(doc.masterSpreads[i].allPageItems); } catch (e) {}
                }
            }

            // ------- spreads -------
            var report = { maxResid: 0 };
            for (i = 0; i < doc.spreads.length; i++) {
                processSpread(doc.spreads[i], s, bleedPt, cfg, report);
            }
            summary.push(doc.spreads.length + " spread(s) scaled to " + cfg.pct + "%.");

            if (cfg.doParents) {
                for (i = 0; i < doc.masterSpreads.length; i++) {
                    processSpread(doc.masterSpreads[i], s, bleedPt, cfg, report);
                }
                summary.push(doc.masterSpreads.length + " parent spread(s) scaled.");
            }

            summary.push(report.maxResid <= 0.05
                ? "Artwork placement verified (max drift " + report.maxResid.toFixed(3) + " pt)."
                : "WARNING: placement residual of " + report.maxResid.toFixed(2) +
                  " pt on at least one item — check positions.");

            if (bleedPt > 0) {
                summary.push("Pages extended by " + (bleedPt / MM2PT).toFixed(2) +
                             " mm bleed per side (page size now = trim + bleed).");
            }

            // ------- style definitions -------
            if (cfg.doStyles && s !== 1) {
                var nP = scaleStyleCollection(doc.allParagraphStyles, PARA_PROPS, s, true);
                var nC = scaleStyleCollection(doc.allCharacterStyles, TEXT_PROPS, s, true);
                var nO = scaleObjectStyles(doc, s);
                var nT = 0, nCe = 0;
                try { nT = scaleStyleCollection(doc.allTableStyles, TABLE_PROPS, s, true); } catch (e) {}
                try { nCe = scaleStyleCollection(doc.allCellStyles, CELL_PROPS, s, true); } catch (e) {}
                try { applyProps(doc.textDefaults, snapshotProps(doc.textDefaults, PARA_PROPS), s); } catch (e) {}
                summary.push("Styles scaled — para: " + nP + ", char: " + nC +
                             ", object: " + nO + ", table: " + nT + ", cell: " + nCe + ".");
            }

            // ------- grids -------
            if (cfg.doGrids && s !== 1) {
                try {
                    var gp = doc.gridPreferences;
                    if (isNum(gp.baselineDivision)) gp.baselineDivision = Math.max(0.1, gp.baselineDivision * s);
                    if (isNum(gp.baselineStart)) gp.baselineStart = gp.baselineStart * s;
                    if (isNum(gp.horizontalGridlineDivision)) gp.horizontalGridlineDivision *= s;
                    if (isNum(gp.verticalGridlineDivision)) gp.verticalGridlineDivision *= s;
                    summary.push("Baseline & document grids scaled.");
                } catch (e) {}
            }

            // ------- effects fallback (older versions only) -------
            if (!effectsPrefOK && s !== 1) {
                var allIt = doc.allPageItems;
                for (i = 0; i < allIt.length; i++) scaleEffectsOn(allIt[i], s);
                if (cfg.doParents) {
                    for (i = 0; i < doc.masterSpreads.length; i++) {
                        try {
                            var mIt = doc.masterSpreads[i].allPageItems;
                            for (j = 0; j < mIt.length; j++) scaleEffectsOn(mIt[j], s);
                        } catch (e) {}
                    }
                }
                summary.push("Effects scaled manually (adjustEffectsWhenScaling unavailable in this version).");
            }

            // ------- redefine scaling to 100% -------
            if (cfg.doRedefine) {
                var bake = function (items) {
                    var n = 0;
                    for (var k = 0; k < items.length; k++) {
                        if (isGraphicClass(items[k])) continue; // keep placed-image scale % meaningful
                        try { items[k].redefineScaling(); n++; } catch (e) {}
                    }
                    return n;
                };
                var baked = bake(doc.allPageItems);
                if (cfg.doParents) {
                    for (i = 0; i < doc.masterSpreads.length; i++) {
                        try { baked += bake(doc.masterSpreads[i].allPageItems); } catch (e) {}
                    }
                }
                summary.push("Scaling baked to 100% on " + baked + " item(s).");
            }

            // ------- zero document bleed -------
            if (cfg.doZeroBleed) {
                try {
                    doc.documentPreferences.documentBleedUniformSize = true;
                    doc.documentPreferences.documentBleedTopOffset = 0;
                    summary.push("Document bleed values zeroed.");
                } catch (e) {}
            }

            // ------- overset check -------
            if (cfg.doOverset) {
                var ovs = 0;
                try {
                    var stories = doc.stories.everyItem().getElements();
                    for (i = 0; i < stories.length; i++) {
                        try { if (stories[i].overflows) ovs++; } catch (e) {}
                    }
                } catch (e) {}
                summary.push(ovs === 0 ? "No overset stories." :
                             ("WARNING: " + ovs + " overset stor" + (ovs === 1 ? "y" : "ies") + " — check text frames."));
            }

            // ------- final size report -------
            var b1 = doc.pages[0].bounds;
            summary.push("Page 1: " +
                ((b0[3] - b0[1]) / MM2PT).toFixed(1) + " x " + ((b0[2] - b0[0]) / MM2PT).toFixed(1) + " mm  ->  " +
                ((b1[3] - b1[1]) / MM2PT).toFixed(1) + " x " + ((b1[2] - b1[0]) / MM2PT).toFixed(1) + " mm.");

            if (cfg.pct > 100) {
                summary.push("Note: scaling UP reduces effective PPI of placed images — check the Links panel.");
            }
            if (cfg.pct < 12) {
                summary.push("Note: extreme downscales can hit InDesign's 0.1 pt minimum type size — spot-check small text.");
            }

        } finally {
            // ------- restore locks -------
            for (i = 0; i < lockedItems.length; i++) { try { lockedItems[i].locked = true; } catch (e) {} }
            for (i = 0; i < lockedLayers.length; i++) { try { lockedLayers[i].locked = true; } catch (e) {} }

            // ------- restore prefs -------
            try { doc.zeroPoint = savedPrefs.zeroPoint; } catch (e) {}
            try { doc.viewPreferences.rulerOrigin = savedPrefs.rulerOrigin; } catch (e) {}
            try { app.transformPreferences.whenScaling = savedPrefs.whenScaling; } catch (e) {}
            try { app.transformPreferences.adjustStrokeWeightWhenScaling = savedPrefs.adjustStroke; } catch (e) {}
            try {
                if (savedPrefs.adjustEffects !== undefined)
                    app.transformPreferences.adjustEffectsWhenScaling = savedPrefs.adjustEffects;
            } catch (e) {}
            try { app.scriptPreferences.enableRedraw = savedPrefs.enableRedraw; } catch (e) {}
            try { app.scriptPreferences.measurementUnit = savedPrefs.measurementUnit; } catch (e) {}
        }

        alert("Scale Artwork — done\n\n" + summary.join("\n"));
    }

    // ---------------------------------------------------------------------
    // Run
    // ---------------------------------------------------------------------
    var cfg = showDialog();
    if (!cfg) return;

    var doc = app.activeDocument;

    // Pre-flight: multi-page spreads are unreliable for this operation
    var multi = false;
    for (var si = 0; si < doc.spreads.length; si++) {
        if (doc.spreads[si].pages.length > 1) { multi = true; break; }
    }
    if (multi) {
        if (!confirm("This document contains spreads with more than one page.\n" +
                     "Page-level scaling and bleed extension are only reliable on " +
                     "single-page spreads (typical artwork docs).\n\nContinue anyway?")) return;
    }

    // Optionally run on a copy
    if (cfg.onCopy) {
        if (!doc.saved) {
            alert("The document has never been saved, so a copy can't be created.\n" +
                  "Save it first, or untick 'Run on a copy'.");
            return;
        }
        var f = doc.fullName;
        var base = f.fsName.replace(/\.indd$/i, "");
        var copyFile = new File(base + "_" + cfg.pct + "pc.indd");
        var n = 1;
        while (copyFile.exists) { copyFile = new File(base + "_" + cfg.pct + "pc_" + (n++) + ".indd"); }
        doc.saveACopy(copyFile);
        doc = app.open(copyFile);
    }

    app.doScript(function () { main(doc, cfg); },
        ScriptLanguage.JAVASCRIPT, [],
        UndoModes.ENTIRE_SCRIPT, "Scale Artwork Document");

})();
