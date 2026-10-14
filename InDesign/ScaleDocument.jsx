/* =========================================================================
   ScaleDocument.jsx  —  v1.2
   Scale a large-format InDesign document to a new size, with optional
   bleed added by ENLARGING THE PAGE (trim + bleed becomes the page size).

   v1.2 — New size can be entered as a PERCENTAGE, a target WIDTH or a
          target HEIGHT (mm / in / pt). Width/height are measured against
          the ACTIVE page's trim; the dialog shows the current size and a
          live readout of the result. One uniform factor is applied to the
          whole document.
          Added a "verify & fix" pass: live values that InDesign's
          Apply-to-Content may leave alone (paragraph rules, underlines,
          strikethroughs, paragraph borders/shading, table & cell strokes,
          column widths/row heights, frame insets, column gutters,
          anchored-object offsets) are snapshotted before scaling and
          corrected afterwards if they did not change. Values InDesign
          already scaled are left alone, so nothing is scaled twice.
          Style-definition scaling now also covers paragraph borders and
          shading corner radii, span/split column gutters, hyphenation
          zone, object-style text frame options, anchored-object offsets
          and frame-fitting crops. Optional scaling of custom dashed /
          dotted stroke style definitions.
          FIX: the table-style pass skipped index 0, which is [Basic
          Table] (a real, usable style). Placeholder styles ([None],
          [No Paragraph Style]) are now skipped by NAME.
          A progress palette shows the current phase, a detail line and
          a bar while the script runs, so InDesign no longer looks hung.

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
   1. Scales every page and every page item (incl. pasteboard items) by
      one factor, anchored at each page's top-left. The factor comes from
      a percentage, or from a target width / height of the active page.
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
   5. Verifies live text, frames and tables and fixes any measured value
      that did not scale (see v1.2 note above).
   6. Runs redefineScaling() on all items afterwards so every frame
      reads 100% / 100% (placed graphics are skipped so their scale %
      stays meaningful in the Links panel).
   7. Optionally scales guides, margins/columns, and the baseline /
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
   - Target width/height refer to the TRIM of the active page. Bleed is
     added on top. Documents with mixed page sizes still get one uniform
     factor, derived from the active page.
   - Style-definition scaling writes computed values back to each style,
     which makes those attributes explicit (flattens based-on
     inheritance for the scaled attributes only). Values stay correct.
   - The verify & fix pass only touches a value when it is still exactly
     at its pre-scale size. Where it has to act on a paragraph / range it
     creates a local override (the same thing Apply to Content does).
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

    var UNITS = [
        { name: "mm", pt: 72 / 25.4 },
        { name: "in", pt: 72 },
        { name: "pt", pt: 1 }
    ];

    // ---------------------------------------------------------------------
    // Progress palette (non-modal). Every call is wrapped so a UI hiccup
    // can never abort the run.
    //   phase(from, to, label)  — new stage; bar band [from..to]; forced redraw
    //   step(frac, detail)      — bar position within the band (+ activity text)
    //   detail(text)            — activity text only (what is being touched now)
    //   Text and bar values are always set (cheap); the actual window repaint
    //   is throttled to ~10/s so that reporting never slows the run down.
    // ---------------------------------------------------------------------
    var PROG = (function () {
        var win = null, bar = null, stPhase = null, stDetail = null, stTime = null;
        var base = 0, span = 0;
        var t0 = 0, lastDraw = 0;
        var REDRAW_MS = 100;
        function safe(f) { try { f(); } catch (e) {} }
        function now() { return new Date().getTime(); }
        function elapsed() {
            var sec = Math.floor((now() - t0) / 1000);
            var m = Math.floor(sec / 60), s = sec % 60;
            return m + ":" + (s < 10 ? "0" : "") + s;
        }
        function draw(force) {
            if (!win) return;
            var t = now();
            if (!force && t - lastDraw < REDRAW_MS) return;
            lastDraw = t;
            stTime.text = elapsed();
            win.update();
        }
        return {
            open: function () {
                t0 = now();
                safe(function () {
                    win = new Window("palette", "Scale Artwork Document", undefined, { closeButton: false });
                    win.orientation = "column";
                    win.alignChildren = "fill";
                    win.margins = 15;
                    var row = win.add("group");
                    row.orientation = "row";
                    row.alignChildren = ["fill", "center"];
                    stPhase = row.add("statictext", undefined, "Starting...");
                    stPhase.characters = 64;
                    stTime = row.add("statictext", undefined, "0:00");
                    stTime.characters = 6;
                    stTime.justify = "right";
                    bar = win.add("progressbar", undefined, 0, 100);
                    bar.preferredSize = [520, 14];
                    stDetail = win.add("statictext", undefined, " ", { truncate: "middle" });
                    stDetail.characters = 76;
                    win.show();
                    draw(true);
                });
            },
            phase: function (from, to, label) {
                base = from; span = to - from;
                safe(function () {
                    stPhase.text = label;
                    stDetail.text = " ";
                    bar.value = from;
                    draw(true);
                });
            },
            step: function (frac, detail) {
                safe(function () {
                    if (frac < 0) frac = 0;
                    if (frac > 1) frac = 1;
                    bar.value = base + span * frac;
                    if (detail !== undefined) stDetail.text = detail;
                    draw(false);
                });
            },
            detail: function (text) {
                safe(function () {
                    stDetail.text = text;
                    draw(false);
                });
            },
            close: function () {
                safe(function () { if (win) win.close(); win = null; });
            }
        };
    })();

    // ---------------------------------------------------------------------
    // Small helpers (function declarations are hoisted, so these are usable
    // from the dialog code below)
    // ---------------------------------------------------------------------
    function isNum(v) { return typeof v === "number" && isFinite(v); }

    function fmt(n) {
        // Up to 3 decimals, trailing zeros removed: 420 / 37.5 / 0.118
        if (!isNum(n)) return "?";
        var str = n.toFixed(3);
        str = str.replace(/0+$/, "").replace(/\.$/, "");
        if (str === "-0") str = "0";
        return str;
    }

    function getReferencePage(doc) {
        // The active page of the layout window (usually page 1); falls
        // back to page 1 when there is no layout window.
        try {
            var win = app.activeWindow;
            if (win && win instanceof LayoutWindow) {
                var p = win.activePage;
                if (p && p.isValid) return p;
            }
        } catch (e) {}
        return doc.pages[0];
    }

    function pageSizePt(page) {
        var saved = app.scriptPreferences.measurementUnit;
        app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
        var b;
        try { b = page.bounds; } // [y1, x1, y2, x2]
        finally { app.scriptPreferences.measurementUnit = saved; }
        return { w: b[3] - b[1], h: b[2] - b[0] };
    }

    // ---------------------------------------------------------------------
    // Dialog
    // ---------------------------------------------------------------------
    function showDialog(doc) {
        var refPage = getReferencePage(doc);
        var cur = pageSizePt(refPage);
        var refIndex = 0;
        try { refIndex = refPage.documentOffset; } catch (e) {}

        var w = new Window("dialog", "Scale Artwork Document");
        w.orientation = "column";
        w.alignChildren = "fill";

        // --- Size ---
        var pSize = w.add("panel", undefined, "New size");
        pSize.orientation = "column";
        pSize.alignChildren = "left";
        pSize.margins = 15;

        var gMode = pSize.add("group");
        gMode.alignChildren = "center";
        gMode.add("statictext", undefined, "Set by:");
        var ddMode = gMode.add("dropdownlist", undefined,
            ["Percent of original", "Target width", "Target height"]);
        ddMode.selection = 0;
        var etVal = gMode.add("edittext", undefined, "50");
        etVal.characters = 9;
        var stAfter = gMode.add("statictext", undefined, "%");
        stAfter.characters = 3;
        gMode.add("statictext", undefined, "   Units:");
        var ddUnit = gMode.add("dropdownlist", undefined, ["mm", "in", "pt"]);
        ddUnit.selection = 0;

        var stCur = pSize.add("statictext", undefined, "");
        stCur.characters = 78;
        var stNew = pSize.add("statictext", undefined, "");
        stNew.characters = 78;

        // --- Bleed ---
        var pBleed = w.add("panel", undefined, "Bleed (extends the page — document bleed settings are NOT used)");
        pBleed.orientation = "column";
        pBleed.alignChildren = "left";
        pBleed.margins = 15;
        var gB = pBleed.add("group");
        gB.add("statictext", undefined, "Bleed each side:");
        var etBleed = gB.add("edittext", undefined, "0");
        etBleed.characters = 7;
        var stBleedUnit = gB.add("statictext", undefined, "mm");
        stBleedUnit.characters = 3;
        var rbOrig = pBleed.add("radiobutton", undefined, "Value is at ORIGINAL size (scales down with the artwork)");
        var rbFinal = pBleed.add("radiobutton", undefined, "Value is FINAL size (applied as-is after scaling)");
        rbOrig.value = true;

        // --- Options ---
        var pOpt = w.add("panel", undefined, "Options");
        pOpt.orientation = "column";
        pOpt.alignChildren = "left";
        pOpt.margins = 15;
        var cbStyles = pOpt.add("checkbox", undefined, "Scale style definitions (paragraph / character / object / table / cell)");
        var cbLiveFix = pOpt.add("checkbox", undefined, "Verify live text && tables after scaling: fix rules, underlines, borders, insets, cell strokes left unscaled");
        var cbStrokeStyles = pOpt.add("checkbox", undefined, "Scale custom dashed / dotted stroke style definitions (dash && gap lengths)");
        var cbParents = pOpt.add("checkbox", undefined, "Include parent (master) pages");
        var cbGuides = pOpt.add("checkbox", undefined, "Scale ruler guides");
        var cbMargins = pOpt.add("checkbox", undefined, "Scale margins && columns");
        var cbGrids = pOpt.add("checkbox", undefined, "Scale baseline && document grids");
        var cbRedefine = pOpt.add("checkbox", undefined, "Bake all scaling to 100% when done (redefine scaling)");
        var cbZeroBleed = pOpt.add("checkbox", undefined, "Zero out document bleed values (avoid double bleed on export)");
        var cbOverset = pOpt.add("checkbox", undefined, "Check for overset text afterwards");
        cbStyles.value = true;
        cbLiveFix.value = true;
        cbStrokeStyles.value = true;
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

        // --- Live behaviour ---
        function unit() { return UNITS[ddUnit.selection.index]; }
        function modeIdx() { return ddMode.selection.index; }

        function currentScale() {
            var v = parseFloat(etVal.text);
            if (isNaN(v) || v <= 0) return NaN;
            var m = modeIdx();
            if (m === 0) return v / 100;
            var pt = v * unit().pt;
            return m === 1 ? pt / cur.w : pt / cur.h;
        }

        function refresh() {
            var u = unit();
            var isPct = modeIdx() === 0;
            stAfter.text = isPct ? "%" : u.name;
            stBleedUnit.text = u.name;
            stCur.text = "Reference: page " + (refIndex + 1) + " (active page), trim " +
                         fmt(cur.w / u.pt) + " x " + fmt(cur.h / u.pt) + " " + u.name;
            var s = currentScale();
            if (!isNum(s)) {
                stNew.text = "Result: enter a positive number.";
            } else {
                stNew.text = "Result: " + fmt(cur.w * s / u.pt) + " x " + fmt(cur.h * s / u.pt) +
                             " " + u.name + " trim  =  " + fmt(s * 100) + " %   (bleed is added on top)";
            }
        }

        ddMode.onChange = function () {
            var u = unit();
            var m = modeIdx();
            if (m === 0) etVal.text = "50";
            else if (m === 1) etVal.text = fmt(cur.w / u.pt);
            else etVal.text = fmt(cur.h / u.pt);
            refresh();
        };

        var prevUnitPt = unit().pt;
        ddUnit.onChange = function () {
            var newPt = unit().pt;
            if (newPt !== prevUnitPt) {
                if (modeIdx() !== 0) {
                    var v = parseFloat(etVal.text);
                    if (!isNaN(v)) etVal.text = fmt(v * prevUnitPt / newPt);
                }
                var b = parseFloat(etBleed.text);
                if (!isNaN(b)) etBleed.text = fmt(b * prevUnitPt / newPt);
                prevUnitPt = newPt;
            }
            refresh();
        };

        etVal.onChanging = refresh;
        etVal.onChange = refresh;
        refresh();

        if (w.show() !== 1) return null;

        var scale = currentScale();
        var bleedVal = parseFloat(etBleed.text);
        if (!isNum(scale) || scale <= 0) { alert("New size must be a positive number."); return null; }
        if (isNaN(bleedVal) || bleedVal < 0) { alert("Bleed must be zero or a positive number."); return null; }

        return {
            scale: scale,
            pct: scale * 100,
            mode: modeIdx(),
            unitName: unit().name,
            unitPt: unit().pt,
            refPageIndex: refIndex,
            bleedPt: bleedVal * unit().pt,  // before the original/final adjustment
            bleedIsOriginal: rbOrig.value,
            doStyles: cbStyles.value,
            doLiveFix: cbLiveFix.value,
            doStrokeStyles: cbStrokeStyles.value,
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
    function scaleNumProp(obj, prop, s) {
        // Scales obj[prop] only if it currently holds a plain finite number.
        // Enum values (e.g. Leading.AUTO) and unset character-style
        // attributes are objects in ExtendScript, so they're skipped.
        try {
            var v = obj[prop];
            if (isNum(v)) obj[prop] = v * s;
        } catch (e) { /* property absent in this version — ignore */ }
    }

    function scaleInsetProp(obj, prop, s) {
        // For properties that may be a single number or a 4-value array
        try {
            var v = obj[prop];
            if (isNum(v)) { obj[prop] = v * s; }
            else if (v && v.length === 4) {
                obj[prop] = [v[0] * s, v[1] * s, v[2] * s, v[3] * s];
            }
        } catch (e) {}
    }

    function isGraphicClass(item) {
        var n = "";
        try { n = String(item.constructor.name); } catch (e) { return false; }
        return n === "Image" || n === "EPS" || n === "PDF" || n === "WMF" ||
               n === "PICT" || n === "ImportedPage" || n === "Graphic" ||
               n === "EPSText" || n === "Movie" || n === "Sound";
    }

    function className(obj) {
        try { return String(obj.constructor.name); } catch (e) { return ""; }
    }

    function isPlaceholderStyle(st) {
        // [None], [No Paragraph Style], [No Character Style] ... but NOT
        // [Basic Paragraph] / [Basic Table], which are real, editable styles.
        var n = "";
        try { n = String(st.name); } catch (e) { return false; }
        return /^\[(None|No [^\]]*)\]$/.test(n);
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
    // Property names that don't exist in this InDesign version are
    // silently skipped.
    // ---------------------------------------------------------------------
    var TEXT_PROPS = [
        "pointSize", "leading", "baselineShift",
        "underlineOffset", "underlineWeight",
        "strikeThroughOffset", "strikeThroughWeight"
    ];
    var PARA_PROPS = TEXT_PROPS.concat([
        "spaceBefore", "spaceAfter",
        "firstLineIndent", "leftIndent", "rightIndent", "lastLineIndent",
        "hyphenationZone",
        "ruleAboveOffset", "ruleAboveWeight", "ruleAboveLeftIndent", "ruleAboveRightIndent",
        "ruleBelowOffset", "ruleBelowWeight", "ruleBelowLeftIndent", "ruleBelowRightIndent",
        // Span / split columns (CS5+)
        "spanColumnInsideGutter", "spanColumnOutsideGutter",
        "spanColumnMinSpaceBefore", "spanColumnMinSpaceAfter",
        "splitColumnInsideGutter", "splitColumnOutsideGutter",
        // Paragraph borders & shading (CC 2018+)
        "paragraphBorderTopLineWeight", "paragraphBorderBottomLineWeight",
        "paragraphBorderLeftLineWeight", "paragraphBorderRightLineWeight",
        "paragraphBorderTopOffset", "paragraphBorderBottomOffset",
        "paragraphBorderLeftOffset", "paragraphBorderRightOffset",
        "paragraphBorderTopLeftCornerRadius", "paragraphBorderTopRightCornerRadius",
        "paragraphBorderBottomLeftCornerRadius", "paragraphBorderBottomRightCornerRadius",
        "paragraphShadingTopOffset", "paragraphShadingBottomOffset",
        "paragraphShadingLeftOffset", "paragraphShadingRightOffset",
        "paragraphShadingTopLeftCornerRadius", "paragraphShadingTopRightCornerRadius",
        "paragraphShadingBottomLeftCornerRadius", "paragraphShadingBottomRightCornerRadius"
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

    function styleName(st) {
        try { return String(st.name); } catch (e) { return "?"; }
    }

    function scaleStyleCollection(styles, propList, s, label, f0, f1) {
        // label / f0 / f1: progress text and bar band for this collection
        var arr = [], snaps = [], i, n;
        if (label === undefined) label = "Style";
        if (f0 === undefined) { f0 = 0; f1 = 1; }
        for (i = 0; i < styles.length; i++) arr.push(styles[i]);
        n = arr.length;
        var span = (f1 - f0) / 2;
        // Pass 1: read everything (computed values)
        for (i = 0; i < n; i++) {
            if (isPlaceholderStyle(arr[i])) { snaps.push(null); continue; }
            PROG.step(f0 + span * i / n, label + " " + (i + 1) + "/" + n + ": " + styleName(arr[i]) + " — reading");
            snaps.push(snapshotProps(arr[i], propList));
        }
        // Pass 2: write scaled values
        var count = 0;
        for (i = 0; i < n; i++) {
            if (!snaps[i]) continue;
            PROG.step(f0 + span + span * i / n, label + " " + (i + 1) + "/" + n + ": " + styleName(arr[i]) + " — writing");
            applyProps(arr[i], snaps[i], s);
            count++;
        }
        return count;
    }

    var FRAME_PREF_PROPS = [
        "textColumnGutter", "textColumnFixedWidth", "minimumFirstBaselineOffset",
        // Column rules (CC 2020+) — names vary by version; absent ones are skipped
        "columnRuleWidth", "columnRuleStrokeWeight", "columnRuleOffset",
        "columnRuleInsetStart", "columnRuleInsetEnd"
    ];
    var BASELINE_GRID_PROPS = ["baselineFrameGridIncrement", "startingOffsetForBaselineFrameGrid"];
    var ANCHOR_PROPS = ["anchorXoffset", "anchorYoffset", "anchorSpaceAbove"];
    var FIT_PROPS = ["topCrop", "bottomCrop", "leftCrop", "rightCrop"];

    function scaleObjectStyles(doc, s, f0, f1) {
        var styles = doc.allObjectStyles, count = 0, k, n = styles.length;
        if (f0 === undefined) { f0 = 0; f1 = 1; }
        for (var i = 0; i < n; i++) {
            var os = styles[i];
            if (isPlaceholderStyle(os)) continue;
            var tag = "Object style " + (i + 1) + "/" + n + ": " + styleName(os) + " — ";
            PROG.step(f0 + (f1 - f0) * i / n, tag + "stroke weight & corner radii");
            scaleNumProp(os, "strokeWeight", s);
            scaleNumProp(os, "cornerRadius", s);
            scaleNumProp(os, "topLeftCornerRadius", s);
            scaleNumProp(os, "topRightCornerRadius", s);
            scaleNumProp(os, "bottomLeftCornerRadius", s);
            scaleNumProp(os, "bottomRightCornerRadius", s);

            // Text frame options held in the style definition. Prefer the
            // textFramePreferences sub-object; fall back to a direct
            // insetSpacing property on older versions. Never both.
            var viaPrefs = false;
            PROG.detail(tag + "text frame options (insets, gutter, baseline)");
            try {
                var tfp = os.textFramePreferences;
                if (tfp) {
                    scaleInsetProp(tfp, "insetSpacing", s);
                    for (k = 0; k < FRAME_PREF_PROPS.length; k++) scaleNumProp(tfp, FRAME_PREF_PROPS[k], s);
                    viaPrefs = true;
                }
            } catch (e) {}
            if (!viaPrefs) scaleInsetProp(os, "insetSpacing", s);

            try {
                var bg = os.baselineFrameGridOptions;
                for (k = 0; k < BASELINE_GRID_PROPS.length; k++) scaleNumProp(bg, BASELINE_GRID_PROPS[k], s);
            } catch (e) {}

            // Text wrap offsets held in the style definition
            PROG.detail(tag + "text wrap offset");
            try { scaleInsetProp(os.textWrapPreferences, "textWrapOffset", s); } catch (e) {}

            // Anchored object offsets
            PROG.detail(tag + "anchored object offsets");
            try {
                var ao = os.anchoredObjectSettings;
                for (k = 0; k < ANCHOR_PROPS.length; k++) scaleNumProp(ao, ANCHOR_PROPS[k], s);
            } catch (e) {}

            // Frame fitting crop amounts
            PROG.detail(tag + "frame fitting crops");
            try {
                var ff = os.frameFittingOptions;
                for (k = 0; k < FIT_PROPS.length; k++) scaleNumProp(ff, FIT_PROPS[k], s);
            } catch (e) {}

            PROG.detail(tag + "effects (shadows, glows, feathers)");
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
    // Custom dashed / dotted stroke style definitions.
    // Custom styles store dash / gap lengths as absolute points, so they
    // don't follow the stroke weight. InDesign's built-in ones are relative
    // to the weight and read-only; they're skipped by name and by try/catch.
    // ---------------------------------------------------------------------
    var BUILTIN_STROKE_NAMES = {
        "Dashed (3 and 2)": 1, "Dashed (4 and 4)": 1, "Dashed": 1,
        "Dotted": 1, "Japanese Dots": 1
    };

    function scaleStrokeStyles(doc, s) {
        var n = 0, i;
        var scaleArr = function (st, prop) {
            try {
                PROG.detail("Stroke style: " + styleName(st) + " — " + prop);
                if (BUILTIN_STROKE_NAMES[String(st.name)]) return false;
                var a = st[prop];
                if (!(a instanceof Array) || a.length === 0) return false;
                var out = [];
                for (var j = 0; j < a.length; j++) out.push(isNum(a[j]) ? a[j] * s : a[j]);
                st[prop] = out;
                return true;
            } catch (e) { return false; }
        };
        try {
            var ds = doc.dashedStrokeStyles;
            for (i = 0; i < ds.length; i++) if (scaleArr(ds[i], "dashArray")) n++;
        } catch (e) {}
        try {
            var dt = doc.dottedStrokeStyles;
            for (i = 0; i < dt.length; i++) if (scaleArr(dt[i], "dotArray")) n++;
        } catch (e) {}
        return n;
    }

    // ---------------------------------------------------------------------
    // Live "verify & fix" engine.
    //   Before scaling, snapshot measured values on live paragraphs, style
    //   ranges, tables, cells, text frames and anchored objects. After
    //   scaling (items + style definitions), re-read each value:
    //     - unchanged (still == pre)      -> InDesign left it; set pre * s
    //     - already == pre * s            -> fine, count as OK
    //     - anything else                 -> leave it, count as "other"
    //   A value is only ever written when it is provably unscaled, so this
    //   can't double-scale. Zero values are skipped (0 * s = 0).
    // ---------------------------------------------------------------------
    var LIVE_PARA_PROPS = [
        "spaceBefore", "spaceAfter",
        "firstLineIndent", "leftIndent", "rightIndent", "lastLineIndent",
        "hyphenationZone",
        "ruleAboveWeight", "ruleAboveOffset", "ruleAboveLeftIndent", "ruleAboveRightIndent",
        "ruleBelowWeight", "ruleBelowOffset", "ruleBelowLeftIndent", "ruleBelowRightIndent",
        "paragraphBorderTopLineWeight", "paragraphBorderBottomLineWeight",
        "paragraphBorderLeftLineWeight", "paragraphBorderRightLineWeight",
        "paragraphBorderTopOffset", "paragraphBorderBottomOffset",
        "paragraphBorderLeftOffset", "paragraphBorderRightOffset",
        "paragraphShadingTopOffset", "paragraphShadingBottomOffset",
        "paragraphShadingLeftOffset", "paragraphShadingRightOffset"
    ];
    var LIVE_CHAR_PROPS = [
        "pointSize", "leading", "baselineShift",
        "underlineWeight", "underlineOffset",
        "strikeThroughWeight", "strikeThroughOffset"
    ];
    var LIVE_TABLE_PROPS = TABLE_PROPS;
    var LIVE_CELL_PROPS = CELL_PROPS;
    var LIVE_FRAME_PROPS = ["insetSpacing"].concat(FRAME_PREF_PROPS);

    var TOL = 0.002; // pt; InDesign rounds stored values to ~0.001

    function snapVal(v) {
        if (isNum(v)) return v;
        if (v && v.length === 4 && isNum(v[0]) && isNum(v[1]) && isNum(v[2]) && isNum(v[3])) {
            return [v[0], v[1], v[2], v[3]];
        }
        return null; // enum / mixed / absent
    }

    function near(a, b) { return Math.abs(a - b) <= TOL; }
    function arrNear(a, b) {
        if (!a || !b || a.length !== b.length) return false;
        for (var i = 0; i < a.length; i++) if (!near(a[i], b[i])) return false;
        return true;
    }
    function mulArr(a, s) { return [a[0] * s, a[1] * s, a[2] * s, a[3] * s]; }
    function allZero(a) { return a[0] === 0 && a[1] === 0 && a[2] === 0 && a[3] === 0; }

    function readBulk(spec, objs, prop) {
        // Reads prop for every object. Uses an everyItem() specifier when
        // available (one call), falling back to per-object reads.
        var n = objs.length, a = null, i;
        if (spec) {
            try {
                a = spec[prop];
                if (n === 1) a = [a];
                if (!(a instanceof Array) || a.length !== n) a = null;
            } catch (e) { a = null; }
        }
        if (!a) {
            a = [];
            for (i = 0; i < n; i++) {
                try { a.push(objs[i][prop]); } catch (e) { a.push(null); }
            }
        }
        return a;
    }

    function snapObjects(objs, spec, props, label) {
        var snap = { objs: objs, spec: spec, vals: {}, label: label || "" };
        for (var p = 0; p < props.length; p++) {
            PROG.detail(snap.label + " (" + objs.length + ") — " + props[p]);
            var raw = readBulk(spec, objs, props[p]);
            var arr = [], any = false;
            for (var i = 0; i < raw.length; i++) {
                var v = snapVal(raw[i]);
                if (v !== null) any = true;
                arr.push(v);
            }
            if (any) snap.vals[props[p]] = arr;
        }
        return snap;
    }

    function verifySnap(snap, s, counts) {
        var objs = snap.objs, n = objs.length;
        for (var prop in snap.vals) {
            PROG.detail(snap.label + " (" + n + ") — " + prop + "   [fixed so far: " + counts.fixed + "]");
            var pre = snap.vals[prop];
            var cur = readBulk(snap.spec, objs, prop);
            for (var i = 0; i < n; i++) {
                var p = pre[i];
                if (p === null) continue;
                var c = snapVal(cur[i]);
                if (c === null) continue;
                if (isNum(p)) {
                    if (p === 0) continue;
                    if (Math.abs(p * s - p) <= 2 * TOL) continue;   // can't tell scaled from unscaled
                    if (near(c, p * s)) { counts.ok++; continue; }
                    if (near(c, p)) {
                        try { objs[i][prop] = p * s; counts.fixed++; }
                        catch (e) { counts.failed++; }
                    } else { counts.other++; }
                } else {
                    if (allZero(p)) continue;
                    var ps = mulArr(p, s);
                    if (arrNear(c, ps)) { counts.ok++; continue; }
                    if (arrNear(c, p)) {
                        try { objs[i][prop] = ps; counts.fixed++; }
                        catch (e) { counts.failed++; }
                    } else { counts.other++; }
                }
            }
        }
    }

    function takeLiveSnapshot(doc, cfg) {
        var entries = [];
        var i, t, c;
        var add = function (objs, spec, props, label) {
            if (objs && objs.length > 0) entries.push(snapObjects(objs, spec, props, label));
        };

        // ---- stories: paragraphs, style ranges, tables ----
        var stories = [];
        try { stories = doc.stories.everyItem().getElements(); } catch (e) {}
        var nSt = stories.length;
        for (i = 0; i < nSt; i++) {
            var st = stories[i];
            var sLab = "Story " + (i + 1) + "/" + nSt;
            PROG.step(0.8 * i / nSt, sLab);
            try { add(st.paragraphs.everyItem().getElements(), st.paragraphs.everyItem(), LIVE_PARA_PROPS, sLab + " paragraphs"); } catch (e) {}
            try { add(st.textStyleRanges.everyItem().getElements(), st.textStyleRanges.everyItem(), LIVE_CHAR_PROPS, sLab + " text ranges"); } catch (e) {}

            var tables = [];
            try { tables = st.tables.everyItem().getElements(); } catch (e) {}
            for (t = 0; t < tables.length; t++) {
                var tb = tables[t];
                var tLab = sLab + " table " + (t + 1) + "/" + tables.length;
                add([tb], null, LIVE_TABLE_PROPS, tLab);
                try { add(tb.columns.everyItem().getElements(), tb.columns.everyItem(), ["width"], tLab + " columns"); } catch (e) {}
                try { add(tb.rows.everyItem().getElements(), tb.rows.everyItem(), ["minimumHeight", "height"], tLab + " rows"); } catch (e) {}
                var cells = [];
                try {
                    cells = tb.cells.everyItem().getElements();
                    add(cells, tb.cells.everyItem(), LIVE_CELL_PROPS, tLab + " cells");
                } catch (e) {}
                for (c = 0; c < cells.length; c++) {
                    var cLab = tLab + " cell " + (c + 1) + "/" + cells.length;
                    try { add(cells[c].paragraphs.everyItem().getElements(), cells[c].paragraphs.everyItem(), LIVE_PARA_PROPS, cLab + " paragraphs"); } catch (e) {}
                    try { add(cells[c].textStyleRanges.everyItem().getElements(), cells[c].textStyleRanges.everyItem(), LIVE_CHAR_PROPS, cLab + " text ranges"); } catch (e) {}
                }
            }
        }

        // ---- page items: text frame prefs, frame baseline grids, anchored offsets ----
        PROG.step(0.82, "Collecting page items...");
        var items = [];
        try { items = doc.allPageItems; } catch (e) {}
        if (cfg.doParents) {
            for (i = 0; i < doc.masterSpreads.length; i++) {
                try { items = items.concat(doc.masterSpreads[i].allPageItems); } catch (e) {}
            }
        }
        var framePrefs = [], gridOpts = [], anchors = [];
        var nIt = items.length;
        for (i = 0; i < nIt; i++) {
            var it = items[i];
            if (i % 50 === 0) {
                PROG.step(0.82 + 0.1 * i / nIt, "Scanning page item " + (i + 1) + "/" + nIt + " (" + className(it) + ")");
            }
            try {
                if (className(it) === "TextFrame") {
                    framePrefs.push(it.textFramePreferences);
                    try { gridOpts.push(it.baselineFrameGridOptions); } catch (e1) {}
                }
            } catch (e) {}
            try {
                if (className(it.parent) === "Character") anchors.push(it.anchoredObjectSettings);
            } catch (e) {}
        }
        PROG.step(0.92, "Text frames: " + framePrefs.length + ", anchored objects: " + anchors.length);
        add(framePrefs, null, LIVE_FRAME_PROPS, "Text frames");
        add(gridOpts, null, BASELINE_GRID_PROPS, "Frame baseline grids");
        add(anchors, null, ANCHOR_PROPS, "Anchored objects");

        return entries;
    }

    function applyLiveFixes(entries, s) {
        var counts = { fixed: 0, ok: 0, other: 0, failed: 0 };
        for (var i = 0; i < entries.length; i++) {
            PROG.step(i / entries.length, entries[i].label);
            try { verifySnap(entries[i], s, counts); } catch (e) {}
        }
        return counts;
    }

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

    function spreadProg(report, f, text) {
        // Progress within this spread's slice of the bar
        PROG.step(report.pf0 + (report.pf1 - report.pf0) * f, report.label + " — " + text);
    }

    function processSpread(spread, s, bleedPt, cfg, report) {
        var pages = spread.pages, i, k;
        var oldRules = [];
        if (!report.label) { report.label = "Spread"; report.pf0 = 0; report.pf1 = 1; }

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
        spreadProg(report, 0, "scaling " + items.length + " item(s)");
        if (s !== 1) {
            for (i = 0; i < items.length; i++) {
                if (i % 10 === 0 && i > 0) {
                    spreadProg(report, 0.45 * i / items.length,
                               "scaling item " + i + "/" + items.length + " (" + className(items[i]) + ")");
                }
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
        spreadProg(report, 0.45, "resizing page(s)" + (bleedPt > 0 ? " and adding bleed" : ""));
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

        spreadProg(report, 0.55, "placing " + items.length + " item(s)");
        for (i = 0; i < items.length; i++) {
            if (i % 10 === 0 && i > 0) {
                spreadProg(report, 0.55 + 0.45 * i / items.length,
                           "placing item " + i + "/" + items.length + " (" + className(items[i]) + ")");
            }
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
        var bleedPt = cfg.bleedPt;
        if (cfg.bleedIsOriginal) bleedPt = bleedPt * s; // entered at original size → scales down
        var U = cfg.unitPt, UN = cfg.unitName;

        var summary = [];
        var savedPrefs = {};
        var lockedLayers = [];
        var lockedItems = [];
        var i, j;
        var b0, refPage;

        try {
            PROG.phase(0, 3, "Preparing: setting preferences, unlocking layers and items...");

            // ------- save & set app/doc prefs -------
            savedPrefs.measurementUnit = app.scriptPreferences.measurementUnit;
            app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

            // ------- record starting size (reference page), now guaranteed in points -------
            refPage = (cfg.refPageIndex < doc.pages.length) ? doc.pages[cfg.refPageIndex] : doc.pages[0];
            b0 = refPage.bounds; // [y1, x1, y2, x2]

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

            // ------- live snapshot (before anything moves; units are points) -------
            var live = null;
            if (cfg.doLiveFix && s !== 1) {
                PROG.phase(3, 18, "Recording live text, frame and table values (before scaling)...");
                try { live = takeLiveSnapshot(doc, cfg); }
                catch (e) { live = null; summary.push("Live verify skipped (snapshot failed: " + e + ")."); }
            }

            // ------- spreads -------
            var report = { maxResid: 0 };
            var nSp = doc.spreads.length;
            PROG.phase(18, 55, "Scaling pages and artwork (" + nSp + " spread" + (nSp === 1 ? "" : "s") + ")...");
            for (i = 0; i < nSp; i++) {
                report.label = "Spread " + (i + 1) + " of " + nSp;
                report.pf0 = i / nSp;
                report.pf1 = (i + 1) / nSp;
                processSpread(doc.spreads[i], s, bleedPt, cfg, report);
            }
            summary.push(nSp + " spread(s) scaled to " + fmt(cfg.pct) + "%.");

            if (cfg.doParents) {
                var nMs = doc.masterSpreads.length;
                PROG.phase(55, 60, "Scaling parent (master) spreads (" + nMs + ")...");
                for (i = 0; i < nMs; i++) {
                    report.label = "Parent spread " + (i + 1) + " of " + nMs;
                    report.pf0 = i / nMs;
                    report.pf1 = (i + 1) / nMs;
                    processSpread(doc.masterSpreads[i], s, bleedPt, cfg, report);
                }
                summary.push(nMs + " parent spread(s) scaled.");
            }

            summary.push(report.maxResid <= 0.05
                ? "Artwork placement verified (max drift " + report.maxResid.toFixed(3) + " pt)."
                : "WARNING: placement residual of " + report.maxResid.toFixed(2) +
                  " pt on at least one item — check positions.");

            if (bleedPt > 0) {
                summary.push("Pages extended by " + fmt(bleedPt / U) + " " + UN +
                             " bleed per side (page size now = trim + bleed).");
            }

            // ------- style definitions -------
            if (cfg.doStyles && s !== 1) {
                PROG.phase(60, 70, "Scaling style definitions...");
                var nP = scaleStyleCollection(doc.allParagraphStyles, PARA_PROPS, s, "Paragraph style", 0.0, 0.3);
                var nC = scaleStyleCollection(doc.allCharacterStyles, TEXT_PROPS, s, "Character style", 0.3, 0.45);
                var nO = scaleObjectStyles(doc, s, 0.45, 0.8);
                var nT = 0, nCe = 0;
                try { nT = scaleStyleCollection(doc.allTableStyles, TABLE_PROPS, s, "Table style", 0.8, 0.88); } catch (e) {}
                try { nCe = scaleStyleCollection(doc.allCellStyles, CELL_PROPS, s, "Cell style", 0.88, 0.95); } catch (e) {}
                PROG.step(0.95, "Text defaults");
                try { applyProps(doc.textDefaults, snapshotProps(doc.textDefaults, PARA_PROPS), s); } catch (e) {}
                summary.push("Styles scaled — para: " + nP + ", char: " + nC +
                             ", object: " + nO + ", table: " + nT + ", cell: " + nCe + ".");
            }

            // ------- custom stroke styles -------
            if (cfg.doStrokeStyles && s !== 1) {
                PROG.phase(70, 72, "Scaling custom dashed / dotted stroke styles...");
                var nS = scaleStrokeStyles(doc, s);
                if (nS > 0) summary.push("Custom dashed/dotted stroke styles scaled: " + nS + ".");
            }

            // ------- live verify & fix -------
            if (live) {
                PROG.phase(72, 88, "Verifying live text, frames and tables; fixing unscaled values...");
                var lc = applyLiveFixes(live, s);
                var msg = "Live text/table check: " + lc.fixed + " value(s) corrected, " +
                          lc.ok + " already scaled by InDesign.";
                if (lc.other > 0) msg += " " + lc.other + " differed unexpectedly (left as-is).";
                if (lc.failed > 0) msg += " " + lc.failed + " could not be written.";
                summary.push(msg);
            }

            // ------- grids -------
            if (cfg.doGrids && s !== 1) {
                PROG.phase(88, 89, "Scaling baseline and document grids...");
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
                PROG.phase(89, 90, "Scaling effects manually (older InDesign)...");
                var allIt = doc.allPageItems;
                for (i = 0; i < allIt.length; i++) {
                    if (i % 10 === 0) PROG.step(i / allIt.length, "Item " + (i + 1) + "/" + allIt.length + " (" + className(allIt[i]) + ")");
                    scaleEffectsOn(allIt[i], s);
                }
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
                PROG.phase(90, 96, "Baking scaling to 100% (redefine scaling)...");
                var bake = function (items, f0, f1) {
                    var n = 0;
                    for (var k = 0; k < items.length; k++) {
                        if (k % 10 === 0) {
                            PROG.step(f0 + (f1 - f0) * k / items.length,
                                      "Item " + (k + 1) + "/" + items.length + " (" + className(items[k]) + ")");
                        }
                        if (isGraphicClass(items[k])) continue; // keep placed-image scale % meaningful
                        try { items[k].redefineScaling(); n++; } catch (e) {}
                    }
                    return n;
                };
                var baked = bake(doc.allPageItems, 0, 0.8);
                if (cfg.doParents) {
                    for (i = 0; i < doc.masterSpreads.length; i++) {
                        try { baked += bake(doc.masterSpreads[i].allPageItems, 0.8, 1); } catch (e) {}
                    }
                }
                summary.push("Scaling baked to 100% on " + baked + " item(s).");
            }

            PROG.phase(96, 99, "Finishing: bleed values, overset check, size report...");

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
            var b1 = refPage.bounds;
            summary.push("Page " + (cfg.refPageIndex + 1) + ": " +
                fmt((b0[3] - b0[1]) / U) + " x " + fmt((b0[2] - b0[0]) / U) + " " + UN + "  ->  " +
                fmt((b1[3] - b1[1]) / U) + " x " + fmt((b1[2] - b1[0]) / U) + " " + UN +
                (bleedPt > 0 ? " (incl. bleed)." : "."));

            if (cfg.pct > 100) {
                summary.push("Note: scaling UP reduces effective PPI of placed images — check the Links panel.");
            }
            if (cfg.pct < 12) {
                summary.push("Note: extreme downscales can hit InDesign's 0.1 pt minimum type size — spot-check small text.");
            }

        } finally {
            PROG.phase(99, 100, "Restoring locks and preferences...");

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

            PROG.close();
        }

        alert("Scale Artwork — done\n\n" + summary.join("\n"));
    }

    // ---------------------------------------------------------------------
    // Run
    // ---------------------------------------------------------------------
    var doc = app.activeDocument;

    var cfg = showDialog(doc);
    if (!cfg) return;

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
    if (cfg.onCopy && !doc.saved) {
        alert("The document has never been saved, so a copy can't be created.\n" +
              "Save it first, or untick 'Run on a copy'.");
        return;
    }

    PROG.open();

    if (cfg.onCopy) {
        PROG.phase(0, 0, "Saving a copy of the document and opening it...");
        var f = doc.fullName;
        var base = f.fsName.replace(/\.indd$/i, "");
        var tag = fmt(cfg.pct);
        var copyFile = new File(base + "_" + tag + "pc.indd");
        var n = 1;
        while (copyFile.exists) { copyFile = new File(base + "_" + tag + "pc_" + (n++) + ".indd"); }
        doc.saveACopy(copyFile);
        doc = app.open(copyFile);
    }

    try {
        app.doScript(function () { main(doc, cfg); },
            ScriptLanguage.JAVASCRIPT, [],
            UndoModes.ENTIRE_SCRIPT, "Scale Artwork Document");
    } catch (err) {
        PROG.close();
        alert("Scale Artwork — stopped with an error:\n\n" + err +
              (err.line ? "\n(line " + err.line + ")" : "") +
              "\n\nThe whole run is one undo step (Ctrl/Cmd-Z).");
    }

})();
