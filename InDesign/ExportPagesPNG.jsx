/**
 * ExportPagesPNG.jsx — v1.2
 * Export document pages as PNG files — one or a batch — with per-page
 * filenames, pixel-targeted sizing, and thumbnail preview.
 * InDesign ExtendScript (2023+ Mac/Win)
 *
 * v1.2 CHANGES
 * - Progress palette with "x of x" during thumbnail generation (previously
 *   looked halted on multi-page docs) and a matching counter on the export
 *   phase.
 *
 * v1.1 CHANGES
 * - Product-frame detection rebuilt. v1.0 scanned page.pageItems, which
 *   misses frames InDesign assigns to the spread instead of the page
 *   (anything hanging over the page edge or onto the pasteboard — common
 *   on signage artwork). v1.1 scans the spread's allPageItems flat
 *   (nested groups included, no fragile recursion), TRIMS names/labels
 *   before the case-insensitive "Product" compare, matches items to the
 *   page via parentPage, and accepts pasteboard Product frames on
 *   single-page spreads. Page rows now distinguish "no Product frame"
 *   from "Product frame empty".
 *
 * WHAT IT DOES
 *   - Lists every page with a thumbnail (temp low-res PNG, auto-deleted),
 *     a tick to include it, and an editable filename.
 *   - Default filename comes from a text frame on that page whose Layers
 *     name OR script label is "Product" (case-insensitive). Carriage
 *     returns / forced line breaks become spaces; illegal filename
 *     characters are stripped. Fallback: <docname>_p<n>.
 *   - Size: enter target WIDTH px or HEIGHT px (the other follows the
 *     page's aspect), or BOTH to verify: InDesign exports by uniform DPI
 *     only, so if both imply different DPIs (aspect mismatch > 0.5%) you
 *     are alerted per page with the actual output size (width wins) and
 *     can proceed or cancel. Expect +/-1 px from DPI rounding.
 *   - Per-page DPI is computed from that page's size, so mixed-size
 *     documents export correctly. DPI is clamped to InDesign's 1-2400
 *     range and clamps are reported.
 *   - Options: transparent background, overwrite existing files (off ->
 *     auto-unique suffix). Duplicate names within the batch get _p<n>.
 *   - Progress palette during export; full report at the end.
 *
 * The document is never modified — export prefs are snapshotted and
 * restored. Install in the Scripts Panel folder; run with a doc open.
 */

#target "InDesign"

(function () {
    if (!app.documents.length) { alert("Open a document first."); return; }
    var doc = app.activeDocument;
    if (doc.pages.length === 0) { alert("Document has no pages."); return; }

    var THUMB_H = 64;          // px, uniform row height
    var THUMB_ASK_OVER = 20;   // ask before generating this many thumbnails
    var DPI_MIN = 1, DPI_MAX = 2400;

    // ---------------- Gather page metadata ----------------

    var oldUnit = app.scriptPreferences.measurementUnit;
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
    var meta = [];
    try {
        for (var p = 0; p < doc.pages.length; p++) {
            var pg = doc.pages[p];
            var b = pg.bounds; // [y1,x1,y2,x2] pt
            var wIn = (b[3] - b[1]) / 72, hIn = (b[2] - b[0]) / 72;
            var prod = findProductText(pg);
            meta.push({
                absNum: pg.documentOffset + 1,
                pageName: pg.name,
                wIn: wIn, hIn: hIn,
                wMM: Math.round(wIn * 25.4 * 10) / 10,
                hMM: Math.round(hIn * 25.4 * 10) / 10,
                defName: prod.name !== null ? prod.name : (baseDocName() + "_p" + (pg.documentOffset + 1)),
                fromProduct: prod.name !== null,
                productFrameFound: prod.frameFound
            });
        }
    } finally {
        app.scriptPreferences.measurementUnit = oldUnit;
    }

    // ---------------- Thumbnails ----------------

    var thumbs = null; // array of File or null
    var wantThumbs = true;
    if (meta.length > THUMB_ASK_OVER) {
        wantThumbs = confirm("Generate thumbnails for " + meta.length + " pages?\n(This exports a tiny preview per page first \u2014 it can take a moment.)");
    }
    var prefsSnap = snapshotPngPrefs();
    var thumbDir = null;
    var tProg = null;
    try {
        if (wantThumbs) {
            thumbDir = new Folder(Folder.temp.fsName + "/idsnap_" + (new Date().getTime()));
            thumbDir.create();
            thumbs = [];
            tProg = progressWin("Generating thumbnails\u2026", meta.length);
            for (var t = 0; t < meta.length; t++) {
                tProg.set(t, "Page " + meta[t].absNum + "  \u2014  " + (t + 1) + " of " + meta.length);
                var tf = new File(thumbDir.fsName + "/t" + meta[t].absNum + ".png");
                var tdpi = clamp(THUMB_H / meta[t].hIn, DPI_MIN, DPI_MAX);
                try {
                    exportPagePNG(doc, meta[t].absNum, tf, tdpi, false);
                    thumbs.push(tf);
                } catch (eT) { thumbs.push(null); }
            }
            tProg.set(meta.length, "Done.");
        }
    } finally {
        restorePngPrefs(prefsSnap);
        if (tProg) tProg.close();
    }

    // ---------------- Dialog ----------------

    var cfg = showDialog(meta, thumbs);

    // thumbnails no longer needed
    cleanupThumbs();

    if (!cfg) return;

    // ---------------- Aspect verification (both fields given) ----------------

    if (cfg.wPx !== null && cfg.hPx !== null) {
        var mismatches = [];
        for (var m = 0; m < cfg.rows.length; m++) {
            var r0 = cfg.rows[m];
            if (!r0.doExport) continue;
            var mm = meta[r0.index];
            var dW = cfg.wPx / mm.wIn, dH = cfg.hPx / mm.hIn;
            var diff = Math.abs(dW - dH) / Math.max(dW, dH);
            if (diff > 0.005) {
                mismatches.push("Page " + mm.absNum + " (" + mm.wMM + " x " + mm.hMM + " mm): would be " +
                    cfg.wPx + " x " + Math.round(mm.hIn * dW) + " px, not " + cfg.wPx + " x " + cfg.hPx + " px");
            }
        }
        if (mismatches.length) {
            var msg = "Requested " + cfg.wPx + " x " + cfg.hPx + " px does not match the aspect ratio of " +
                mismatches.length + " selected page(s).\nInDesign scales uniformly \u2014 WIDTH will win:\n\n" +
                mismatches.join("\n") + "\n\nContinue?";
            if (!confirm(msg)) return;
        }
    }

    // ---------------- Export ----------------

    var lines = [], nOk = 0, fails = [];
    var usedNames = {}; // collision tracking within the batch
    var toExport = [];
    for (var q = 0; q < cfg.rows.length; q++) { if (cfg.rows[q].doExport) toExport.push(cfg.rows[q]); }
    if (!toExport.length) { alert("Nothing ticked \u2014 no pages exported."); return; }

    var prog = progressWin("Exporting PNGs\u2026", toExport.length);
    var snap2 = snapshotPngPrefs();
    try {
        for (var e = 0; e < toExport.length; e++) {
            var row = toExport[e];
            var pm = meta[row.index];
            prog.set(e, "Page " + pm.absNum + "  \u2192  " + row.fileName + ".png   \u2014  " + (e + 1) + " of " + toExport.length);

            // per-page DPI (width wins when both are given)
            var dpi;
            if (cfg.wPx !== null) dpi = cfg.wPx / pm.wIn;
            else dpi = cfg.hPx / pm.hIn;
            var clamped = false;
            var dpiC = clamp(dpi, DPI_MIN, DPI_MAX);
            if (dpiC !== dpi) clamped = true;

            // filename: batch-unique, then disk collision policy
            var base = row.fileName;
            if (usedNames[base.toLowerCase()]) base = base + "_p" + pm.absNum;
            usedNames[base.toLowerCase()] = true;

            var out = new File(cfg.folder.fsName + "/" + base + ".png");
            var renamedForDisk = false;
            if (out.exists && !cfg.overwrite) {
                var n = 2, cand = out;
                while (cand.exists) { cand = new File(cfg.folder.fsName + "/" + base + " (" + (n++) + ").png"); }
                out = cand; renamedForDisk = true;
            }

            try {
                exportPagePNG(doc, pm.absNum, out, dpiC, cfg.transparent);
                nOk++;
                var pxW = Math.round(pm.wIn * dpiC), pxH = Math.round(pm.hIn * dpiC);
                lines.push("EXPORTED  Page " + pm.absNum + "  \u2192  " + out.name +
                    "   (" + pxW + " x " + pxH + " px @ " + (Math.round(dpiC * 100) / 100) + " dpi" +
                    (clamped ? ", DPI CLAMPED" : "") +
                    (renamedForDisk ? ", renamed \u2014 file existed" : "") + ")");
            } catch (eX) {
                fails.push("Page " + pm.absNum + " ('" + base + "'): " + eX);
            }
        }
        prog.set(toExport.length, "Done.");
    } finally {
        restorePngPrefs(snap2);
        prog.close();
    }

    // ---------------- Report ----------------

    var rep = [];
    rep.push("Folder: " + cfg.folder.fsName);
    rep.push("Size: " + (cfg.wPx !== null ? cfg.wPx + " px wide" : "") +
             (cfg.wPx !== null && cfg.hPx !== null ? " x " : "") +
             (cfg.hPx !== null ? cfg.hPx + " px high" : "") +
             (cfg.wPx !== null && cfg.hPx !== null ? "  (width wins on mismatch)" : ""));
    rep.push("Transparent background: " + (cfg.transparent ? "yes" : "no") +
             "   |   Overwrite existing: " + (cfg.overwrite ? "yes" : "no"));
    rep.push("");
    for (var L = 0; L < lines.length; L++) rep.push(lines[L]);
    if (fails.length) {
        rep.push("");
        rep.push("ISSUES (" + fails.length + "):");
        for (var F = 0; F < fails.length; F++) rep.push("  " + fails[F]);
    }
    rep.push("");
    rep.push("Summary: " + nOk + " of " + toExport.length + " page(s) exported.");
    showReport("Export Pages to PNG \u2014 Report", rep.join("\n"));
    return;

    // =====================================================================
    // Helpers
    // =====================================================================

    function baseDocName() {
        var n = doc.name;
        return n.replace(/\.indd$/i, "");
    }

    function clamp(v, lo, hi) { return v < lo ? lo : (v > hi ? hi : v); }

    // ---- Product frame lookup ----
    // Scans the page's SPREAD flat via allPageItems (nested groups included),
    // trims names/labels before the case-insensitive compare, matches items
    // to the page via parentPage, and accepts pasteboard Product frames when
    // the spread has a single page. Returns { name: string|null, frameFound }.
    function findProductText(pageRef) {
        function tidy(v) { return String(v === null || v === undefined ? "" : v).replace(/^\s+|\s+$/g, "").toLowerCase(); }
        var best = null, frameFound = false;
        var spread = null, singlePage = false;
        try { spread = pageRef.parent; } catch (_) {}
        try { singlePage = spread && spread.pages.length === 1; } catch (_) {}

        var pool = null;
        try { pool = spread ? spread.allPageItems : pageRef.pageItems; } catch (_) {}
        if (!pool) return { name: null, frameFound: false };

        for (var i = 0; i < pool.length; i++) {
            var item = pool[i];
            try {
                if (!item || !item.isValid) continue;
                var cn = item.constructor ? item.constructor.name : "";
                if (cn !== "TextFrame") continue;
                var nm = "", lb = "";
                try { nm = tidy(item.name); } catch (_) {}
                try { lb = tidy(item.label); } catch (_) {}
                if (nm !== "product" && lb !== "product") continue;

                // Does this frame belong to THIS page?
                var onThisPage = false;
                try {
                    var pp = item.parentPage;
                    if (pp && pp.isValid) onThisPage = (pp.documentOffset === pageRef.documentOffset);
                    else onThisPage = singlePage; // pasteboard frame, unambiguous on a 1-page spread
                } catch (_) { onThisPage = singlePage; }
                if (!onThisPage) continue;

                frameFound = true;
                var txt = "";
                try { txt = item.texts[0].contents; } catch (_) { try { txt = item.contents; } catch (__) {} }
                if (typeof txt === "string") {
                    var clean = sanitizeName(txt);
                    if (clean !== "") { best = clean; break; }
                }
            } catch (_) {}
        }
        return { name: best, frameFound: frameFound };
    }

    function sanitizeName(s) {
        var out = String(s);
        out = out.replace(/[\r\n\u2028\u2029]+/g, " ");     // returns / breaks -> space
        out = out.replace(/[\\\/:\*\?"<>\|]/g, "");          // illegal filename chars
        out = out.replace(/\s+/g, " ");                       // collapse whitespace
        out = out.replace(/^\s+|\s+$/g, "");                  // trim
        out = out.replace(/[\. ]+$/g, "");                    // no trailing dots/spaces (Windows)
        return out;
    }

    // ---- PNG export ----

    function exportPagePNG(docRef, absPageNum, file, dpi, transparent) {
        var pp = app.pngExportPreferences;
        pp.pngExportRange = PNGExportRangeEnum.EXPORT_RANGE;
        pp.pageString = "+" + absPageNum;      // absolute page number, section-proof
        pp.exportResolution = dpi;
        pp.transparentBackground = transparent === true;
        pp.antiAlias = true;
        try { pp.pngQuality = PNGQualityEnum.MAXIMUM; } catch (_) {}
        try { pp.pngColorSpace = PNGColorSpaceEnum.RGB; } catch (_) {}
        try { pp.exportingSpread = false; } catch (_) {}
        try { pp.simulateOverprint = false; } catch (_) {}
        try { pp.useDocumentBleeds = false; } catch (_) {}
        docRef.exportFile(ExportFormat.PNG_FORMAT, file, false);
    }

    function snapshotPngPrefs() {
        var pp = app.pngExportPreferences, snap = {}, keys = [
            "pngExportRange", "pageString", "exportResolution", "transparentBackground",
            "antiAlias", "pngQuality", "pngColorSpace", "exportingSpread",
            "simulateOverprint", "useDocumentBleeds"
        ];
        for (var i = 0; i < keys.length; i++) { try { snap[keys[i]] = pp[keys[i]]; } catch (_) {} }
        return snap;
    }
    function restorePngPrefs(snap) {
        var pp = app.pngExportPreferences;
        for (var k in snap) { if (snap.hasOwnProperty(k)) { try { pp[k] = snap[k]; } catch (_) {} } }
    }

    function cleanupThumbs() {
        try {
            if (thumbs) { for (var i = 0; i < thumbs.length; i++) { try { if (thumbs[i]) thumbs[i].remove(); } catch (_) {} } }
            if (thumbDir) { try { thumbDir.remove(); } catch (_) {} }
        } catch (_) {}
    }

    // ---- Dialog ----

    function showDialog(metaArr, thumbArr) {
        var w = new Window("dialog", "Export Pages to PNG \u2014 v1.0");
        w.orientation = "column"; w.alignChildren = "fill";

        // Output folder
        var gF = w.add("group");
        gF.add("statictext", [0, 0, 90, 20], "Output folder:");
        var etFolder = gF.add("edittext", [0, 0, 520, 24], (doc.saved ? doc.filePath.fsName : Folder.desktop.fsName));
        var btnBrowse = gF.add("button", undefined, "Browse\u2026");
        btnBrowse.onClick = function () {
            var f = Folder.selectDialog("Choose the output folder", new Folder(etFolder.text));
            if (f) etFolder.text = f.fsName;
        };

        // Size
        var gS = w.add("panel", undefined, "Output size \u2014 enter Width OR Height (the other follows the page). Enter BOTH to verify aspect.");
        gS.orientation = "row"; gS.margins = 15;
        gS.add("statictext", undefined, "Width:");
        var etW = gS.add("edittext", [0, 0, 80, 24], "");
        gS.add("statictext", undefined, "px      Height:");
        var etH = gS.add("edittext", [0, 0, 80, 24], "");
        gS.add("statictext", undefined, "px");
        var stInfo = w.add("statictext", [0, 0, 860, 18], "");

        function updateInfo() {
            var wv = parseFloat(etW.text), hv = parseFloat(etH.text);
            var m0 = metaArr[0];
            if (!isNaN(wv) && wv > 0 && (isNaN(hv) || etH.text === "")) {
                var d = wv / m0.wIn;
                stInfo.text = "Page 1 \u2192 " + Math.round(wv) + " x " + Math.round(m0.hIn * d) + " px @ " + (Math.round(d * 100) / 100) + " dpi";
            } else if (!isNaN(hv) && hv > 0 && (isNaN(wv) || etW.text === "")) {
                var d2 = hv / m0.hIn;
                stInfo.text = "Page 1 \u2192 " + Math.round(m0.wIn * d2) + " x " + Math.round(hv) + " px @ " + (Math.round(d2 * 100) / 100) + " dpi";
            } else if (!isNaN(wv) && !isNaN(hv) && wv > 0 && hv > 0) {
                var dW = wv / m0.wIn, dH = hv / m0.hIn;
                var diff = Math.abs(dW - dH) / Math.max(dW, dH);
                stInfo.text = "Page 1 \u2192 forced " + Math.round(wv) + " x " + Math.round(hv) + " px" +
                    (diff > 0.005 ? "  \u26A0 aspect differs from page (" + Math.round(diff * 1000) / 10 + "%) \u2014 width will win" : "  (aspect matches)");
            } else {
                stInfo.text = "";
            }
        }
        etW.onChanging = updateInfo; etH.onChanging = updateInfo;

        // Options
        var gO = w.add("group");
        var chkTrans = gO.add("checkbox", undefined, "Transparent background");
        chkTrans.value = false;
        var chkOver = gO.add("checkbox", undefined, "Overwrite existing files (off = auto-suffix)");
        chkOver.value = true;

        // Pages list (scrollable when long)
        var pnl = w.add("panel", undefined, "Pages \u2014 tick to export, edit filenames (\".png\" is added automatically)");
        pnl.alignChildren = "left"; pnl.margins = 12;
        var rowH = (thumbArr ? THUMB_H + 10 : 30);
        var visRows = Math.min(metaArr.length, thumbArr ? 6 : 12);
        var needScroll = metaArr.length > visRows;
        var host = pnl.add("group");
        var viewport = host.add("group");
        viewport.orientation = "column";
        var inner = viewport.add("group");
        inner.orientation = "column"; inner.alignChildren = "left"; inner.spacing = 4;

        var rowCtls = [];
        for (var i = 0; i < metaArr.length; i++) {
            var mm = metaArr[i];
            var g = inner.add("group");
            g.alignChildren = "center";
            var chk = g.add("checkbox", undefined, "");
            chk.value = true;
            if (thumbArr && thumbArr[i]) {
                try { g.add("image", undefined, thumbArr[i]); } catch (_) { g.add("statictext", [0,0,40,20], ""); }
            }
            var lbl = g.add("statictext", [0, 0, 190, 18], "Pg " + mm.absNum + "  (" + mm.wMM + " x " + mm.hMM + " mm)" +
                (mm.fromProduct ? "" : (mm.productFrameFound ? "  \u2022 Product frame empty" : "  \u2022 no Product frame")));
            var et = g.add("edittext", [0, 0, 330, 24], mm.defName);
            rowCtls.push({ chk: chk, et: et });
        }

        var sb = null;
        if (needScroll) {
            viewport.maximumSize = [900, visRows * rowH];
            viewport.size = [880, visRows * rowH];
            sb = host.add("scrollbar", [0, 0, 18, visRows * rowH]);
            sb.onChanging = function () { inner.location = [inner.location[0], -sb.value]; };
        }

        var gSel = w.add("group");
        var btnAll = gSel.add("button", undefined, "Select All");
        var btnNone = gSel.add("button", undefined, "Select None");
        btnAll.onClick = function () { for (var j = 0; j < rowCtls.length; j++) rowCtls[j].chk.value = true; };
        btnNone.onClick = function () { for (var j = 0; j < rowCtls.length; j++) rowCtls[j].chk.value = false; };

        var gB = w.add("group"); gB.alignment = "right";
        gB.add("button", undefined, "Cancel", { name: "cancel" });
        var btnGo = gB.add("button", undefined, "Export", { name: "ok" });

        w.onShow = function () {
            if (sb) {
                var innerH = inner.size ? inner.size.height : metaArr.length * rowH;
                var viewH = viewport.size.height;
                sb.minvalue = 0;
                sb.maxvalue = Math.max(0, innerH - viewH);
                sb.stepdelta = rowH;
                sb.jumpdelta = viewH;
            }
        };

        // validate on OK
        var result = null;
        btnGo.onClick = function () {
            var wv = etW.text === "" ? null : parseFloat(etW.text);
            var hv = etH.text === "" ? null : parseFloat(etH.text);
            if ((wv === null && hv === null) || (wv !== null && (isNaN(wv) || wv <= 0)) || (hv !== null && (isNaN(hv) || hv <= 0))) {
                alert("Enter a positive Width and/or Height in pixels.");
                return;
            }
            var folder = new Folder(etFolder.text);
            if (!folder.exists) {
                if (!confirm("Folder does not exist:\n" + etFolder.text + "\n\nCreate it?")) return;
                if (!folder.create()) { alert("Could not create the folder."); return; }
            }
            var rows = [], anyBad = [];
            for (var j = 0; j < rowCtls.length; j++) {
                var nm = sanitizeName(rowCtls[j].et.text);
                if (rowCtls[j].chk.value && nm === "") anyBad.push("Pg " + metaArr[j].absNum);
                rows.push({ index: j, doExport: rowCtls[j].chk.value, fileName: nm });
            }
            if (anyBad.length) { alert("Empty filename on: " + anyBad.join(", ")); return; }
            result = {
                folder: folder,
                wPx: wv !== null ? Math.round(wv) : null,
                hPx: hv !== null ? Math.round(hv) : null,
                transparent: chkTrans.value,
                overwrite: chkOver.value,
                rows: rows
            };
            w.close(1);
        };

        return (w.show() === 1) ? result : null;
    }

    // ---- Progress & report ----

    function progressWin(title, max) {
        var w = new Window("palette", title);
        w.orientation = "column"; w.alignChildren = "fill";
        var t = w.add("statictext", [0, 0, 380, 18], "");
        var pb = w.add("progressbar", [0, 0, 380, 12], 0, Math.max(1, max));
        w.show();
        return {
            set: function (v, msg) { try { pb.value = v; t.text = msg; w.update(); } catch (_) {} },
            close: function () { try { w.close(); } catch (_) {} }
        };
    }

    function showReport(title, text) {
        var w = new Window("dialog", title);
        w.orientation = "column"; w.alignChildren = "fill";
        var et = w.add("edittext", [0, 0, 760, 420], text, { multiline: true, readonly: true, scrolling: true });
        var g = w.add("group"); g.alignment = "right";
        g.add("button", undefined, "Close", { name: "ok" });
        w.show();
    }

})();
