#target photoshop

/*
 * =============================================================================
 *  Smart Object Artwork Placer
 *  Photoshop ExtendScript (.jsx)  -  for Photoshop CC 2025 (v26.x), Windows
 * =============================================================================
 *
 *  PURPOSE
 *  -------
 *  Automates placing artwork into Smart Object layers of print product mockup
 *  templates (flyers, brochures, booklets). The template PSD must already be
 *  open in Photoshop, with its Smart Object layers pre-distorted to match the
 *  product shot. For every Smart Object, the script opens its contents, clears
 *  it, places the assigned artwork, scales it to FILL the canvas (plus optional
 *  bleed overhang), flattens, and saves so the parent document updates.
 *  Finally it saves a layered PSD copy and a flattened JPG (quality 8) to a
 *  user-chosen working directory.
 *
 *  WORKFLOW  (multi-step ScriptUI wizard)
 *  --------------------------------------
 *    Step 1  Project Setup .......... filename (no extension) + working folder
 *    Step 2  Smart Object Discovery . scan + order the Smart Objects
 *    Step 3  Bleed Settings ......... top / bottom / left / right + mm|px
 *    Step 4  Artwork Source ......... image file / PDF / folder / multi-page PDF
 *    Step 5  Review & Process ....... summary, progress bar, processing + export
 *
 *  Author: Abacus AI Agent
 *  License: provided as-is for the requesting user's workflow.
 * =============================================================================
 */

(function () {

    // -------------------------------------------------------------------------
    //  PRE-FLIGHT
    // -------------------------------------------------------------------------
    if (app.documents.length === 0) {
        alert("Smart Object Artwork Placer\n\n" +
              "Please open a mockup template (a PSD with Smart Object layers) " +
              "before running this script.");
        return;
    }

    // -------------------------------------------------------------------------
    //  GLOBAL STATE
    // -------------------------------------------------------------------------
    var CONFIG = {
        parentDoc:    app.activeDocument,
        filename:     "",
        workingDir:   null,
        soLayers:     [],     // [{layer, name, idx, order}] in palette order
        orderedSOs:   [],      // soLayers sorted by user-assigned order
        isSingle:     false,
        bleed:        { top: 0, bottom: 0, left: 0, right: 0, unit: "mm" },
        sourceType:   "",      // imageFile | pdfFile | imageFolder | pdfMulti
        sourceFile:   null,
        sourceFolder: null,
        imageFiles:   [],
        pdfPageCount: 0,
        singlePdfPage: 1,
        pdfStartPage: 1,
        pdfMode:      "sequential", // sequential | custom
        pdfMapping:   []
    };

    var ui = {}; // holds references to dynamically-created controls

    // =========================================================================
    //  GENERIC HELPERS
    // =========================================================================
    function trimStr(s) { return String(s).replace(/^\s+|\s+$/g, ""); }

    function sanitizeName(s) {
        // Strip characters illegal in Windows filenames.
        return trimStr(s).replace(/[\\\/:*?"<>|]/g, "_");
    }

    // Natural / human sort so file1, file2, ... file10 order correctly.
    function naturalCompare(a, b) {
        a = String(a).toLowerCase();
        b = String(b).toLowerCase();
        var ax = [], bx = [];
        a.replace(/(\d+)|(\D+)/g, function (_, $1, $2) { ax.push([$1 || Infinity, $2 || ""]); return ""; });
        b.replace(/(\d+)|(\D+)/g, function (_, $1, $2) { bx.push([$1 || Infinity, $2 || ""]); return ""; });
        while (ax.length && bx.length) {
            var an = ax.shift(), bn = bx.shift();
            var nn = (an[0] - bn[0]) || ((an[1] < bn[1]) ? -1 : (an[1] > bn[1]) ? 1 : 0);
            if (nn) { return nn; }
        }
        return ax.length - bx.length;
    }

    // Convert a bleed value to pixels. mm uses the document resolution (ppi).
    function computeBleedPx(val, unit, resolution) {
        if (!val || val <= 0) { return 0; }
        if (unit === "px") { return val; }
        return (val / 25.4) * resolution; // 1 inch = 25.4 mm
    }

    // =========================================================================
    //  SMART OBJECT ENUMERATION  (recursive, returns palette order top->bottom)
    // =========================================================================
    function findSmartObjects(doc) {
        var out = [];
        function scan(layers) {
            for (var i = 0; i < layers.length; i++) {
                var L = layers[i];
                if (L.typename === "LayerSet") {
                    scan(L.layers); // recurse into groups (templates are usually flat)
                } else {
                    var isSO = false;
                    try { isSO = (L.kind === LayerKind.SMARTOBJECT); } catch (e) { isSO = false; }
                    if (isSO) { out.push({ layer: L, name: L.name, idx: 0, order: 0 }); }
                }
            }
        }
        scan(doc.layers);
        for (var k = 0; k < out.length; k++) { out[k].idx = k; out[k].order = k + 1; }
        return out;
    }

    // =========================================================================
    //  PDF PAGE-COUNT DETECTION  (heuristic; user can override manually)
    // =========================================================================
    function getPDFPageCount(file) {
        var count = 0;
        try {
            var f = new File(file);
            f.encoding = "BINARY";
            if (f.open("r")) {
                var content = f.read();
                f.close();
                // Primary: count "/Type /Page" markers (avoid matching "/Pages").
                var m = content.match(/\/Type\s*\/Page[^s]/g);
                if (m) { count = m.length; }
                // Fallback: the page-tree "/Count N" (take the largest value found).
                if (count === 0) {
                    var counts = content.match(/\/Count\s+(\d+)/g);
                    if (counts && counts.length) {
                        var max = 0;
                        for (var i = 0; i < counts.length; i++) {
                            var nm = counts[i].match(/(\d+)/);
                            if (nm) { var v = parseInt(nm[1], 10); if (v > max) { max = v; } }
                        }
                        count = max;
                    }
                }
            }
        } catch (e) { count = 0; }
        return count;
    }

    // =========================================================================
    //  COLOR-MODE HELPERS
    // =========================================================================
    function targetOpenMode(doc) {
        try {
            switch (doc.mode) {
                case DocumentMode.GRAYSCALE: return OpenDocumentMode.GRAYSCALE;
                case DocumentMode.CMYK:      return OpenDocumentMode.CMYK;
                case DocumentMode.LAB:       return OpenDocumentMode.LAB;
                default:                     return OpenDocumentMode.RGB;
            }
        } catch (e) { return OpenDocumentMode.RGB; }
    }

    function matchColorMode(srcDoc, targetDoc) {
        try {
            if (srcDoc.mode === targetDoc.mode) { return; }
            var cm;
            switch (targetDoc.mode) {
                case DocumentMode.GRAYSCALE:    cm = ChangeMode.GRAYSCALE; break;
                case DocumentMode.CMYK:         cm = ChangeMode.CMYK; break;
                case DocumentMode.LAB:          cm = ChangeMode.LAB; break;
                case DocumentMode.INDEXEDCOLOR: cm = ChangeMode.INDEXEDCOLOR; break;
                case DocumentMode.BITMAP:       cm = ChangeMode.BITMAP; break;
                default:                        cm = ChangeMode.RGB;
            }
            app.activeDocument = srcDoc;
            srcDoc.changeMode(cm);
        } catch (e) { /* best effort */ }
    }

    // =========================================================================
    //  PHOTOSHOP OPERATIONS
    // =========================================================================

    // Open the contents (.psb) of a Smart Object using the Action Manager so we
    // can suppress dialogs. Includes a wait-and-verify loop (a known timing
    // quirk: activeDocument may not switch immediately).
    function openSmartObjectContents(soLayer, parentDoc) {
        app.activeDocument = parentDoc;
        parentDoc.activeLayer = soLayer;
        var before = app.documents.length;
        executeAction(stringIDToTypeID("placedLayerEditContents"),
                      new ActionDescriptor(), DialogModes.NO);
        var tries = 0;
        while (app.documents.length <= before && tries < 50) {
            $.sleep(100);
            tries++;
        }
        return app.activeDocument; // the freshly opened PSB
    }

    // Reduce the document to a single, empty layer (clean slate).
    function cleanSlate(doc) {
        app.activeDocument = doc;
        // Remove everything except one layer.
        try {
            while (doc.layers.length > 1) {
                doc.layers[doc.layers.length - 1].remove();
            }
        } catch (e) { /* ignore */ }
        // Convert a background layer to a normal layer so it can be cleared.
        try {
            var L = doc.layers[0];
            doc.activeLayer = L;
            if (L.isBackgroundLayer) { L.isBackgroundLayer = false; }
        } catch (e) { /* ignore */ }
        // Empty the remaining layer.
        try {
            doc.selection.selectAll();
            doc.selection.clear();
            doc.selection.deselect();
        } catch (e) { /* ignore */ }
    }

    // Open the artwork (image or specific PDF page) as a temporary document,
    // flatten it, then transfer its content into targetDoc. Returns the placed
    // ArtLayer (sitting at the top of targetDoc).
    function placeArtwork(targetDoc, file, isPdf, pageNum) {
        var srcDoc;
        if (isPdf) {
            var opts = new PDFOpenOptions();
            opts.antiAlias = true;
            opts.bitsPerChannel = BitsPerChannelType.EIGHT;
            opts.mode = targetOpenMode(targetDoc);
            try { opts.resolution = targetDoc.resolution; } catch (e) { opts.resolution = 300; }
            opts.usePageNumber = true;   // page refers to the PDF page number
            opts.page = pageNum;
            opts.cropPage = CropToType.MEDIABOX;
            opts.suppressWarnings = true;
            srcDoc = app.open(file, opts);
        } else {
            srcDoc = app.open(file);
        }

        app.activeDocument = srcDoc;
        try { srcDoc.flatten(); } catch (e) { /* may already be flat */ }
        matchColorMode(srcDoc, targetDoc);

        var placed = null;
        // Preferred: duplicate the layer directly into the target document.
        // duplicate() returns the new layer (which belongs to targetDoc).
        try {
            var dupLayer = srcDoc.artLayers[0].duplicate(targetDoc, ElementPlacement.PLACEATBEGINNING);
            app.activeDocument = targetDoc;
            placed = dupLayer ? dupLayer : targetDoc.artLayers[0];
        } catch (eDup) {
            // Fallback: copy / paste (also reliable, pastes centered).
            try {
                app.activeDocument = srcDoc;
                srcDoc.selection.selectAll();
                srcDoc.selection.copy();
                srcDoc.selection.deselect();
                app.activeDocument = targetDoc;
                placed = targetDoc.paste();
            } catch (ePaste) {
                try { srcDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e) {}
                throw new Error("Unable to transfer artwork into the Smart Object (" + ePaste.message + ").");
            }
        }

        try { srcDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e) {}
        app.activeDocument = targetDoc;
        if (!placed) { placed = targetDoc.artLayers[0]; }
        return placed;
    }

    // Scale the layer to FILL the canvas (+ bleed) preserving aspect ratio,
    // then center it on the canvas.
    function scaleAndCenter(doc, layer) {
        var oldUnits = app.preferences.rulerUnits;
        app.preferences.rulerUnits = Units.PIXELS;
        try {
            var res = doc.resolution;
            var canvasW = doc.width.value;   // pixels (rulerUnits == PIXELS)
            var canvasH = doc.height.value;

            var leftPx   = computeBleedPx(CONFIG.bleed.left,   CONFIG.bleed.unit, res);
            var rightPx  = computeBleedPx(CONFIG.bleed.right,  CONFIG.bleed.unit, res);
            var topPx    = computeBleedPx(CONFIG.bleed.top,    CONFIG.bleed.unit, res);
            var bottomPx = computeBleedPx(CONFIG.bleed.bottom, CONFIG.bleed.unit, res);

            // Target area = canvas plus bleed overhang on each side.
            var targetW = canvasW + leftPx + rightPx;
            var targetH = canvasH + topPx + bottomPx;

            var b = layer.bounds;
            var artW = b[2].value - b[0].value;
            var artH = b[3].value - b[1].value;
            if (artW <= 0 || artH <= 0) { return; }

            // Compute per-axis scale factors and take the LARGER (fill, not fit).
            var scaleW = targetW / artW;
            var scaleH = targetH / artH;
            var scale  = Math.max(scaleW, scaleH) * 100; // resize() takes percent

            layer.resize(scale, scale, AnchorPosition.MIDDLECENTER);

            // Center the (now scaled) layer on the canvas.
            b = layer.bounds;
            var layerCX = (b[0].value + b[2].value) / 2;
            var layerCY = (b[1].value + b[3].value) / 2;
            layer.translate(canvasW / 2 - layerCX, canvasH / 2 - layerCY);
        } finally {
            app.preferences.rulerUnits = oldUnits;
        }
    }

    // Save a document with a JPEG/Photoshop options object; retry without an
    // embedded profile if the document has no color profile.
    function safeSaveAs(d, file, opts) {
        try {
            d.saveAs(file, opts, false, Extension.LOWERCASE);
        } catch (e) {
            try {
                opts.embedColorProfile = false;
                d.saveAs(file, opts, false, Extension.LOWERCASE);
            } catch (e2) {
                throw e2;
            }
        }
    }

    // Save layered PSD copy + flattened JPG (quality 8). Both stay open.
    function exportFiles() {
        var doc = CONFIG.parentDoc;
        app.activeDocument = doc;
        var dir = CONFIG.workingDir.fsName;

        // 1) Layered PSD (this becomes the open, saved working file).
        var psdFile = new File(dir + "/" + CONFIG.filename + ".psd");
        var psdOpts = new PhotoshopSaveOptions();
        psdOpts.layers = true;
        psdOpts.embedColorProfile = true;
        psdOpts.alphaChannels = true;
        psdOpts.annotations = true;
        psdOpts.spotColors = true;
        safeSaveAs(doc, psdFile, psdOpts);

        // 2) Flattened JPG from a duplicate (keeps the layered PSD intact/open).
        var dup = doc.duplicate(CONFIG.filename + " (JPG)", false);
        app.activeDocument = dup;
        dup.flatten();
        var jpgFile = new File(dir + "/" + CONFIG.filename + ".jpg");
        var jpgOpts = new JPEGSaveOptions();
        jpgOpts.quality = 8;                              // out of 12
        jpgOpts.formatOptions = FormatOptions.STANDARDBASELINE;
        jpgOpts.embedColorProfile = true;
        safeSaveAs(dup, jpgFile, jpgOpts);

        // Leave the JPG active so the user can immediately crop it to 500x500.
        return { psd: psdFile, jpg: jpgFile };
    }

    // =========================================================================
    //  ARTWORK ASSIGNMENT
    // =========================================================================
    function artworkCount() {
        switch (CONFIG.sourceType) {
            case "imageFile":   return 1;
            case "pdfFile":     return 1;
            case "imageFolder": return CONFIG.imageFiles.length;
            case "pdfMulti":
                if (CONFIG.pdfMode === "custom") { return CONFIG.pdfMapping.length; }
                if (CONFIG.pdfPageCount > 0) {
                    var avail = CONFIG.pdfPageCount - CONFIG.pdfStartPage + 1;
                    return avail < 0 ? 0 : avail;
                }
                return CONFIG.orderedSOs.length; // unknown count -> assume enough
        }
        return 0;
    }

    // Returns {file, isPdf, page} for the i-th Smart Object, or null if none.
    function getArtworkForIndex(i) {
        if (CONFIG.sourceType === "imageFile") {
            return (i === 0) ? { file: CONFIG.sourceFile, isPdf: false, page: 1 } : null;
        }
        if (CONFIG.sourceType === "pdfFile") {
            return (i === 0) ? { file: CONFIG.sourceFile, isPdf: true, page: CONFIG.singlePdfPage } : null;
        }
        if (CONFIG.sourceType === "imageFolder") {
            return (i < CONFIG.imageFiles.length)
                ? { file: CONFIG.imageFiles[i], isPdf: false, page: 1 } : null;
        }
        if (CONFIG.sourceType === "pdfMulti") {
            var page;
            if (CONFIG.pdfMode === "custom") {
                if (i >= CONFIG.pdfMapping.length) { return null; }
                page = CONFIG.pdfMapping[i];
            } else {
                page = CONFIG.pdfStartPage + i;
                if (CONFIG.pdfPageCount > 0 && page > CONFIG.pdfPageCount) { return null; }
            }
            if (!page || page < 1) { return null; }
            return { file: CONFIG.sourceFile, isPdf: true, page: page };
        }
        return null;
    }

    function sourceDescription() {
        switch (CONFIG.sourceType) {
            case "imageFile":
                return "Single image  (" + (CONFIG.sourceFile ? CONFIG.sourceFile.name : "?") + ")";
            case "pdfFile":
                return "PDF page " + CONFIG.singlePdfPage + "  (" + (CONFIG.sourceFile ? CONFIG.sourceFile.name : "?") + ")";
            case "imageFolder":
                return CONFIG.imageFiles.length + " image(s) from folder  (" +
                       (CONFIG.sourceFolder ? CONFIG.sourceFolder.name : "?") + ")";
            case "pdfMulti":
                if (CONFIG.pdfMode === "custom") {
                    return "Multi-page PDF, custom page mapping  (" + (CONFIG.sourceFile ? CONFIG.sourceFile.name : "?") + ")";
                }
                return "Multi-page PDF, sequential from page " + CONFIG.pdfStartPage +
                       "  (" + (CONFIG.sourceFile ? CONFIG.sourceFile.name : "?") + ")";
        }
        return "(not selected)";
    }

    // =========================================================================
    //  SCAN THE DOCUMENT BEFORE BUILDING THE UI
    // =========================================================================
    CONFIG.soLayers = findSmartObjects(CONFIG.parentDoc);
    if (CONFIG.soLayers.length === 0) {
        alert("No Smart Object layers were found in the active document.\n\n" +
              "Open a template whose artwork areas are Smart Objects and try again.");
        return;
    }
    CONFIG.isSingle = (CONFIG.soLayers.length === 1);
    CONFIG.orderedSOs = CONFIG.soLayers.slice(); // default order

    // =========================================================================
    //  BUILD THE WIZARD WINDOW
    // =========================================================================
    var win = new Window("dialog", "Smart Object Artwork Placer");
    win.alignChildren = ["fill", "top"];
    win.spacing = 10;
    win.margins = 16;

    // ---- Header ----
    var header = win.add("panel");
    header.alignChildren = ["fill", "top"];
    header.margins = 12;
    var titleTxt = header.add("statictext", undefined, "");
    try { titleTxt.graphics.font = ScriptUI.newFont(titleTxt.graphics.font.name, "Bold", 16); } catch (e) {}
    titleTxt.preferredSize = [470, 24];
    var subTxt = header.add("statictext", undefined, "", { multiline: true });
    subTxt.preferredSize = [470, 32];

    // ---- Content (stacked step panels; only one visible at a time) ----
    var content = win.add("group");
    content.orientation = "stack";
    content.alignChildren = ["fill", "top"];

    var g1 = content.add("group"); g1.orientation = "column"; g1.alignChildren = ["fill", "top"]; g1.spacing = 8;
    var g2 = content.add("group"); g2.orientation = "column"; g2.alignChildren = ["fill", "top"]; g2.spacing = 6;
    var g3 = content.add("group"); g3.orientation = "column"; g3.alignChildren = ["fill", "top"]; g3.spacing = 8;
    var g4 = content.add("group"); g4.orientation = "column"; g4.alignChildren = ["fill", "top"]; g4.spacing = 8;
    var g5 = content.add("group"); g5.orientation = "column"; g5.alignChildren = ["fill", "top"]; g5.spacing = 8;

    // ---- Footer (navigation) ----
    var footer = win.add("group");
    footer.alignment = "fill";
    var cancelBtn = footer.add("button", undefined, "Cancel");
    cancelBtn.alignment = "left";
    var filler = footer.add("statictext", undefined, "");
    filler.alignment = ["fill", "center"];
    var backBtn = footer.add("button", undefined, "\u25C0 Back");
    var nextBtn = footer.add("button", undefined, "Next \u25B6");
    win.cancelElement = cancelBtn;

    // -------------------------------------------------------------------------
    //  STEP 1  -  Project Setup
    // -------------------------------------------------------------------------
    g1.add("statictext", undefined,
        "Enter a base filename (used for BOTH the PSD and the JPG export). Do not include an extension.");
    var s1row1 = g1.add("group");
    s1row1.add("statictext", undefined, "Filename:");
    ui.fileName = s1row1.add("edittext", undefined, "");
    ui.fileName.characters = 34;

    g1.add("statictext", undefined, "Choose the working directory where the files will be saved.");
    var s1row2 = g1.add("group");
    ui.dirField = s1row2.add("edittext", undefined, "");
    ui.dirField.characters = 34;
    ui.dirField.enabled = false;
    var browseDirBtn = s1row2.add("button", undefined, "Browse\u2026");
    browseDirBtn.onClick = function () {
        var f = Folder.selectDialog("Select the working directory");
        if (f) { CONFIG.workingDir = f; ui.dirField.text = f.fsName; }
    };

    // -------------------------------------------------------------------------
    //  STEP 2  -  Smart Object Discovery & Ordering
    // -------------------------------------------------------------------------
    function buildStep2() {
        if (CONFIG.isSingle) {
            g2.add("statictext", undefined, "Found 1 Smart Object layer:");
            var only = g2.add("statictext", undefined, "    \u2022 " + CONFIG.soLayers[0].name);
            only.preferredSize = [470, 20];
            g2.add("statictext", undefined, "Ordering is not required for a single Smart Object \u2014 click Next to continue.");
        } else {
            g2.add("statictext", undefined,
                "Found " + CONFIG.soLayers.length + " Smart Object layer(s), listed in layer-palette order (top to bottom).");
            g2.add("statictext", undefined,
                "Set the order number to control which artwork is placed into which Smart Object.");

            var listPanel = g2.add("panel");
            listPanel.alignChildren = ["fill", "top"];
            listPanel.margins = 10;
            listPanel.spacing = 3;

            var hr = listPanel.add("group");
            var h1 = hr.add("statictext", undefined, "Order"); h1.preferredSize = [46, 20];
            hr.add("statictext", undefined, "Smart Object Layer");

            ui.orderFields = [];
            for (var i = 0; i < CONFIG.soLayers.length; i++) {
                var row = listPanel.add("group");
                var ef = row.add("edittext", undefined, String(i + 1));
                ef.preferredSize = [46, 22];
                var nm = row.add("statictext", undefined, CONFIG.soLayers[i].name);
                nm.preferredSize = [380, 22];
                ui.orderFields.push(ef);
            }
        }
    }

    // -------------------------------------------------------------------------
    //  STEP 3  -  Bleed Settings
    // -------------------------------------------------------------------------
    function buildStep3() {
        g3.add("statictext", undefined,
            "Bleed scales the artwork LARGER than the Smart Object canvas (overhang for print trim).", { multiline: true }).preferredSize = [470, 18];
        g3.add("statictext", undefined,
            "Leave a field blank or 0 for no bleed on that side.");

        var unitRow = g3.add("group");
        unitRow.add("statictext", undefined, "Units:");
        ui.unitDD = unitRow.add("dropdownlist", undefined, ["mm", "px"]);
        ui.unitDD.selection = 0; // mm default

        var bp = g3.add("panel");
        bp.alignChildren = ["left", "top"];
        bp.margins = 12;
        bp.spacing = 6;

        function bleedRow(label, def) {
            var r = bp.add("group");
            var lt = r.add("statictext", undefined, label); lt.preferredSize = [70, 22];
            var ef = r.add("edittext", undefined, def); ef.characters = 8;
            return ef;
        }
        ui.bTop    = bleedRow("Top:", "0");
        ui.bBottom = bleedRow("Bottom:", "0");
        ui.bLeft   = bleedRow("Left:", "0");
        ui.bRight  = bleedRow("Right:", "0");
    }

    // -------------------------------------------------------------------------
    //  STEP 4  -  Artwork Source Selection (adapts to single vs. multi SO)
    // -------------------------------------------------------------------------
    function buildStep4() {
        if (CONFIG.isSingle) {
            // ---------- Single Smart Object ----------
            g4.add("statictext", undefined, "Select one artwork file to place into the Smart Object.");
            var rg = g4.add("group");
            ui.s_radioImage = rg.add("radiobutton", undefined, "Image file (JPG / PNG / TIFF / PSD)");
            ui.s_radioPdf   = rg.add("radiobutton", undefined, "PDF file");
            ui.s_radioImage.value = true;

            var prow = g4.add("group");
            ui.s_path = prow.add("edittext", undefined, "");
            ui.s_path.characters = 34; ui.s_path.enabled = false;
            ui.s_browseBtn = prow.add("button", undefined, "Browse\u2026");

            ui.s_pdfPanel = g4.add("panel");
            ui.s_pdfPanel.alignChildren = ["left", "top"];
            ui.s_pdfPanel.margins = 10; ui.s_pdfPanel.spacing = 6;
            ui.s_pdfCountTxt = ui.s_pdfPanel.add("statictext", undefined, "PDF page count: \u2014");
            ui.s_pdfCountTxt.preferredSize = [320, 20];
            var pr2 = ui.s_pdfPanel.add("group");
            pr2.add("statictext", undefined, "Use page:");
            ui.s_pageField = pr2.add("edittext", undefined, "1"); ui.s_pageField.characters = 5;
            ui.s_pdfPanel.visible = false;

            ui.s_radioImage.onClick = ui.s_radioPdf.onClick = function () {
                ui.s_path.text = ""; CONFIG.sourceFile = null;
                ui.s_pdfPanel.visible = ui.s_radioPdf.value;
                win.layout.layout(true);
            };
            ui.s_browseBtn.onClick = function () {
                var f;
                if (ui.s_radioPdf.value) {
                    f = File.openDialog("Select a PDF file", "PDF:*.pdf");
                } else {
                    f = File.openDialog("Select an image file",
                        "Images:*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.psd");
                }
                if (f) {
                    CONFIG.sourceFile = f;
                    ui.s_path.text = f.fsName;
                    if (ui.s_radioPdf.value) {
                        var c = getPDFPageCount(f);
                        CONFIG.pdfPageCount = c;
                        ui.s_pdfCountTxt.text = "PDF page count: " +
                            (c > 0 ? c : "unknown \u2014 enter the page manually");
                    }
                }
            };
        } else {
            // ---------- Multiple Smart Objects ----------
            var n = CONFIG.soLayers.length;
            g4.add("statictext", undefined, "Select the artwork source for the " + n + " Smart Objects.");
            var rgm = g4.add("group");
            ui.m_radioFolder = rgm.add("radiobutton", undefined, "Folder of images");
            ui.m_radioPdf    = rgm.add("radiobutton", undefined, "Multi-page PDF");
            ui.m_radioFolder.value = true;

            var prowm = g4.add("group");
            ui.m_path = prowm.add("edittext", undefined, "");
            ui.m_path.characters = 34; ui.m_path.enabled = false;
            ui.m_browseBtn = prowm.add("button", undefined, "Browse\u2026");

            ui.m_folderStatus = g4.add("statictext", undefined, "");
            ui.m_folderStatus.preferredSize = [470, 20];

            // ---- PDF options sub-panel ----
            ui.m_pdfPanel = g4.add("panel");
            ui.m_pdfPanel.alignChildren = ["left", "top"];
            ui.m_pdfPanel.margins = 10; ui.m_pdfPanel.spacing = 6;
            ui.m_pdfCountTxt = ui.m_pdfPanel.add("statictext", undefined, "PDF page count: \u2014");
            ui.m_pdfCountTxt.preferredSize = [430, 20];

            var modeRow = ui.m_pdfPanel.add("group");
            ui.m_modeSeq   = modeRow.add("radiobutton", undefined, "Sequential from start page:");
            ui.m_startField = modeRow.add("edittext", undefined, "1"); ui.m_startField.characters = 5;
            ui.m_modeSeq.value = true;

            ui.m_modeCustom = ui.m_pdfPanel.add("radiobutton", undefined, "Custom page per Smart Object");

            ui.m_customPanel = ui.m_pdfPanel.add("panel");
            ui.m_customPanel.alignChildren = ["left", "top"];
            ui.m_customPanel.margins = 8; ui.m_customPanel.spacing = 3;
            ui.m_customFields = [];
            ui.m_customLabels = [];
            for (var i = 0; i < n; i++) {
                var crow = ui.m_customPanel.add("group");
                crow.add("statictext", undefined, "Page:");
                var pf = crow.add("edittext", undefined, String(i + 1)); pf.characters = 4;
                var lbl = crow.add("statictext", undefined, ""); lbl.preferredSize = [360, 20];
                ui.m_customFields.push(pf);
                ui.m_customLabels.push(lbl);
            }
            ui.m_customPanel.visible = false;
            ui.m_pdfPanel.visible = false;

            // Refresh custom-mapping labels from the (possibly re-ordered) SOs.
            ui._updateCustomMapping = function () {
                var ord = CONFIG.orderedSOs.length ? CONFIG.orderedSOs : CONFIG.soLayers;
                for (var k = 0; k < ui.m_customLabels.length && k < ord.length; k++) {
                    ui.m_customLabels[k].text = "\u2192  " + ord[k].name;
                }
            };

            function toggleMultiSource() {
                var isPdf = ui.m_radioPdf.value;
                ui.m_pdfPanel.visible = isPdf;
                ui.m_folderStatus.visible = !isPdf;
                win.layout.layout(true);
            }
            ui.m_radioFolder.onClick = ui.m_radioPdf.onClick = function () {
                ui.m_path.text = "";
                CONFIG.sourceFolder = null; CONFIG.sourceFile = null; CONFIG.imageFiles = [];
                ui.m_folderStatus.text = "";
                toggleMultiSource();
            };
            ui.m_modeSeq.onClick = ui.m_modeCustom.onClick = function () {
                var custom = ui.m_modeCustom.value;
                ui.m_customPanel.visible = custom;
                ui.m_startField.enabled = !custom;
                win.layout.layout(true);
            };
            ui.m_browseBtn.onClick = function () {
                if (ui.m_radioPdf.value) {
                    var f = File.openDialog("Select a multi-page PDF", "PDF:*.pdf");
                    if (f) {
                        CONFIG.sourceFile = f;
                        ui.m_path.text = f.fsName;
                        var c = getPDFPageCount(f);
                        CONFIG.pdfPageCount = c;
                        ui.m_pdfCountTxt.text = "PDF page count: " +
                            (c > 0 ? c : "unknown \u2014 enter the pages manually");
                        ui._updateCustomMapping();
                    }
                } else {
                    var fol = Folder.selectDialog("Select a folder of images");
                    if (fol) {
                        CONFIG.sourceFolder = fol;
                        ui.m_path.text = fol.fsName;
                        var files = fol.getFiles(function (ff) {
                            if (ff instanceof Folder) { return false; }
                            return /\.(jpg|jpeg|png|tif|tiff|psd)$/i.test(ff.name);
                        });
                        files.sort(function (a, b) { return naturalCompare(a.name, b.name); });
                        CONFIG.imageFiles = files;
                        ui.m_folderStatus.text = "Found " + files.length +
                            " image(s) \u2014 assigned to Smart Objects in order.";
                    }
                }
            };
        }
    }

    // -------------------------------------------------------------------------
    //  STEP 5  -  Review & Process
    // -------------------------------------------------------------------------
    function buildStep5() {
        g5.add("statictext", undefined, "Review your settings, then start processing.");
        ui.summary = g5.add("statictext", undefined, "", { multiline: true });
        ui.summary.preferredSize = [470, 170];

        ui.startBtn = g5.add("button", undefined, "\u25B6  Start Processing");
        ui.progress = g5.add("progressbar", undefined, 0, 100);
        ui.progress.preferredSize = [470, 16];
        ui.status = g5.add("statictext", undefined, "");
        ui.status.preferredSize = [470, 22];

        ui.startBtn.onClick = onStartProcessing;
    }

    function refreshSummary() {
        var s = "";
        s += "Filename:        " + CONFIG.filename + "   (" + CONFIG.filename + ".psd + " + CONFIG.filename + ".jpg)\n";
        s += "Working folder:  " + (CONFIG.workingDir ? CONFIG.workingDir.fsName : "(none)") + "\n";
        s += "Smart Objects:   " + CONFIG.orderedSOs.length + "  (processed in this order)\n";
        for (var i = 0; i < CONFIG.orderedSOs.length; i++) {
            s += "      " + (i + 1) + ". " + CONFIG.orderedSOs[i].name + "\n";
        }
        s += "Bleed (" + CONFIG.bleed.unit + "):    Top " + CONFIG.bleed.top +
             "  |  Bottom " + CONFIG.bleed.bottom +
             "  |  Left " + CONFIG.bleed.left +
             "  |  Right " + CONFIG.bleed.right + "\n";
        s += "Artwork source:  " + sourceDescription() + "\n";

        var artN = artworkCount(), soN = CONFIG.orderedSOs.length;
        if (artN !== soN) {
            s += "\n\u26A0  " + artN + " artwork item(s) vs " + soN +
                 " Smart Object(s). You'll be asked to confirm before processing.";
        }
        ui.summary.text = s;
    }

    // =========================================================================
    //  PROCESSING DRIVER
    // =========================================================================
    function onStartProcessing() {
        var soN = CONFIG.orderedSOs.length;
        var artN = artworkCount();

        // Mismatch handling: confirm or cancel.
        if (artN !== soN) {
            var msg = "The number of artwork items (" + artN + ") does not match the number " +
                      "of Smart Objects (" + soN + ").\n\n" +
                      "Click OK to continue and fill what we can, or Cancel to go back and adjust.";
            if (!confirm(msg)) { return; }
        }

        // Lock the UI during processing.
        ui.startBtn.enabled = false;
        nextBtn.enabled = false;
        backBtn.enabled = false;
        cancelBtn.enabled = false;

        var oldUnits = app.preferences.rulerUnits;
        var oldDialogs = app.displayDialogs;
        app.displayDialogs = DialogModes.NO;

        var processed = 0, skipped = 0, errors = [];

        try {
            var total = soN;
            for (var i = 0; i < total; i++) {
                var so = CONFIG.orderedSOs[i];
                ui.status.text = "Processing " + (i + 1) + " of " + total + ":  " + so.name;
                ui.progress.value = Math.round((i / total) * 90);
                win.update();

                var art = getArtworkForIndex(i);
                if (!art) { skipped++; continue; }

                try {
                    if (!art.file.exists) { throw new Error("File not found: " + art.file.fsName); }

                    app.activeDocument = CONFIG.parentDoc;
                    CONFIG.parentDoc.activeLayer = so.layer;

                    var psb = openSmartObjectContents(so.layer, CONFIG.parentDoc);
                    cleanSlate(psb);
                    var placed = placeArtwork(psb, art.file, art.isPdf, art.page);
                    scaleAndCenter(psb, placed);
                    psb.flatten();
                    try { psb.save(); } catch (eSave) { /* close still updates parent */ }
                    psb.close(SaveOptions.SAVECHANGES);

                    app.activeDocument = CONFIG.parentDoc;
                    processed++;
                } catch (eSO) {
                    errors.push("\u2022 " + so.name + ":  " + eSO.message);
                    try { app.activeDocument = CONFIG.parentDoc; } catch (e2) {}
                }
            }

            // Export.
            ui.status.text = "Exporting PSD and JPG\u2026";
            ui.progress.value = 95;
            win.update();
            var outFiles = exportFiles();

            ui.progress.value = 100;
            ui.status.text = "Complete.";
            win.update();

            var rep = "Processing complete!\n\n";
            rep += "Smart Objects filled:  " + processed + "\n";
            if (skipped > 0) { rep += "Skipped (no artwork):  " + skipped + "\n"; }
            if (errors.length > 0) {
                rep += "Errors:  " + errors.length + "\n" + errors.join("\n") + "\n";
            }
            rep += "\nSaved PSD:  " + outFiles.psd.fsName +
                   "\nSaved JPG:  " + outFiles.jpg.fsName;
            rep += "\n\nBoth files remain open. You can now crop the JPG to 500\u00d7500 px.";
            alert(rep);
            win.close(1);
        } catch (eFatal) {
            alert("A fatal error occurred during processing:\n\n" + eFatal.message);
            ui.startBtn.enabled = true;
            backBtn.enabled = true;
            cancelBtn.enabled = true;
        } finally {
            app.preferences.rulerUnits = oldUnits;
            app.displayDialogs = oldDialogs;
        }
    }

    // =========================================================================
    //  WIZARD NAVIGATION & VALIDATION
    // =========================================================================
    var STEP_TITLES = {
        1: "Project Setup",
        2: "Smart Object Discovery & Ordering",
        3: "Bleed Settings",
        4: "Artwork Source",
        5: "Review & Process"
    };
    var STEP_SUBS = {
        1: "Choose a base filename and the folder where the PSD and JPG will be saved.",
        2: "Confirm the Smart Objects detected in the template and set their fill order.",
        3: "Optionally extend the artwork beyond the canvas on each side for print bleed.",
        4: "Pick the artwork that will be placed into the Smart Object(s).",
        5: "Confirm everything, then run the automated placement and export."
    };
    var steps = [null, g1, g2, g3, g4, g5];
    var currentStep = 1;

    function showStep(n) {
        currentStep = n;
        for (var i = 1; i <= 5; i++) { steps[i].visible = (i === n); }
        titleTxt.text = "Step " + n + " of 5   \u2014   " + STEP_TITLES[n];
        subTxt.text = STEP_SUBS[n];
        backBtn.enabled = (n > 1);
        nextBtn.visible = (n < 5);
        if (n === 4 && !CONFIG.isSingle && ui._updateCustomMapping) { ui._updateCustomMapping(); }
        if (n === 5) { refreshSummary(); }
        win.layout.layout(true);
    }

    function parseBleedField(t) {
        t = trimStr(t);
        if (t === "") { return 0; }
        var v = parseFloat(t);
        if (isNaN(v) || v < 0) { return NaN; }
        return v;
    }

    function validateStep(n) {
        if (n === 1) {
            var fn = sanitizeName(ui.fileName.text);
            if (fn === "") { alert("Please enter a base filename."); return false; }
            if (!CONFIG.workingDir) { alert("Please choose a working directory."); return false; }
            CONFIG.filename = fn;
            return true;
        }
        if (n === 2) {
            if (CONFIG.isSingle) { CONFIG.orderedSOs = CONFIG.soLayers.slice(); return true; }
            var arr = [];
            for (var i = 0; i < CONFIG.soLayers.length; i++) {
                var v = parseInt(trimStr(ui.orderFields[i].text), 10);
                if (isNaN(v) || v < 1) {
                    alert("Please enter a valid order number (1 or greater) for every Smart Object.");
                    return false;
                }
                CONFIG.soLayers[i].order = v;
                arr.push(CONFIG.soLayers[i]);
            }
            arr.sort(function (a, b) { return (a.order - b.order) || (a.idx - b.idx); });
            CONFIG.orderedSOs = arr;
            return true;
        }
        if (n === 3) {
            CONFIG.bleed.unit = ui.unitDD.selection.text;
            var top = parseBleedField(ui.bTop.text);
            var bottom = parseBleedField(ui.bBottom.text);
            var left = parseBleedField(ui.bLeft.text);
            var right = parseBleedField(ui.bRight.text);
            if (isNaN(top) || isNaN(bottom) || isNaN(left) || isNaN(right)) {
                alert("Bleed values must be non-negative numbers (or blank for 0).");
                return false;
            }
            CONFIG.bleed.top = top; CONFIG.bleed.bottom = bottom;
            CONFIG.bleed.left = left; CONFIG.bleed.right = right;
            return true;
        }
        if (n === 4) {
            if (CONFIG.isSingle) {
                if (!CONFIG.sourceFile) { alert("Please select an artwork file."); return false; }
                if (!CONFIG.sourceFile.exists) { alert("The selected file no longer exists."); return false; }
                if (ui.s_radioPdf.value) {
                    var p = parseInt(trimStr(ui.s_pageField.text), 10);
                    if (isNaN(p) || p < 1) { alert("Please enter a valid PDF page number."); return false; }
                    if (CONFIG.pdfPageCount > 0 && p > CONFIG.pdfPageCount) {
                        alert("Page " + p + " exceeds the PDF page count (" + CONFIG.pdfPageCount + ")."); return false;
                    }
                    CONFIG.sourceType = "pdfFile";
                    CONFIG.singlePdfPage = p;
                } else {
                    CONFIG.sourceType = "imageFile";
                }
                return true;
            } else {
                if (ui.m_radioPdf.value) {
                    if (!CONFIG.sourceFile) { alert("Please select a PDF file."); return false; }
                    if (!CONFIG.sourceFile.exists) { alert("The selected PDF no longer exists."); return false; }
                    if (ui.m_modeCustom.value) {
                        var map = [];
                        for (var j = 0; j < ui.m_customFields.length; j++) {
                            var pv = parseInt(trimStr(ui.m_customFields[j].text), 10);
                            if (isNaN(pv) || pv < 1) {
                                alert("Please enter a valid page number for every Smart Object."); return false;
                            }
                            if (CONFIG.pdfPageCount > 0 && pv > CONFIG.pdfPageCount) {
                                alert("Page " + pv + " exceeds the PDF page count (" + CONFIG.pdfPageCount + ")."); return false;
                            }
                            map.push(pv);
                        }
                        CONFIG.pdfMode = "custom"; CONFIG.pdfMapping = map;
                    } else {
                        var sp = parseInt(trimStr(ui.m_startField.text), 10);
                        if (isNaN(sp) || sp < 1) { alert("Please enter a valid start page."); return false; }
                        if (CONFIG.pdfPageCount > 0 && sp > CONFIG.pdfPageCount) {
                            alert("Start page exceeds the PDF page count (" + CONFIG.pdfPageCount + ")."); return false;
                        }
                        CONFIG.pdfMode = "sequential"; CONFIG.pdfStartPage = sp;
                    }
                    CONFIG.sourceType = "pdfMulti";
                } else {
                    if (!CONFIG.sourceFolder) { alert("Please select a folder of images."); return false; }
                    if (CONFIG.imageFiles.length === 0) {
                        alert("No supported images (JPG/PNG/TIFF/PSD) were found in the selected folder."); return false;
                    }
                    CONFIG.sourceType = "imageFolder";
                }
                return true;
            }
        }
        return true;
    }

    nextBtn.onClick = function () {
        if (!validateStep(currentStep)) { return; }
        if (currentStep < 5) { showStep(currentStep + 1); }
    };
    backBtn.onClick = function () { if (currentStep > 1) { showStep(currentStep - 1); } };
    cancelBtn.onClick = function () { win.close(0); };

    // =========================================================================
    //  ASSEMBLE & LAUNCH
    // =========================================================================
    buildStep2();
    buildStep3();
    buildStep4();
    buildStep5();
    showStep(1);

    win.center();
    win.show();

})();
