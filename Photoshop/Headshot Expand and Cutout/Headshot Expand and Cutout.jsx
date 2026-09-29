#target photoshop
/*
<javascriptresource>
<name>Headshot Expand and Cutout...</name>
<about>Batch Generative Expand on all four sides, with optional Select Subject + Select and Mask background removal.</about>
<category>Headshots</category>
</javascriptresource>
*/

/* =============================================================================
   Headshot Expand and Cutout  v1.2  -  Photoshop ExtendScript (.jsx)
   -----------------------------------------------------------------------------
   v1.2  Cutout without expanding (one-layer images) no longer fails with
         "Merge Visible is not currently available"; JPEG saving no longer
         flattens an already-flat copy.
   Also the processing library for "Headshot Hot Folder.jsx" - keep the two
   scripts in the same folder.

   Per image:
     1. Flatten, convert to RGB 8-bit (Generative Fill needs both), optional
        downsize of the long edge.
     2. Extend the canvas by N% of the original width (left/right) and height
        (top/bottom). 30% on every side = canvas 160% x 160% of the original.
     3. Select the new border (plus a few px of overlap into the photo) and run
        Generative Fill with a blank prompt - the same thing Generative Expand does.
     4. Optional background removal: stamp the result, Select Subject, then
        Select and Mask (radius / smooth / feather / shift edge / decontaminate)
        into a layer mask. New background = transparent, white or a custom colour.
     5. Save as the source format, JPEG, PNG, TIFF or PSD, and write a
        tab-separated log into the destination folder.

   Requirements
     - Photoshop 2024 (v25) or later, signed in and online.
       Photoshop 2026 v27.3.1+ uses the Firefly Fill & Expand model (2K output).
     - Each image = one Generative Fill = normally 1 generative credit, unless
       your plan includes unlimited standard generations (e.g. Creative Cloud Pro).
       Background removal alone uses no credits.
     - Recommended: Preferences > Image Processing > Select Subject and Remove
       Background = Cloud (Detailed Results). Better hair; no credits used.

   Install
     Copy into Photoshop's Presets/Scripts folder and restart Photoshop
     (File > Scripts > Headshot Expand and Cutout...), or use File > Scripts > Browse...

   Notes
     - Esc stops a batch. The current image is abandoned; the log is kept.
     - Generated pixels are rendered at roughly 2048 px across the whole canvas and
       upscaled to fit, so on big files the new border is softer than the original.
     - PSD output keeps the hidden "Generative Expand" layer (with its variations)
       and the original under the Cutout. Switching variation there does NOT
       update the Cutout - it is a stamped copy.
   ============================================================================= */

// ---------------------------------------------------------------------------
// Settings you might want to tweak
// ---------------------------------------------------------------------------
var HEC = {
    NAME: 'Headshot Expand and Cutout',
    VERSION: '1.2',
    SETTINGS_KEY: 'hecHeadshotExpandCutoutSettings_v1',

    // Firefly model for Generative Fill. 'me_md' = Firefly Fill & Expand (Photoshop 27.3+).
    // '' = don't specify; Photoshop uses its current default. If Photoshop rejects the
    // model, the script retries without it and stops asking for it for the rest of the run.
    GEN_MODEL: 'me_md',
    GEN_ATTEMPTS: 2,             // tries per image when a generation fails or returns nothing
    MAX_GEN_FAILS_IN_A_ROW: 3,   // stop the batch after this many consecutive generative failures

    PNG_COMPRESSION: 6,          // 0-9
    // File types picked up from the source folder. Add |dng|cr2|cr3|nef|arw|raf to include
    // raw files (they open with Camera Raw defaults, no dialog).
    FILE_TYPES: /\.(jpe?g|jpe|png|tiff?|psd|psb|webp|heic|heif)$/i,
    LOG_PREFIX: 'HeadshotExpand_log_'
};

// Defaults for the dialog. Last-used values are remembered between runs.
var DEFAULTS = {
    sourceMode: 'folder',        // 'folder' | 'active'
    sourceFolder: '',
    includeSubfolders: false,
    destFolder: '',
    format: 'same',              // 'same' | 'jpg' | 'png' | 'tif' | 'psd'
    jpegQuality: 11,
    suffix: '_expanded',
    skipExisting: true,
    top: 30,
    bottom: 30,
    left: 30,
    right: 30,
    overlap: 4,                  // px the generated border reaches into the original
    maxLongEdge: 0,              // downsize before expanding (0 = off)
    prompt: '',
    removeBg: false,
    smRadius: 3,                 // Select and Mask: edge detection radius (px)
    smSmooth: 0,
    smFeather: 0,                // px
    smShift: -5,                 // shift edge %, negative pulls the edge in
    smDecontaminate: true,
    smDeconAmount: 50,           // %
    bgMode: 'transparent',       // 'transparent' | 'white' | 'colour'
    bgColour: 'FFFFFF'
};

var FORMAT_KEYS = ['same', 'jpg', 'png', 'tif', 'psd'];
var BG_KEYS = ['transparent', 'white', 'colour'];
var GEN_STATE = { modelOK: true };

// ---------------------------------------------------------------------------
// Small utilities (ES3 - ExtendScript has no trim/indexOf on arrays/JSON)
// ---------------------------------------------------------------------------
function sTID(s) { return stringIDToTypeID(s); }
function cTID(s) { return charIDToTypeID(s); }
function trimStr(s) { return String(s).replace(/^\s+|\s+$/g, ''); }
function nowMs() { return new Date().getTime(); }
function secondsSince(t0) { return Math.round((nowMs() - t0) / 100) / 10; }
function pad2(n) { return (n < 10 ? '0' : '') + n; }

function fileTimeStamp() {
    var d = new Date();
    return d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '-' +
        pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds());
}

function clockTime() {
    var d = new Date();
    return pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds());
}

function formatDuration(ms) {
    var s = Math.round(ms / 1000);
    var m = Math.floor(s / 60);
    return m > 0 ? (m + 'm ' + (s % 60) + 's') : (s + 's');
}

function arrayIndex(arr, v) {
    for (var i = 0; i < arr.length; i++) { if (arr[i] === v) return i; }
    return -1;
}

function errMsg(e) {
    if (!e) return 'Unknown error';
    var m = (e.message !== undefined && e.message !== '') ? e.message : String(e);
    return trimStr(String(m).replace(/[\r\n\t]+/g, ' '));
}

function isUserCancel(e) {
    if (!e) return false;
    if (e.number === 8007) return true;
    return /user cancel/i.test(errMsg(e));
}

function isCreditError(e) { return /credit/i.test(errMsg(e)); }

function getExt(name) {
    var m = /\.([^.\/\\]+)$/.exec(String(name));
    return m ? m[1].toLowerCase() : '';
}

function stripExt(name) { return String(name).replace(/\.[^.\/\\]+$/, ''); }

function hexToRgb(hex) {
    var h = trimStr(hex).replace(/^#/, '');
    if (/^[0-9a-f]{3}$/i.test(h)) {
        h = h.charAt(0) + h.charAt(0) + h.charAt(1) + h.charAt(1) + h.charAt(2) + h.charAt(2);
    }
    if (!/^[0-9a-f]{6}$/i.test(h)) return null;
    return { r: parseInt(h.substr(0, 2), 16), g: parseInt(h.substr(2, 2), 16), b: parseInt(h.substr(4, 2), 16) };
}

function cloneSettings(src) {
    var o = {};
    for (var k in src) { if (src.hasOwnProperty(k)) o[k] = src[k]; }
    return o;
}

function serializeSettings(S) {
    var parts = [];
    for (var k in DEFAULTS) {
        if (DEFAULTS.hasOwnProperty(k)) parts.push(k + '=' + encodeURIComponent(String(S[k])));
    }
    return parts.join('&');
}

// Sets one setting from text, coerced to the type of its default. Unknown keys are ignored.
function applySetting(S, k, v) {
    if (!DEFAULTS.hasOwnProperty(k)) return;
    var t = typeof DEFAULTS[k];
    if (t === 'number') {
        var n = parseFloat(v);
        if (!isNaN(n)) S[k] = n;
    } else if (t === 'boolean') {
        S[k] = /^(true|yes|on|1)$/i.test(trimStr(v));
    } else {
        S[k] = String(v);
    }
}

function parseSettings(str, base) {
    var S = cloneSettings(base);
    var parts = String(str || '').split('&');
    for (var i = 0; i < parts.length; i++) {
        var eq = parts[i].indexOf('=');
        if (eq < 1) continue;
        var v;
        try { v = decodeURIComponent(parts[i].substr(eq + 1)); } catch (e) { continue; }
        applySetting(S, parts[i].substr(0, eq), v);
    }
    if (arrayIndex(FORMAT_KEYS, S.format) < 0) S.format = DEFAULTS.format;
    if (arrayIndex(BG_KEYS, S.bgMode) < 0) S.bgMode = DEFAULTS.bgMode;
    if (S.sourceMode !== 'active') S.sourceMode = 'folder';
    return S;
}

// ---- Readable settings files (used by the hot folder queues) ----
// Keys written to a queue's _settings.txt. destFolder blank = the queue's Output folder.
var QUEUE_SETTING_KEYS = ['destFolder', 'format', 'jpegQuality', 'suffix', 'top', 'bottom', 'left', 'right',
    'overlap', 'maxLongEdge', 'prompt', 'removeBg', 'smRadius', 'smSmooth', 'smFeather', 'smShift',
    'smDecontaminate', 'smDeconAmount', 'bgMode', 'bgColour'];

// Same limits as the dialog; applied to hand-edited files.
var SETTING_LIMITS = {
    jpegQuality: [0, 12, true], top: [0, 200], bottom: [0, 200], left: [0, 200], right: [0, 200],
    overlap: [0, 100, true], maxLongEdge: [0, 30000, true], smRadius: [0, 250, true], smSmooth: [0, 100, true],
    smFeather: [0, 250], smShift: [-100, 100], smDeconAmount: [0, 100]
};

function sanitizeSettings(S) {
    for (var k in SETTING_LIMITS) {
        if (!SETTING_LIMITS.hasOwnProperty(k)) continue;
        var lim = SETTING_LIMITS[k];
        var v = Number(S[k]);
        if (isNaN(v)) v = DEFAULTS[k];
        v = Math.min(lim[1], Math.max(lim[0], v));
        S[k] = lim[2] ? Math.round(v) : v;
    }
    if (arrayIndex(FORMAT_KEYS, S.format) < 0) S.format = DEFAULTS.format;
    if (arrayIndex(BG_KEYS, S.bgMode) < 0) S.bgMode = DEFAULTS.bgMode;
    if (S.bgMode === 'colour' && !hexToRgb(S.bgColour)) S.bgColour = 'FFFFFF';
    S.bgColour = String(S.bgColour).replace(/^#/, '').toUpperCase();
    S.suffix = String(S.suffix).replace(/[\\\/:*?"<>|]/g, '');
    return S;
}

function settingsToText(S, keys, headerLines) {
    var lines = headerLines ? headerLines.slice(0) : [];
    for (var i = 0; i < keys.length; i++) {
        lines.push(keys[i] + '=' + String(S[keys[i]]).replace(/[\r\n]+/g, ' '));
    }
    return lines.join('\n') + '\n';
}

function settingsFromText(text, base) {
    var S = cloneSettings(base);
    var lines = String(text || '').replace(/^\uFEFF/, '').split(/\r\n|\r|\n/); // Notepad adds a BOM
    for (var i = 0; i < lines.length; i++) {
        var t = trimStr(lines[i]);
        if (!t || t.charAt(0) === '#') continue;
        var eq = t.indexOf('=');
        if (eq < 1) continue;
        applySetting(S, trimStr(t.substr(0, eq)), trimStr(t.substr(eq + 1)));
    }
    return sanitizeSettings(S);
}

function readTextFile(file) {
    file.encoding = 'UTF-8';
    if (!file.open('r')) throw new Error('Cannot read ' + file.fsName);
    try { return file.read(); } finally { file.close(); }
}

function writeTextFile(file, text) {
    file.encoding = 'UTF-8';
    if (!file.open('w')) throw new Error('Cannot write ' + file.fsName);
    try { file.write(text); } finally { file.close(); }
}

function loadSettings() {
    try {
        var d = app.getCustomOptions(HEC.SETTINGS_KEY);
        return parseSettings(d.getString(sTID('hecSettings')), DEFAULTS);
    } catch (e) {
        return cloneSettings(DEFAULTS);
    }
}

function saveSettings(S) {
    try {
        var d = new ActionDescriptor();
        d.putString(sTID('hecSettings'), serializeSettings(S));
        app.putCustomOptions(HEC.SETTINGS_KEY, d, true);
    } catch (e) { /* not critical */ }
}

// ---------------------------------------------------------------------------
// Geometry and naming (pure functions)
// ---------------------------------------------------------------------------
// Pixels added to each side. Left/right are % of the original width,
// top/bottom % of the original height, so equal values keep the aspect ratio.
function computeExpansion(w, h, S) {
    var ex = {
        origW: w,
        origH: h,
        top: Math.max(0, Math.round(h * S.top / 100)),
        bottom: Math.max(0, Math.round(h * S.bottom / 100)),
        left: Math.max(0, Math.round(w * S.left / 100)),
        right: Math.max(0, Math.round(w * S.right / 100))
    };
    ex.newW = w + ex.left + ex.right;
    ex.newH = h + ex.top + ex.bottom;
    ex.any = (ex.top + ex.bottom + ex.left + ex.right) > 0;
    return ex;
}

// Rectangle [left, top, right, bottom] in the expanded canvas that is KEPT.
// Expanded sides are pulled in by `overlap` px so the generated border blends over
// the original edge; sides that weren't expanded run to the canvas edge.
function computeKeepRect(ex, overlap) {
    var ov = Math.max(0, Math.round(overlap || 0));
    ov = Math.min(ov, Math.floor(Math.min(ex.origW, ex.origH) / 4));
    return [
        ex.left > 0 ? ex.left + ov : 0,
        ex.top > 0 ? ex.top + ov : 0,
        ex.right > 0 ? ex.left + ex.origW - ov : ex.newW,
        ex.bottom > 0 ? ex.top + ex.origH - ov : ex.newH
    ];
}

function formatFromExt(ext) {
    if (ext === 'jpg' || ext === 'jpeg' || ext === 'jpe') return 'jpg';
    if (ext === 'png') return 'png';
    if (ext === 'tif' || ext === 'tiff') return 'tif';
    if (ext === 'psd' || ext === 'psb') return 'psd';
    return '';
}

// Works out the output format for one source file.
function resolveFormat(S, srcExt) {
    var notes = [];
    var fmt = S.format;
    if (fmt === 'same') {
        fmt = formatFromExt(srcExt);
        if (!fmt) {
            fmt = 'jpg';
            notes.push((srcExt ? '.' + srcExt : 'unsaved document') + ' -> JPEG');
        } else if (srcExt === 'psb') {
            notes.push('PSB saved as PSD');
        }
    }
    if (S.removeBg && S.bgMode === 'transparent' && fmt === 'jpg') {
        fmt = 'png';
        notes.push('transparent background - saved as PNG instead of JPEG');
    }
    return { fmt: fmt, note: notes.join('; ') };
}

// Sub-path of `folder` below `root` (URI form, '' if not below it).
function relativeDir(root, folder) {
    var r = String(root.fullName);
    var f = String(folder.fullName);
    if (f.length > r.length && f.toLowerCase().indexOf(r.toLowerCase() + '/') === 0) {
        return f.substr(r.length + 1);
    }
    return '';
}

function samePath(a, b) {
    return String(a.fsName).toLowerCase() === String(b.fsName).toLowerCase();
}

// ---------------------------------------------------------------------------
// Files, log and progress
// ---------------------------------------------------------------------------
function collectFiles(folder, recurse, exclude, out) {
    out = out || [];
    var items = folder.getFiles();
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        var nm = File.decode(it.name);
        if (nm.charAt(0) === '.') continue;
        if (it instanceof Folder) {
            if (recurse && !(exclude && samePath(it, exclude))) collectFiles(it, recurse, exclude, out);
        } else if (HEC.FILE_TYPES.test(nm)) {
            out.push(it);
        }
    }
    return out;
}

function sortFiles(files) {
    files.sort(function (a, b) {
        var x = String(a.fsName).toLowerCase();
        var y = String(b.fsName).toLowerCase();
        return x < y ? -1 : (x > y ? 1 : 0);
    });
    return files;
}

// Creates the folder and any missing parents.
function ensureFolder(folder, depth) {
    if (folder.exists) return;
    depth = depth || 0;
    var parent = folder.parent;
    if (parent && !parent.exists && depth < 32 && String(parent.fsName) !== String(folder.fsName)) {
        ensureFolder(parent, depth + 1);
    }
    if (!folder.create()) throw new Error('Cannot create folder ' + folder.fsName);
}

function findOpenDocument(file) {
    for (var i = 0; i < app.documents.length; i++) {
        try {
            if (samePath(app.documents[i].fullName, file)) return app.documents[i];
        } catch (e) { /* unsaved document */ }
    }
    return null;
}

function Logger(folder) {
    this.file = null;
    if (folder) {
        this.file = new File(folder.fullName + '/' + HEC.LOG_PREFIX + fileTimeStamp() + '.txt');
        this.file.encoding = 'UTF-8';
    }
}
Logger.prototype.line = function (s) {
    if (!this.file) return;
    try {
        if (this.file.open('a')) {
            this.file.writeln(s);
            this.file.close();
        }
    } catch (e) { /* logging must never stop the batch */ }
};
Logger.prototype.row = function (status, src, out, secs, note) {
    this.line([clockTime(), status, src ? src.fsName : '', out ? out.fsName : '',
        secs === '' ? '' : secs + 's', note || ''].join('\t'));
};

function Progress(total) {
    this.win = null;
    try {
        var w = new Window('palette', HEC.NAME);
        w.orientation = 'column';
        w.alignChildren = ['fill', 'top'];
        w.margins = 14;
        this.label = w.add('statictext', undefined, 'Preparing...');
        this.label.preferredSize.width = 460;
        this.bar = w.add('progressbar', undefined, 0, Math.max(1, total));
        this.bar.preferredSize.width = 460;
        w.add('statictext', undefined, 'Esc stops the batch (the current image is abandoned).');
        w.show();
        this.win = w;
    } catch (e) { this.win = null; }
}
Progress.prototype.set = function (value, text) {
    if (!this.win) return;
    try {
        this.label.text = text;
        this.bar.value = value;
        this.win.update();
    } catch (e) { /* cosmetic only */ }
};
Progress.prototype.close = function () {
    if (this.win) { try { this.win.close(); } catch (e) { } }
};

// ---------------------------------------------------------------------------
// Action Manager helpers
// ---------------------------------------------------------------------------
function getActiveDocID() {
    var r = new ActionReference();
    r.putProperty(sTID('property'), sTID('documentID'));
    r.putEnumerated(sTID('document'), sTID('ordinal'), sTID('targetEnum'));
    return executeActionGet(r).getInteger(sTID('documentID'));
}

function getActiveLayerID() {
    var r = new ActionReference();
    r.putProperty(sTID('property'), sTID('layerID'));
    r.putEnumerated(sTID('layer'), sTID('ordinal'), sTID('targetEnum'));
    return executeActionGet(r).getInteger(sTID('layerID'));
}

function layerExists(id) {
    try {
        var r = new ActionReference();
        r.putProperty(sTID('property'), sTID('layerID'));
        r.putIdentifier(sTID('layer'), id);
        executeActionGet(r);
        return true;
    } catch (e) { return false; }
}

function layerHasUserMask(id) {
    try {
        var r = new ActionReference();
        r.putProperty(sTID('property'), sTID('hasUserMask'));
        r.putIdentifier(sTID('layer'), id);
        return executeActionGet(r).getBoolean(sTID('hasUserMask'));
    } catch (e) { return false; }
}

function selectLayerByID(id) {
    var d = new ActionDescriptor();
    var r = new ActionReference();
    r.putIdentifier(sTID('layer'), id);
    d.putReference(sTID('null'), r);
    d.putBoolean(sTID('makeVisible'), false);
    executeAction(sTID('select'), d, DialogModes.NO);
}

function setLayerVisible(id, visible) {
    var d = new ActionDescriptor();
    var list = new ActionList();
    var r = new ActionReference();
    r.putIdentifier(sTID('layer'), id);
    list.putReference(r);
    d.putList(sTID('null'), list);
    executeAction(sTID(visible ? 'show' : 'hide'), d, DialogModes.NO);
}

function deleteLayerByID(id) {
    var d = new ActionDescriptor();
    var r = new ActionReference();
    r.putIdentifier(sTID('layer'), id);
    d.putReference(sTID('null'), r);
    executeAction(sTID('delete'), d, DialogModes.NO);
}

function renameActiveLayer(doc, name) {
    try { doc.activeLayer.name = name; } catch (e) { /* cosmetic only */ }
}

function hasSelection(doc) {
    try {
        var b = doc.selection.bounds;
        return !!(b && b.length === 4);
    } catch (e) { return false; }
}

function deselect(doc) {
    try { doc.selection.deselect(); } catch (e) { }
}

// Stamp Visible (Cmd/Ctrl+Opt/Alt+Shift+E): new merged layer above the active layer.
// Photoshop only allows it with two or more visible layers and a visible active layer.
function stampVisible() {
    var d = new ActionDescriptor();
    d.putBoolean(sTID('duplicate'), true);
    executeAction(sTID('mergeVisible'), d, DialogModes.NO);
}

// Layer > Duplicate Layer on the active layer; the copy becomes the active layer.
function duplicateActiveLayer(name) {
    var d = new ActionDescriptor();
    var r = new ActionReference();
    r.putEnumerated(sTID('layer'), sTID('ordinal'), sTID('targetEnum'));
    d.putReference(sTID('null'), r);
    d.putString(sTID('name'), name);
    executeAction(sTID('duplicate'), d, DialogModes.NO);
}

// True when the document is already a single Background layer ("Flatten Image" is
// not available then).
function isFlat(doc) {
    return doc.layers.length === 1 && doc.layers[0].typename === 'ArtLayer' && doc.layers[0].isBackgroundLayer;
}

// A new top layer holding everything visible, named Cutout. Returns its ID.
// One layer (cutout without expanding): duplicate it - Stamp Visible needs two.
// Otherwise stamp; if Photoshop refuses, copy in a merged duplicate of the document.
function makeCompositeLayer(doc, origID, genID) {
    if (!genID) {
        selectLayerByID(origID);
        duplicateActiveLayer('Cutout');
        return getActiveLayerID();
    }
    selectLayerByID(genID);
    try {
        stampVisible();
        var id = getActiveLayerID();
        if (id !== genID) {
            renameActiveLayer(doc, 'Cutout');
            return id;
        }
    } catch (e) {
        if (isUserCancel(e)) throw e;
    }
    var tmp = doc.duplicate(stripExt(doc.name) + '_merged', true);
    try {
        app.activeDocument = tmp;
        tmp.layers[0].duplicate(doc, ElementPlacement.PLACEATBEGINNING);
    } finally {
        tmp.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = doc;
    }
    doc.activeLayer = doc.layers[0];
    renameActiveLayer(doc, 'Cutout');
    return getActiveLayerID();
}

// Select > Subject on the active layer. Cloud or device processing follows
// Preferences > Image Processing.
function selectSubject() {
    var d = new ActionDescriptor();
    d.putBoolean(sTID('sampleAllLayers'), false);
    executeAction(sTID('autoCutout'), d, DialogModes.NO);
}

// Select and Mask on the current selection, no dialog. `output` is one of
// selectionOutputToSelection / selectionOutputToUserMask /
// selectionOutputToNewLayerWithUserMask. Decontaminate needs a new-layer output.
function selectAndMask(S, decontaminate, output) {
    var d = new ActionDescriptor();
    d.putInteger(sTID('smartBrushRadius'), Math.round(S.smRadius));
    d.putInteger(sTID('smartBrushSmooth'), Math.round(S.smSmooth));
    d.putUnitDouble(sTID('smartBrushFeather'), cTID('#Pxl'), S.smFeather);
    d.putUnitDouble(sTID('smartBrushContrast'), cTID('#Prc'), 0);
    d.putUnitDouble(sTID('smartBrushShiftEdge'), cTID('#Prc'), S.smShift);
    d.putBoolean(sTID('sampleAllLayers'), false);
    d.putBoolean(sTID('smartBrushUseSmartRadius'), S.smRadius > 0);
    d.putBoolean(sTID('smartBrushDecontaminate'), !!decontaminate);
    if (decontaminate) d.putUnitDouble(sTID('smartBrushDeconAmount'), cTID('#Prc'), S.smDeconAmount);
    d.putEnumerated(sTID('refineEdgeOutput'), sTID('refineEdgeOutput'), sTID(output));
    executeAction(sTID('smartBrushWorkspace'), d, DialogModes.NO);
}

// Layer mask on the active layer from the current selection (Reveal Selection).
function makeMaskFromSelection() {
    var d = new ActionDescriptor();
    d.putClass(sTID('new'), sTID('channel'));
    var r = new ActionReference();
    r.putEnumerated(sTID('channel'), sTID('channel'), sTID('mask'));
    d.putReference(sTID('at'), r);
    d.putEnumerated(sTID('using'), sTID('userMaskEnabled'), sTID('revealSelection'));
    executeAction(sTID('make'), d, DialogModes.NO);
}

// Solid Color fill layer, created above the active layer.
function addSolidColourLayer(rgb, name) {
    var d = new ActionDescriptor();
    var r = new ActionReference();
    r.putClass(sTID('contentLayer'));
    d.putReference(sTID('null'), r);
    var layerDesc = new ActionDescriptor();
    layerDesc.putString(sTID('name'), name);
    var typeDesc = new ActionDescriptor();
    var c = new ActionDescriptor();
    c.putDouble(cTID('Rd  '), rgb.r);
    c.putDouble(cTID('Grn '), rgb.g);
    c.putDouble(cTID('Bl  '), rgb.b);
    typeDesc.putObject(sTID('color'), sTID('RGBColor'), c);
    layerDesc.putObject(sTID('type'), sTID('solidColorLayer'), typeDesc);
    d.putObject(sTID('using'), sTID('contentLayer'), layerDesc);
    executeAction(sTID('make'), d, DialogModes.NO);
}

// Generative Fill descriptor, as recorded by ScriptListener in Photoshop 2024-2026.
// A blank prompt on an empty border = Generative Expand.
function buildGenFillDesc(prompt, model, docID, layerID) {
    var d = new ActionDescriptor();
    var r = new ActionReference();
    r.putEnumerated(sTID('document'), sTID('ordinal'), sTID('targetEnum'));
    d.putReference(sTID('null'), r);
    d.putInteger(sTID('documentID'), docID);
    d.putInteger(sTID('layerID'), layerID);
    d.putString(sTID('prompt'), prompt);
    d.putString(sTID('serviceID'), 'clio');
    d.putEnumerated(sTID('workflowType'), sTID('genWorkflow'), sTID('in_painting'));
    var clio = new ActionDescriptor();
    clio.putString(sTID('gi_PROMPT'), prompt);
    clio.putString(sTID('gi_MODE'), prompt ? 'tinp' : 'ginp');
    clio.putInteger(sTID('gi_SEED'), -1);
    clio.putInteger(sTID('gi_NUM_STEPS'), -1);
    clio.putInteger(sTID('gi_GUIDANCE'), 6);
    clio.putInteger(sTID('gi_SIMILARITY'), 0);
    clio.putBoolean(sTID('gi_CROP'), false);
    clio.putBoolean(sTID('gi_DILATE'), false);
    clio.putInteger(sTID('gi_CONTENT_PRESERVE'), 0);
    clio.putBoolean(sTID('gi_ENABLE_PROMPT_FILTER'), true);
    clio.putBoolean(sTID('dualCrop'), true);
    clio.putString(sTID('gi_ADVANCED'), '{"enable_mts":true}');
    var opts = new ActionDescriptor();
    opts.putObject(sTID('clio'), sTID('clio'), clio);
    d.putObject(sTID('serviceOptionsList'), sTID('null'), opts);
    if (model) {
        d.putString(sTID('serviceVersion'), model);
        var map = new ActionDescriptor();
        map.putString(sTID('in_painting'), model);
        map.putString(sTID('out_painting'), model);
        d.putObject(sTID('workflow_to_active_service_identifier_map'), sTID('null'), map);
    }
    return d;
}

function generativeError(e) {
    var err = new Error('Generative Fill failed: ' + errMsg(e));
    err.hecGenerative = true;
    err.hecCredits = isCreditError(e);
    return err;
}

// Runs Generative Fill on the current selection. Returns the new layer's ID.
function runGenerativeFill(doc, prompt) {
    var docID = getActiveDocID();
    var baseID = getActiveLayerID();
    var lastErr = null;
    for (var attempt = 1; attempt <= HEC.GEN_ATTEMPTS; attempt++) {
        var models = (HEC.GEN_MODEL && GEN_STATE.modelOK) ? [HEC.GEN_MODEL, ''] : [''];
        for (var m = 0; m < models.length; m++) {
            var threw = false;
            try {
                selectLayerByID(baseID);
                var res = executeAction(sTID('syntheticFill'),
                    buildGenFillDesc(prompt, models[m], docID, baseID), DialogModes.NO);
                var newID = getActiveLayerID();
                if (newID === baseID && res && res.hasKey(sTID('layerID'))) {
                    newID = res.getInteger(sTID('layerID'));
                }
                if (newID !== baseID && layerExists(newID)) {
                    if (m > 0) GEN_STATE.modelOK = false; // model was rejected; plain call worked
                    selectLayerByID(newID);
                    renameActiveLayer(doc, 'Generative Expand');
                    return newID;
                }
                lastErr = new Error('no result returned (service busy, or blocked by the content filter)');
            } catch (e) {
                if (isUserCancel(e)) throw e;
                if (isCreditError(e)) throw generativeError(e);
                threw = true;
                lastErr = e;
            }
            // Photoshop accepted the request but returned nothing: a different model
            // descriptor won't help, go to the next attempt.
            if (!threw) break;
        }
    }
    throw generativeError(lastErr);
}

// ---------------------------------------------------------------------------
// Image steps
// ---------------------------------------------------------------------------
function pxOf(unitValue) { return Math.round(unitValue.as('px')); }

// One unlocked pixel layer, RGB, 8 bits/channel. Returns that layer's ID.
function prepareDocument(doc, notes) {
    var needFlatten = doc.layers.length > 1;
    if (!needFlatten) {
        var only = doc.layers[0];
        if (only.typename !== 'ArtLayer' || (!only.isBackgroundLayer && only.kind != LayerKind.NORMAL)) {
            needFlatten = true;
        }
    }
    if (needFlatten) {
        doc.flatten();
        notes.push('flattened');
    }
    if (doc.mode != DocumentMode.RGB) {
        var from = String(doc.mode).replace(/^DocumentMode\./, '');
        if (doc.mode == DocumentMode.BITMAP) doc.changeMode(ChangeMode.GRAYSCALE);
        doc.changeMode(ChangeMode.RGB);
        notes.push(from + ' converted to RGB');
    }
    if (doc.bitsPerChannel != BitsPerChannelType.EIGHT) {
        doc.bitsPerChannel = BitsPerChannelType.EIGHT;
        notes.push('converted to 8-bit');
    }
    var lyr = doc.layers[0];
    doc.activeLayer = lyr;
    if (lyr.isBackgroundLayer) lyr.isBackgroundLayer = false;
    try { lyr.allLocked = false; } catch (e) { }
    lyr.name = 'Original';
    return getActiveLayerID();
}

// Adds the new canvas one side at a time so each side can differ.
// The original layer is a normal layer, so the new area is transparent.
function expandCanvas(doc, ex) {
    var w = ex.origW;
    var h = ex.origH;
    if (ex.top > 0) {
        h += ex.top;
        doc.resizeCanvas(UnitValue(w, 'px'), UnitValue(h, 'px'), AnchorPosition.BOTTOMCENTER);
    }
    if (ex.bottom > 0) {
        h += ex.bottom;
        doc.resizeCanvas(UnitValue(w, 'px'), UnitValue(h, 'px'), AnchorPosition.TOPCENTER);
    }
    if (ex.left > 0) {
        w += ex.left;
        doc.resizeCanvas(UnitValue(w, 'px'), UnitValue(h, 'px'), AnchorPosition.MIDDLERIGHT);
    }
    if (ex.right > 0) {
        w += ex.right;
        doc.resizeCanvas(UnitValue(w, 'px'), UnitValue(h, 'px'), AnchorPosition.MIDDLELEFT);
    }
    if (pxOf(doc.width) !== ex.newW || pxOf(doc.height) !== ex.newH) {
        throw new Error('Canvas is ' + pxOf(doc.width) + 'x' + pxOf(doc.height) +
            ', expected ' + ex.newW + 'x' + ex.newH);
    }
}

function selectBorder(doc, keep) {
    doc.selection.selectAll();
    doc.selection.select([[keep[0], keep[1]], [keep[2], keep[1]], [keep[2], keep[3]], [keep[0], keep[3]]],
        SelectionType.DIMINISH, 0, false);
}

// Select and Mask with Decontaminate -> new layer with mask. Returns the cutout layer
// ID, or 0 if this Photoshop version wouldn't do it (caller falls back to a plain mask).
function decontaminatedCutout(doc, S, stampID, notes) {
    var outputs = ['selectionOutputToNewLayerWithUserMask', 'selectionOutputToNewLayerWithLayerMask'];
    for (var i = 0; i < outputs.length; i++) {
        try {
            selectLayerByID(stampID);
            selectAndMask(S, true, outputs[i]);
        } catch (e) {
            if (isUserCancel(e)) throw e;
            continue; // output type not recognised - try the next spelling
        }
        var nowID = getActiveLayerID();
        if (nowID !== stampID && layerHasUserMask(nowID)) {
            try { deleteLayerByID(stampID); } catch (e2) { }
            selectLayerByID(nowID);
            renameActiveLayer(doc, 'Cutout');
            notes.push('cutout, decontaminated ' + S.smDeconAmount + '%');
            return nowID;
        }
        // Photoshop produced something else - tidy up and let the caller redo it plainly.
        if (nowID !== stampID) { try { deleteLayerByID(nowID); } catch (e3) { } }
        if (layerHasUserMask(stampID)) {
            notes.push('cutout (decontaminate not applied)');
            return stampID;
        }
        // Any selection left now may already be refined - drop it so the plain route
        // starts again from a fresh Select Subject instead of refining twice.
        deselect(doc);
        break;
    }
    notes.push('decontaminate unavailable, plain layer mask used');
    return 0;
}

// Select and Mask -> layer mask on the stamp; falls back to the raw Select Subject edge.
function maskedCutout(doc, S, stampID, notes) {
    selectLayerByID(stampID);
    setLayerVisible(stampID, true);
    if (!hasSelection(doc)) selectSubject();
    if (!hasSelection(doc)) throw new Error('Select Subject found no subject');
    try {
        selectAndMask(S, false, 'selectionOutputToUserMask');
        if (!layerHasUserMask(stampID) && hasSelection(doc)) makeMaskFromSelection();
    } catch (e) {
        if (isUserCancel(e)) throw e;
        notes.push('Select and Mask failed (' + errMsg(e) + '), raw Select Subject edge used');
    }
    if (!layerHasUserMask(stampID)) {
        selectLayerByID(stampID);
        if (!hasSelection(doc)) selectSubject();
        if (!hasSelection(doc)) throw new Error('Select Subject found no subject');
        makeMaskFromSelection();
    }
    if (arrayIndex(notes, 'cutout') < 0) notes.push('cutout');
    return stampID;
}

// Stamp -> Select Subject -> Select and Mask -> mask. Hides the source layers and adds
// the new background colour if one was chosen. Returns the cutout layer ID.
function cutOutSubject(doc, S, notes, origID, genID) {
    deselect(doc);
    var stampID = makeCompositeLayer(doc, origID, genID);
    selectSubject();
    if (!hasSelection(doc)) throw new Error('Select Subject found no subject');

    var cutID = 0;
    if (S.smDecontaminate && S.smDeconAmount > 0) cutID = decontaminatedCutout(doc, S, stampID, notes);
    if (!cutID) cutID = maskedCutout(doc, S, stampID, notes);
    deselect(doc);

    setLayerVisible(origID, false);
    if (genID) setLayerVisible(genID, false);

    if (S.bgMode !== 'transparent') {
        var rgb = hexToRgb(S.bgMode === 'white' ? 'FFFFFF' : S.bgColour) || { r: 255, g: 255, b: 255 };
        selectLayerByID(genID || origID); // new layer lands just above it = under the cutout
        addSolidColourLayer(rgb, 'Background colour');
    }
    selectLayerByID(cutID);
    return cutID;
}

// The whole per-image pipeline on an open document. Nothing is saved here.
function processDocument(doc, S, notes) {
    app.activeDocument = doc;
    deselect(doc);
    var origID = prepareDocument(doc, notes);
    var w = pxOf(doc.width);
    var h = pxOf(doc.height);

    if (S.maxLongEdge > 0 && Math.max(w, h) > S.maxLongEdge) {
        var k = S.maxLongEdge / Math.max(w, h);
        var nw = Math.max(1, Math.round(w * k));
        var nh = Math.max(1, Math.round(h * k));
        doc.resizeImage(UnitValue(nw, 'px'), UnitValue(nh, 'px'), doc.resolution, ResampleMethod.BICUBICSHARPER);
        notes.push('downsized ' + w + 'x' + h + ' to ' + nw + 'x' + nh);
        w = nw;
        h = nh;
    }

    var genID = 0;
    var ex = computeExpansion(w, h, S);
    if (ex.any) {
        expandCanvas(doc, ex);
        selectBorder(doc, computeKeepRect(ex, S.overlap));
        selectLayerByID(origID);
        genID = runGenerativeFill(doc, S.prompt);
        deselect(doc);
        notes.push('expanded ' + w + 'x' + h + ' to ' + ex.newW + 'x' + ex.newH);
    }

    if (S.removeBg) cutOutSubject(doc, S, notes, origID, genID);
    return genID;
}

function saveOutput(doc, file, fmt, S) {
    ensureFolder(file.parent);
    app.activeDocument = doc;
    var opts;
    if (fmt === 'jpg') {
        // Save from a flattened duplicate so the working document keeps its layers.
        var tmp = doc.duplicate(stripExt(doc.name) + '_flat', true);
        try {
            if (!isFlat(tmp)) tmp.flatten();
            if (tmp.bitsPerChannel != BitsPerChannelType.EIGHT) tmp.bitsPerChannel = BitsPerChannelType.EIGHT;
            opts = new JPEGSaveOptions();
            opts.quality = S.jpegQuality;
            opts.embedColorProfile = true;
            opts.formatOptions = FormatOptions.STANDARDBASELINE;
            opts.matte = MatteType.NONE;
            tmp.saveAs(file, opts, true, Extension.LOWERCASE);
        } finally {
            tmp.close(SaveOptions.DONOTSAVECHANGES);
            app.activeDocument = doc;
        }
    } else if (fmt === 'png') {
        opts = new PNGSaveOptions();
        opts.compression = HEC.PNG_COMPRESSION;
        opts.interlaced = false;
        doc.saveAs(file, opts, true, Extension.LOWERCASE);
    } else if (fmt === 'tif') {
        opts = new TiffSaveOptions();
        opts.imageCompression = TIFFEncoding.TIFFLZW;
        opts.layers = false;
        opts.transparency = true;
        opts.alphaChannels = false;
        opts.embedColorProfile = true;
        doc.saveAs(file, opts, true, Extension.LOWERCASE);
    } else {
        opts = new PhotoshopSaveOptions();
        opts.layers = true;
        opts.embedColorProfile = true;
        opts.alphaChannels = true;
        doc.saveAs(file, opts, true, Extension.LOWERCASE);
    }
}

// ---------------------------------------------------------------------------
// Runs
// ---------------------------------------------------------------------------
function settingsSummary(S) {
    var s = 'Expand % T/B/L/R: ' + S.top + '/' + S.bottom + '/' + S.left + '/' + S.right +
        ', overlap ' + S.overlap + 'px' +
        (S.maxLongEdge ? ', downsize to ' + S.maxLongEdge + 'px' : '') +
        (S.prompt ? ', prompt "' + S.prompt + '"' : '') +
        ' | Format: ' + S.format + (S.format === 'jpg' || S.format === 'same' ? ' (JPEG q' + S.jpegQuality + ')' : '');
    if (S.removeBg) {
        s += ' | Cutout: radius ' + S.smRadius + ', smooth ' + S.smSmooth + ', feather ' + S.smFeather +
            ', shift ' + S.smShift + '%' + (S.smDecontaminate ? ', decontaminate ' + S.smDeconAmount + '%' : '') +
            ', background ' + (S.bgMode === 'colour' ? '#' + S.bgColour : S.bgMode);
    }
    return s;
}

function runFolder(S) {
    var src = new Folder(S.sourceFolder);
    var dst = new Folder(S.destFolder);
    if (!src.exists) throw new Error('Source folder not found:\n' + S.sourceFolder);
    ensureFolder(dst);

    var files = sortFiles(collectFiles(src, S.includeSubfolders, dst));
    if (!files.length) {
        alert('No supported images found in\n' + src.fsName);
        return;
    }
    var expands = S.top + S.bottom + S.left + S.right > 0;
    var msg = 'Process ' + files.length + ' image(s)?\n\nFrom: ' + src.fsName + '\nTo: ' + dst.fsName;
    if (expands) {
        msg += '\n\nGenerative Expand runs once per image - normally 1 generative credit each, ' +
            'unless your plan includes unlimited standard generations.';
    }
    if (!confirm(msg)) return;

    var log = new Logger(dst);
    log.line(HEC.NAME + ' ' + HEC.VERSION + ' - Photoshop ' + app.version + ' - ' + new Date().toString());
    log.line(settingsSummary(S));
    log.line(['time', 'status', 'source', 'output', 'seconds', 'notes'].join('\t'));

    var stats = { ok: 0, failed: 0, skipped: 0 };
    var usedOutputs = {};
    var genFailsInARow = 0;
    var stopReason = '';
    var tStart = nowMs();
    var prog = new Progress(files.length);

    try {
        // No continue/break inside try/finally below - ExtendScript's engine is old,
        // so the loop is steered with flags instead.
        for (var i = 0; i < files.length && !stopReason; i++) {
            var f = files[i];
            var displayName = File.decode(f.name);
            prog.set(i, (i + 1) + ' of ' + files.length + ':  ' + displayName);
            var t0 = nowMs();
            var doc = null;
            var outFile = null;
            try {
                var srcExt = getExt(displayName);
                var fmtInfo = resolveFormat(S, srcExt);
                var rel = relativeDir(src, f.parent);
                var outDir = rel ? new Folder(dst.fullName + '/' + rel) : dst;
                var base = stripExt(displayName);
                outFile = new File(outDir.fullName + '/' + base + S.suffix + '.' + fmtInfo.fmt);
                if (samePath(outFile, f)) outFile = new File(outDir.fullName + '/' + base + S.suffix + '_expanded.' + fmtInfo.fmt);
                if (usedOutputs[String(outFile.fsName).toLowerCase()]) {
                    // e.g. photo.jpg and photo.png both in the folder
                    outFile = new File(outDir.fullName + '/' + base + '_' + srcExt + S.suffix + '.' + fmtInfo.fmt);
                }
                if (S.skipExisting && outFile.exists) {
                    stats.skipped++;
                    log.row('SKIPPED', f, outFile, '', 'output already exists');
                } else {
                    var notes = [];
                    if (fmtInfo.note) notes.push(fmtInfo.note);

                    var openDoc = findOpenDocument(f);
                    if (openDoc) {
                        doc = openDoc.duplicate(base + S.suffix, false);
                        notes.push('file was open in Photoshop - processed a copy');
                    } else {
                        doc = app.open(f);
                    }
                    processDocument(doc, S, notes);
                    saveOutput(doc, outFile, fmtInfo.fmt, S);
                    usedOutputs[String(outFile.fsName).toLowerCase()] = true;
                    stats.ok++;
                    genFailsInARow = 0;
                    log.row('OK', f, outFile, secondsSince(t0), notes.join('; '));
                }
            } catch (e) {
                if (isUserCancel(e)) {
                    stopReason = 'Stopped with Esc.';
                    log.row('CANCELLED', f, null, secondsSince(t0), '');
                } else {
                    stats.failed++;
                    log.row('FAILED', f, null, secondsSince(t0), errMsg(e));
                    if (e.hecGenerative) {
                        genFailsInARow++;
                        if (e.hecCredits) {
                            stopReason = 'Stopped: generative credits look exhausted.\n' + errMsg(e);
                        } else if (genFailsInARow >= HEC.MAX_GEN_FAILS_IN_A_ROW) {
                            stopReason = 'Stopped after ' + genFailsInARow + ' Generative Fill failures in a row.\n' +
                                'Check you are signed in, online and have credits.\nLast error: ' + errMsg(e);
                        }
                    }
                }
            } finally {
                if (doc) {
                    try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose) { }
                }
            }
        }
    } finally {
        prog.close();
    }

    var summary = 'Saved: ' + stats.ok + '   Skipped: ' + stats.skipped + '   Failed: ' + stats.failed +
        '\nTime: ' + formatDuration(nowMs() - tStart);
    log.line('# ' + summary.replace(/\n/g, '  ') + (stopReason ? '  ' + stopReason.replace(/\n/g, ' ') : ''));
    alert(HEC.NAME + '\n\n' + (stopReason ? stopReason + '\n\n' : '') + summary +
        (log.file ? '\n\nLog: ' + log.file.fsName : ''));
}

function runActive(S) {
    var src = app.activeDocument;
    var srcExt = '';
    try { srcExt = getExt(src.fullName.name); } catch (e) { srcExt = 'psd'; } // unsaved
    var fmtInfo = resolveFormat(S, srcExt);
    var base = stripExt(src.name);
    var notes = [];
    if (fmtInfo.note) notes.push(fmtInfo.note);
    var t0 = nowMs();

    // Work on a copy - the original stays open and untouched.
    var doc = src.duplicate(base + S.suffix, false);
    processDocument(doc, S, notes);

    var saved = '';
    if (S.destFolder) {
        var dst = new Folder(S.destFolder);
        ensureFolder(dst);
        var outFile = new File(dst.fullName + '/' + base + S.suffix + '.' + fmtInfo.fmt);
        saveOutput(doc, outFile, fmtInfo.fmt, S);
        saved = '\n\nSaved: ' + outFile.fsName;
    }
    app.activeDocument = doc;
    alert(HEC.NAME + '\n\nDone in ' + formatDuration(nowMs() - t0) + '.' +
        (notes.length ? '\n' + notes.join('\n') : '') + saved +
        '\n\nThe result is open as "' + doc.name + '" - your original is unchanged.');
}

// ---------------------------------------------------------------------------
// Dialog
// ---------------------------------------------------------------------------
// opts (optional) - { mode: 'queue', queueName: 'Cutout - PNG', isNew: true, validateName: fn }
// Queue mode (used by the hot folder): the Source panel becomes the queue's folder name,
// the destination becomes an optional output folder, and the result carries .queueName.
function showDialog(S, haveDoc, opts) {
    opts = opts || {};
    var isQueue = opts.mode === 'queue';
    var result = null;
    var title = HEC.NAME + '  v' + HEC.VERSION;
    if (isQueue) title = opts.isNew ? 'New hot folder queue' : 'Queue settings: ' + opts.queueName;
    var dlg = new Window('dialog', title);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.spacing = 10;
    dlg.margins = 16;

    function addPanel(title) {
        var p = dlg.add('panel', undefined, title);
        p.orientation = 'column';
        p.alignChildren = ['left', 'top'];
        p.margins = [12, 18, 12, 10];
        p.spacing = 6;
        return p;
    }
    function addRow(parent) {
        var g = parent.add('group');
        g.orientation = 'row';
        g.alignChildren = ['left', 'center'];
        g.spacing = 6;
        return g;
    }
    function addNum(parent, label, value, chars, tip) {
        if (label) parent.add('statictext', undefined, label);
        var et = parent.add('edittext', undefined, String(value));
        et.characters = chars || 4;
        if (tip) et.helpTip = tip;
        return et;
    }

    var rbFolder = null, rbActive = null, etSrc = null, btSrc = null, cbSub = null, cbSkip = null, etName = null;
    if (isQueue) {
        // Queue name = the drop folder people see
        var pQ = addPanel('Queue (drop folder)');
        var gName = addRow(pQ);
        etName = addNum(gName, 'Folder name:', opts.queueName || '', 30,
            'People drop images into a folder with this name inside the hot folder.');
        etName.enabled = !!opts.isNew;
    } else {
        // Source
        var pSrc = addPanel('Source');
        var gMode = addRow(pSrc);
        rbFolder = gMode.add('radiobutton', undefined, 'Folder of images');
        rbActive = gMode.add('radiobutton', undefined, 'Active document (works on a copy)');
        var gSrc = addRow(pSrc);
        etSrc = gSrc.add('edittext', undefined, S.sourceFolder);
        etSrc.preferredSize.width = 400;
        btSrc = gSrc.add('button', undefined, 'Browse...');
        cbSub = pSrc.add('checkbox', undefined, 'Include subfolders (structure is mirrored in the destination)');
        cbSub.value = S.includeSubfolders;
        rbActive.enabled = haveDoc;
        rbActive.value = haveDoc && S.sourceMode === 'active';
        rbFolder.value = !rbActive.value;
    }

    // Destination / Output
    var pDst = addPanel(isQueue ? 'Output' : 'Destination');
    var gDst = addRow(pDst);
    if (isQueue) gDst.add('statictext', undefined, 'Save to:');
    var etDst = gDst.add('edittext', undefined, S.destFolder);
    etDst.preferredSize.width = isQueue ? 340 : 400;
    etDst.helpTip = isQueue ? 'Leave blank to save into the Output folder inside the queue.' :
        'Optional for Active document - leave blank to just leave the result open.';
    var btDst = gDst.add('button', undefined, 'Browse...');
    var gFmt = addRow(pDst);
    gFmt.add('statictext', undefined, 'Format:');
    var ddFmt = gFmt.add('dropdownlist', undefined,
        ['Same as source', 'JPEG', 'PNG', 'TIFF (LZW)', 'PSD (keeps layers and variations)']);
    ddFmt.selection = Math.max(0, arrayIndex(FORMAT_KEYS, S.format));
    var etQ = addNum(gFmt, '   JPEG quality (0-12):', S.jpegQuality, 3);
    var gSuf = addRow(pDst);
    var etSuf = addNum(gSuf, 'Filename suffix:', S.suffix, 12, 'Added to each file name, e.g. photo_expanded.jpg');
    if (!isQueue) {
        cbSkip = gSuf.add('checkbox', undefined, 'Skip files already in the destination');
        cbSkip.value = S.skipExisting;
        cbSkip.helpTip = 'Lets you re-run a stopped batch without paying for the same image twice.';
    } else {
        pDst.add('statictext', undefined, 'Leave "Save to" blank to use the Output folder inside the queue.');
    }

    // Generative Expand
    var pGen = addPanel('Generative Expand');
    var gPct = addRow(pGen);
    gPct.add('statictext', undefined, 'Add to each side (% of original):');
    var etTop = addNum(gPct, 'Top', S.top, 3);
    var etBottom = addNum(gPct, 'Bottom', S.bottom, 3);
    var etLeft = addNum(gPct, 'Left', S.left, 3);
    var etRight = addNum(gPct, 'Right', S.right, 3);
    var gOv = addRow(pGen);
    var etOv = addNum(gOv, 'Overlap into photo (px):', S.overlap, 3,
        'The generated border reaches this far into the original to hide the join.');
    var etMax = addNum(gOv, '   Downsize long edge first (px, 0 = off):', S.maxLongEdge, 5,
        'Generated pixels top out around 2048 px across the whole canvas. Downsizing big files first keeps the border and the face at a similar sharpness.');
    var gPr = addRow(pGen);
    gPr.add('statictext', undefined, 'Prompt (optional):');
    var etPrompt = gPr.add('edittext', undefined, S.prompt);
    etPrompt.preferredSize.width = 330;
    etPrompt.helpTip = 'Leave blank for a natural continuation (recommended for expanding).';
    pGen.add('statictext', undefined, '30 on every side = canvas 160% x 160%. All zeros = no Generative Expand (cutout only, no credits).');

    // Background removal
    var pBg = addPanel('Background removal');
    var cbBg = pBg.add('checkbox', undefined, 'Remove background  (Select Subject, then Select and Mask)');
    cbBg.value = S.removeBg;
    var gSM = addRow(pBg);
    var etRad = addNum(gSM, 'Edge radius (px):', S.smRadius, 3, 'Smart Radius. Raise for large files or wispy hair (e.g. 6-10 on 24 MP).');
    var etSmooth = addNum(gSM, '  Smooth:', S.smSmooth, 3);
    var etFeather = addNum(gSM, '  Feather (px):', S.smFeather, 4);
    var etShift = addNum(gSM, '  Shift edge (%):', S.smShift, 4, 'Negative values pull the edge in and cut halos.');
    var gDec = addRow(pBg);
    var cbDec = gDec.add('checkbox', undefined, 'Decontaminate colours');
    cbDec.value = S.smDecontaminate;
    var etDec = addNum(gDec, 'Amount (%):', S.smDeconAmount, 3, 'Removes the old background colour from hair edges.');
    var gFill = addRow(pBg);
    gFill.add('statictext', undefined, 'New background:');
    var ddBg = gFill.add('dropdownlist', undefined, ['Transparent', 'White', 'Custom colour']);
    ddBg.selection = Math.max(0, arrayIndex(BG_KEYS, S.bgMode));
    var etHex = addNum(gFill, '  Hex:', S.bgColour, 7);
    var btPick = gFill.add('button', undefined, 'Pick...');
    pBg.add('statictext', undefined, 'Transparent needs PNG, TIFF or PSD - JPEG output is switched to PNG automatically.');

    // Buttons
    var gBtn = dlg.add('group');
    gBtn.alignment = ['right', 'top'];
    var btCancel = gBtn.add('button', undefined, 'Cancel', { name: 'cancel' });
    var btRun = gBtn.add('button', undefined, isQueue ? 'Save' : 'Run', { name: 'ok' });

    function refresh() {
        if (rbFolder) {
            var folderMode = rbFolder.value;
            etSrc.enabled = folderMode;
            btSrc.enabled = folderMode;
            cbSub.enabled = folderMode;
        }
        var fmtIdx = ddFmt.selection ? ddFmt.selection.index : 0;
        etQ.enabled = (fmtIdx === 0 || fmtIdx === 1);
        var bg = cbBg.value;
        gSM.enabled = bg;
        gDec.enabled = bg;
        gFill.enabled = bg;
        etDec.enabled = bg && cbDec.value;
        var custom = ddBg.selection && ddBg.selection.index === 2;
        etHex.enabled = bg && custom;
        btPick.enabled = bg && custom;
    }
    if (rbFolder) {
        rbFolder.onClick = refresh;
        rbActive.onClick = refresh;
    }
    ddFmt.onChange = refresh;
    cbBg.onClick = refresh;
    cbDec.onClick = refresh;
    ddBg.onChange = refresh;

    function browse(et, promptText) {
        var start = trimStr(et.text) ? new Folder(trimStr(et.text)) : null;
        var f = (start && start.exists) ? start.selectDlg(promptText) : Folder.selectDialog(promptText);
        if (f) et.text = f.fsName;
    }
    if (btSrc) btSrc.onClick = function () { browse(etSrc, 'Choose the folder of headshots'); };
    btDst.onClick = function () {
        browse(etDst, isQueue ? 'Choose where this queue saves its results' : 'Choose where to save the results');
    };

    btPick.onClick = function () {
        try {
            var prev = app.foregroundColor;
            var rgb = hexToRgb(etHex.text) || { r: 255, g: 255, b: 255 };
            var c = new SolidColor();
            c.rgb.red = rgb.r;
            c.rgb.green = rgb.g;
            c.rgb.blue = rgb.b;
            app.foregroundColor = c;
            if (app.showColorPicker()) etHex.text = app.foregroundColor.rgb.hexValue;
            app.foregroundColor = prev;
        } catch (e) { /* type the hex value instead */ }
    };

    btRun.onClick = function () {
        var errors = [];
        function num(et, label, min, max, isInt) {
            var v = parseFloat(trimStr(et.text));
            if (isNaN(v) || v < min || v > max) {
                errors.push(label + ' must be a number from ' + min + ' to ' + max + '.');
                return min;
            }
            return isInt ? Math.round(v) : v;
        }
        var R = cloneSettings(S);
        if (isQueue) {
            R.queueName = trimStr(etName.text);
        } else {
            R.sourceMode = rbActive.value ? 'active' : 'folder';
            R.sourceFolder = trimStr(etSrc.text);
            R.includeSubfolders = cbSub.value;
            R.skipExisting = cbSkip.value;
        }
        R.destFolder = trimStr(etDst.text);
        R.format = FORMAT_KEYS[ddFmt.selection ? ddFmt.selection.index : 0];
        R.jpegQuality = num(etQ, 'JPEG quality', 0, 12, true);
        R.suffix = trimStr(etSuf.text);
        R.top = num(etTop, 'Top', 0, 200, false);
        R.bottom = num(etBottom, 'Bottom', 0, 200, false);
        R.left = num(etLeft, 'Left', 0, 200, false);
        R.right = num(etRight, 'Right', 0, 200, false);
        R.overlap = num(etOv, 'Overlap', 0, 100, true);
        R.maxLongEdge = num(etMax, 'Downsize long edge', 0, 30000, true);
        R.prompt = trimStr(etPrompt.text);
        R.removeBg = cbBg.value;
        R.smRadius = num(etRad, 'Edge radius', 0, 250, true);
        R.smSmooth = num(etSmooth, 'Smooth', 0, 100, true);
        R.smFeather = num(etFeather, 'Feather', 0, 250, false);
        R.smShift = num(etShift, 'Shift edge', -100, 100, false);
        R.smDecontaminate = cbDec.value;
        R.smDeconAmount = num(etDec, 'Decontaminate amount', 0, 100, false);
        R.bgMode = BG_KEYS[ddBg.selection ? ddBg.selection.index : 0];
        R.bgColour = trimStr(etHex.text).replace(/^#/, '').toUpperCase();

        if (R.maxLongEdge > 0 && R.maxLongEdge < 256) errors.push('Downsize long edge: use 0 (off) or at least 256 px.');
        if (R.removeBg && R.bgMode === 'colour' && !hexToRgb(R.bgColour)) errors.push('Background colour must be a hex value like F2F2F2.');
        if (/[\\\/:*?"<>|]/.test(R.suffix)) errors.push('The filename suffix contains characters that are not allowed in file names.');
        var expands = R.top + R.bottom + R.left + R.right > 0;
        if (!expands && !R.removeBg) errors.push('Nothing to do - set some expansion or tick Remove background.');
        if (isQueue) {
            if (!R.queueName) {
                errors.push('Give the queue a folder name.');
            } else if (/[\\\/:*?"<>|]/.test(R.queueName) || /^[._]/.test(R.queueName)) {
                errors.push('The folder name cannot contain \\ / : * ? " < > | or start with . or _');
            } else if (opts.validateName) {
                var nameErr = opts.validateName(R.queueName);
                if (nameErr) errors.push(nameErr);
            }
        } else if (R.sourceMode === 'folder') {
            if (!R.sourceFolder || !new Folder(R.sourceFolder).exists) errors.push('Choose a source folder that exists.');
            if (!R.destFolder) errors.push('Choose a destination folder.');
            if (R.sourceFolder && R.destFolder && samePath(new Folder(R.sourceFolder), new Folder(R.destFolder)) && !R.suffix) {
                errors.push('Destination is the source folder - add a filename suffix so originals are not overwritten.');
            }
        }
        if (errors.length) {
            alert(errors.join('\n'));
            return;
        }
        result = R;
        dlg.close(1);
    };
    btCancel.onClick = function () { dlg.close(2); };

    refresh();
    dlg.center();
    dlg.show();
    return result;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
function main() {
    if (typeof app === 'undefined' || !/photoshop/i.test(app.name)) {
        alert('Run this script from Adobe Photoshop.');
        return;
    }
    var psMajor = parseInt(app.version, 10);
    var S = loadSettings();
    S = showDialog(S, app.documents.length > 0);
    if (!S) return;
    saveSettings(S);

    if (S.top + S.bottom + S.left + S.right > 0 && psMajor < 25) {
        alert('Generative Fill needs Photoshop 2024 (v25) or later. This is v' + app.version + '.');
        return;
    }

    var saved = {
        rulerUnits: app.preferences.rulerUnits,
        typeUnits: app.preferences.typeUnits,
        dialogs: app.displayDialogs
    };
    app.preferences.rulerUnits = Units.PIXELS;
    app.preferences.typeUnits = TypeUnits.PIXELS;
    app.displayDialogs = DialogModes.NO;
    try {
        if (S.sourceMode === 'active') runActive(S);
        else runFolder(S);
    } catch (e) {
        if (!isUserCancel(e)) alert(HEC.NAME + ' stopped:\n' + errMsg(e) + (e.line ? '\n(line ' + e.line + ')' : ''));
    } finally {
        app.preferences.rulerUnits = saved.rulerUnits;
        app.preferences.typeUnits = saved.typeUnits;
        app.displayDialogs = saved.dialogs;
    }
}

if (typeof HEC_NO_AUTORUN === 'undefined') {
    main();
}
