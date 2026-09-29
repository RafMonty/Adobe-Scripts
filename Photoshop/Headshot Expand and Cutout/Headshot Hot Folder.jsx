#target photoshop
/*
<javascriptresource>
<name>Headshot Hot Folder...</name>
<about>Watches drop folders and runs Headshot Expand and Cutout on anything that lands in them. For a dedicated machine - Photoshop is busy while it watches.</about>
<category>Headshots</category>
</javascriptresource>
*/

/* =============================================================================
   Headshot Hot Folder  v1.0  -  Photoshop ExtendScript (.jsx)
   -----------------------------------------------------------------------------
   Turns Photoshop into a watched-folder ("hot folder") processor for headshots.
   Needs "Headshot Expand and Cutout.jsx" (v1.1+) in the SAME folder - that
   script does the image work; this one does the watching.

   Folder layout (the setup dialog makes it for you):

     Headshots Hot Folder/              <- the folder you point the watcher at
       _status.txt                      <- what the watcher is doing right now
       _logs/                           <- one tab-separated log per day
       Expand - JPEG/                   <- a queue: people drop images in here
         _settings.txt                  <- this queue's settings
         Output/                        <- results
         Originals/                     <- originals are moved here when done
         Failed/                        <- originals that failed + a .txt saying why
       Cutout - transparent PNG/
       Cutout - white JPEG/

   How it behaves
     - Drop image files, or a folder of them, into a queue. A file is picked up
       once it has stopped changing for a few seconds, so half-copied files are
       left alone. Dropped folders are mirrored in Output/Originals/Failed.
     - Stop: put a file called STOP (or STOP.txt) in the hot folder, click Stop or
       close the status window, or press Esc. Unprocessed files just wait.
     - Photoshop is busy for as long as it watches (that is how scripts work),
       so use a machine or a Photoshop nobody needs, and set it not to sleep.
     - Before the first run on that machine, process one image by hand with
       Headshot Expand and Cutout (Active document) so any first-time Generative
       Fill or sign-in prompts are out of the way.
     - Safety for unattended use:
         * daily limit on generations (credits); cutout-only queues are not limited
         * pauses by itself when credits run out or Firefly keeps failing, then retries
         * a Firefly failure is retried after 15 and then 30 more minutes before the
           file goes to Failed, so short outages don't fail anything
         * a file that was being processed when Photoshop died twice goes to Failed
     - To start watching whenever Photoshop starts: tick "start watching straight
       away" in the setup dialog, then File > Scripts > Script Events Manager,
       event "Start Application", script = this file. Hold Shift while it starts
       to get the setup dialog back.
   ============================================================================= */

var HEC_NO_AUTORUN = true; // load the processing library without opening its dialog
var HF_LIB_ERROR = '';
var HF_LIB_FILE = new File(new File($.fileName).parent.fullName + '/Headshot Expand and Cutout.jsx');
if (HF_LIB_FILE.exists) {
    try {
        $.evalFile(HF_LIB_FILE);
    } catch (eLib) {
        HF_LIB_ERROR = 'Could not load ' + HF_LIB_FILE.fsName + '\n' + (eLib.message || eLib);
    }
} else {
    HF_LIB_ERROR = 'This script needs "Headshot Expand and Cutout.jsx" (v1.1 or later) in the same folder:\n' +
        HF_LIB_FILE.parent.fsName;
}

var HF = {
    NAME: 'Headshot Hot Folder',
    VERSION: '1.0',
    OPTIONS_KEY: 'hecHotFolderOptions_v1',
    SETTINGS_FILE: '_settings.txt',
    STATUS_FILE: '_status.txt',
    LOG_FOLDER: '_logs',
    MARKER_FILE: '_in_progress.txt',   // in _logs; left behind only if Photoshop dies mid-file
    IGNORE_FILE: '_ignored.txt',       // in _logs; originals that could not be moved away
    OUTPUT: 'Output',
    ORIGINALS: 'Originals',
    FAILED: 'Failed',
    STOP_NAME: /^stop(\.[a-z0-9]{1,5})?$/i,   // STOP, STOP.txt, stop.rtf ...
    STABLE_SECONDS: 8,         // a file must be unchanged this long before it is picked up
    MAX_TRIES: 3,              // Firefly failures on one file before it goes to Failed
    CRASH_LIMIT: 2,            // Photoshop died on the same file this many times -> Failed
    FAILS_BEFORE_PAUSE: 3,     // Firefly failures in a row (any files) that pause the watcher
    MAX_DEPTH: 4,              // how deep dropped folders are searched
    EMPTY_FOLDER_SECONDS: 120, // dropped folders are removed once empty for this long
    JUNK: /^(thumbs\.db|desktop\.ini|\.ds_store|icon\r?)$/i,
    ARRIVING: /\.(part|partial|crdownload|download|tmp)$/i
};

var HF_DEFAULTS = {
    root: '',
    interval: 10,        // seconds between checks
    dailyLimit: 100,     // expanded images per day, 0 = no limit
    pauseMinutes: 15,
    autoStart: false
};

// ---------------------------------------------------------------------------
// Options (remembered on this machine)
// ---------------------------------------------------------------------------
function hfLoadOptions() {
    var o = cloneSettings(HF_DEFAULTS);
    try {
        var parts = app.getCustomOptions(HF.OPTIONS_KEY).getString(sTID('hecHotFolder')).split('&');
        for (var i = 0; i < parts.length; i++) {
            var eq = parts[i].indexOf('=');
            if (eq < 1) continue;
            var k = parts[i].substr(0, eq);
            if (!HF_DEFAULTS.hasOwnProperty(k)) continue;
            var v = decodeURIComponent(parts[i].substr(eq + 1));
            var t = typeof HF_DEFAULTS[k];
            if (t === 'number') {
                var n = parseFloat(v);
                if (!isNaN(n)) o[k] = n;
            } else if (t === 'boolean') {
                o[k] = (v === 'true');
            } else {
                o[k] = v;
            }
        }
    } catch (e) { /* first run */ }
    return o;
}

function hfSaveOptions(o) {
    try {
        var parts = [];
        for (var k in HF_DEFAULTS) {
            if (HF_DEFAULTS.hasOwnProperty(k)) parts.push(k + '=' + encodeURIComponent(String(o[k])));
        }
        var d = new ActionDescriptor();
        d.putString(sTID('hecHotFolder'), parts.join('&'));
        app.putCustomOptions(HF.OPTIONS_KEY, d, true);
    } catch (e) { /* not critical */ }
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------
function hfDayKey(ms) {
    var d = ms ? new Date(ms) : new Date();
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
}

function hfDateTime(ms) {
    if (!ms) return '-';
    var d = new Date(ms);
    return hfDayKey(ms) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds());
}

function hfTime(ms) {
    var d = new Date(ms);
    return pad2(d.getHours()) + ':' + pad2(d.getMinutes());
}

function hfKey(file) { return String(file.fsName).toLowerCase(); }

function hfSig(file) {
    var m = file.modified;
    return file.length + '|' + (m ? m.getTime() : 0);
}

// "photo.JPG" -> free name in folder: photo.JPG, photo-2.JPG, photo-3.JPG ...
function hfUniqueFile(folder, base, ext) {
    var f = new File(folder.fullName + '/' + base + ext);
    for (var n = 2; f.exists && n < 1000; n++) {
        f = new File(folder.fullName + '/' + base + '-' + n + ext);
    }
    return f;
}

function hfExtOf(name) {
    var base = stripExt(name);
    return String(name).substr(base.length); // keeps the original case, includes the dot
}

// Copy then delete (ExtendScript can only rename within a folder).
function hfMoveFile(file, destFolder, name) {
    ensureFolder(destFolder);
    var target = hfUniqueFile(destFolder, stripExt(name), hfExtOf(name));
    if (!file.copy(target.fsName)) throw new Error('could not copy to ' + target.fsName);
    return { target: target, removed: file.remove() };
}

// ---------------------------------------------------------------------------
// Queues
// ---------------------------------------------------------------------------
function hfDefaultQueueSettings() {
    var S = cloneSettings(DEFAULTS);
    S.destFolder = '';
    S.skipExisting = false;
    return S;
}

function hfStarterQueues() {
    var a = hfDefaultQueueSettings();
    a.format = 'jpg';
    a.suffix = '_expanded';
    var b = hfDefaultQueueSettings();
    b.format = 'png';
    b.suffix = '_cutout';
    b.removeBg = true;
    b.bgMode = 'transparent';
    var c = hfDefaultQueueSettings();
    c.format = 'jpg';
    c.suffix = '_cutout';
    c.removeBg = true;
    c.bgMode = 'white';
    return [
        { name: 'Expand - JPEG', S: a },
        { name: 'Cutout - transparent PNG', S: b },
        { name: 'Cutout - white JPEG', S: c }
    ];
}

function hfFindQueues(root) {
    var list = [];
    if (!root || !root.exists) return list;
    var items = root.getFiles();
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!(it instanceof Folder)) continue;
        var name = File.decode(it.name);
        if (/^[._]/.test(name)) continue;
        var sf = new File(it.fullName + '/' + HF.SETTINGS_FILE);
        if (sf.exists) list.push({ name: name, folder: it, settingsFile: sf, settings: null, error: '' });
    }
    list.sort(function (a, b) {
        var x = a.name.toLowerCase();
        var y = b.name.toLowerCase();
        return x < y ? -1 : (x > y ? 1 : 0);
    });
    return list;
}

function hfReadQueueSettings(q) {
    var S = settingsFromText(readTextFile(q.settingsFile), hfDefaultQueueSettings());
    S.skipExisting = false;
    return S;
}

function hfWriteQueueSettings(folder, S) {
    writeTextFile(new File(folder.fullName + '/' + HF.SETTINGS_FILE), settingsToText(S, QUEUE_SETTING_KEYS, [
        '# ' + HF.NAME + ' - settings for this drop folder.',
        '# Made by the setup dialog. Safe to edit by hand: one key=value per line.',
        '# destFolder: blank = the Output folder inside this queue.',
        '# format: same, jpg, png, tif or psd.   bgMode: transparent, white or colour.'
    ]));
}

function hfCreateQueue(root, name, S) {
    var folder = new Folder(root.fullName + '/' + name);
    ensureFolder(folder);
    ensureFolder(new Folder(folder.fullName + '/' + HF.OUTPUT));
    ensureFolder(new Folder(folder.fullName + '/' + HF.ORIGINALS));
    ensureFolder(new Folder(folder.fullName + '/' + HF.FAILED));
    hfWriteQueueSettings(folder, S);
    return folder;
}

function hfQueueSummary(S) {
    var fmtNames = { same: 'same format', jpg: 'JPEG', png: 'PNG', tif: 'TIFF', psd: 'PSD' };
    var parts = [];
    if (S.top + S.bottom + S.left + S.right > 0) {
        parts.push('expand ' + S.top + '/' + S.bottom + '/' + S.left + '/' + S.right + '%');
    } else {
        parts.push('no expand');
    }
    if (S.removeBg) parts.push('cutout on ' + (S.bgMode === 'colour' ? '#' + S.bgColour : S.bgMode));
    parts.push(fmtNames[S.format] || S.format);
    return parts.join(', ');
}

// ---------------------------------------------------------------------------
// Watching: finding files that are ready
// ---------------------------------------------------------------------------
// A file is ready once its size and date have stayed the same for STABLE_SECONDS
// and it can be opened (Windows keeps files locked while they are being copied).
function HFTracker(seconds) {
    this.seconds = seconds;
    this.seen = {};
    this.round = 0;
}
HFTracker.prototype.begin = function () { this.round++; };
HFTracker.prototype.isReady = function (file, nowT) {
    var k = hfKey(file);
    var sig = hfSig(file);
    var e = this.seen[k];
    if (!e || e.sig !== sig) {
        this.seen[k] = { sig: sig, since: nowT, round: this.round };
        return false;
    }
    e.round = this.round;
    if (file.length <= 0 || nowT - e.since < this.seconds * 1000) return false;
    if (!file.open('r')) return false;
    file.close();
    return true;
};
HFTracker.prototype.end = function () {
    var stale = [];
    for (var k in this.seen) {
        if (this.seen.hasOwnProperty(k) && this.seen[k].round !== this.round) stale.push(k);
    }
    for (var i = 0; i < stale.length; i++) delete this.seen[stale[i]];
};

function hfIsIgnored(ctx, file) {
    var s = ctx.ignore[hfKey(file)];
    return s !== undefined && s === hfSig(file);
}

// Remember an original we could not move away, so it is not processed (and paid for) again.
function hfIgnore(ctx, file) {
    var k = hfKey(file);
    var sig = hfSig(file);
    ctx.ignore[k] = sig;
    try {
        var f = new File(ctx.logDir.fullName + '/' + HF.IGNORE_FILE);
        f.encoding = 'UTF-8';
        if (f.open('a')) {
            f.writeln(k + '\t' + sig);
            f.close();
        }
    } catch (e) { }
}

function hfLoadIgnoreList(ctx) {
    var f = new File(ctx.logDir.fullName + '/' + HF.IGNORE_FILE);
    if (!f.exists) return;
    try {
        var lines = readTextFile(f).split(/\r\n|\r|\n/);
        for (var i = 0; i < lines.length; i++) {
            var tab = lines[i].indexOf('\t');
            if (tab > 0) ctx.ignore[lines[i].substr(0, tab)] = lines[i].substr(tab + 1);
        }
    } catch (e) { }
}

// Removes a dropped folder once it has been empty (apart from Finder/Explorer junk) for a while.
function hfRemoveIfEmpty(folder, nowT) {
    var items = folder.getFiles();
    for (var i = 0; i < items.length; i++) {
        if (items[i] instanceof Folder || !HF.JUNK.test(File.decode(items[i].name))) return;
    }
    var m = folder.modified;
    if (!m || nowT - m.getTime() < HF.EMPTY_FOLDER_SECONDS * 1000) return;
    for (var j = 0; j < items.length; j++) {
        try { items[j].remove(); } catch (e) { }
    }
    try { folder.remove(); } catch (e2) { }
}

function hfIsReserved(name) {
    var n = String(name).toLowerCase();
    return n === HF.OUTPUT.toLowerCase() || n === HF.ORIGINALS.toLowerCase() || n === HF.FAILED.toLowerCase();
}

function hfScanFolder(ctx, q, folder, rel, depth, nowT, jobs) {
    var items = folder.getFiles();
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        var name = File.decode(it.name);
        if (it instanceof Folder) {
            if (name.charAt(0) === '.') continue;
            if (depth === 0 && hfIsReserved(name)) continue;
            if (depth + 1 > HF.MAX_DEPTH) continue;
            hfScanFolder(ctx, q, it, rel ? rel + '/' + name : name, depth + 1, nowT, jobs);
            hfRemoveIfEmpty(it, nowT);
            continue;
        }
        if (name.charAt(0) === '.' || name.indexOf('~$') === 0 || HF.JUNK.test(name)) continue;
        if (depth === 0 && name === HF.SETTINGS_FILE) continue;
        if (hfIsIgnored(ctx, it)) continue;
        ctx.waiting[q.name] = (ctx.waiting[q.name] || 0) + 1;
        if (HF.ARRIVING.test(name)) continue; // browser/partial download - wait for the real file
        if (!ctx.tracker.isReady(it, nowT)) continue;
        var m = it.modified;
        jobs.push({
            file: it,
            name: name,
            rel: rel,
            queue: q,
            supported: HEC.FILE_TYPES.test(name),
            time: m ? m.getTime() : nowT
        });
    }
}

// Images dropped straight into the hot folder instead of a queue.
function hfCountStrays(root) {
    var n = 0;
    try {
        var items = root.getFiles();
        for (var i = 0; i < items.length; i++) {
            if (items[i] instanceof File && HEC.FILE_TYPES.test(File.decode(items[i].name))) n++;
        }
    } catch (e) { }
    return n;
}

function hfStopRequested(root) {
    try {
        var items = root.getFiles();
        for (var i = 0; i < items.length; i++) {
            if (items[i] instanceof File && HF.STOP_NAME.test(File.decode(items[i].name))) return true;
        }
    } catch (e) { }
    return false;
}

function hfRemoveStopFiles(root) {
    try {
        var items = root.getFiles();
        for (var i = 0; i < items.length; i++) {
            if (items[i] instanceof File && HF.STOP_NAME.test(File.decode(items[i].name))) items[i].remove();
        }
    } catch (e) { }
}

// ---------------------------------------------------------------------------
// Log, status file and status window
// ---------------------------------------------------------------------------
function hfLog(ctx, status, queueName, src, out, secs, note) {
    try {
        ensureFolder(ctx.logDir);
        var f = new File(ctx.logDir.fullName + '/HotFolder_' + hfDayKey() + '.txt');
        var isNew = !f.exists;
        f.encoding = 'UTF-8';
        if (f.open('a')) {
            if (isNew) f.writeln(['time', 'status', 'queue', 'file', 'output', 'seconds', 'notes'].join('\t'));
            f.writeln([clockTime(), status, queueName || '', src || '', out || '',
                (secs === undefined || secs === '') ? '' : secs + 's', note || ''].join('\t'));
            f.close();
        }
    } catch (e) { /* logging must never stop the watcher */ }
}

function hfJobPath(job) { return job.queue.name + '/' + (job.rel ? job.rel + '/' : '') + job.name; }

function hfQueueLines(ctx) {
    var lines = [];
    for (var i = 0; i < ctx.queues.length; i++) {
        var q = ctx.queues[i];
        lines.push(q.name + ': ' + (ctx.waiting[q.name] || 0) + ' waiting' +
            (q.error ? '   (problem: ' + q.error + ' - files will wait)' : ''));
    }
    if (!lines.length) lines.push('No queues found.');
    if (ctx.strays) {
        lines.push(ctx.strays + ' image(s) dropped straight into the hot folder are ignored - move them into a queue folder.');
    }
    return lines;
}

function hfTodayLine(ctx) {
    return 'Today: ' + ctx.doneToday + ' done, ' + ctx.failedToday + ' failed, ' + ctx.gensToday +
        (ctx.opts.dailyLimit > 0 ? ' of ' + ctx.opts.dailyLimit : '') + ' generations used';
}

function hfWriteStatus(ctx) {
    var lines = [
        HF.NAME + ' ' + HF.VERSION + ' - Photoshop ' + app.version,
        'State: ' + ctx.state,
        'Last check: ' + hfDateTime(ctx.lastCheck),
        'Running since: ' + hfDateTime(ctx.startedAt),
        hfTodayLine(ctx),
        '',
        'Queues:'
    ];
    var q = hfQueueLines(ctx);
    for (var i = 0; i < q.length; i++) lines.push('  ' + q[i]);
    lines.push('');
    lines.push('To stop the watcher, put a file named STOP in this folder.');
    try { writeTextFile(new File(ctx.root.fullName + '/' + HF.STATUS_FILE), lines.join('\n') + '\n'); } catch (e) { }
}

function HFWindow(root) {
    this.win = null;
    this.stop = false;
    this.why = '';
    try {
        var self = this;
        var w = new Window('palette', HF.NAME);
        w.orientation = 'column';
        w.alignChildren = ['fill', 'top'];
        w.margins = 14;
        w.spacing = 6;
        w.add('statictext', undefined, 'Watching: ' + root.fsName);
        this.stateText = w.add('statictext', undefined, 'Starting...');
        this.stateText.preferredSize.width = 520;
        this.todayText = w.add('statictext', undefined, '');
        this.todayText.preferredSize.width = 520;
        this.queueText = w.add('statictext', undefined, '', { multiline: true });
        this.queueText.preferredSize = [520, 90];
        w.add('statictext', undefined, 'Stop: click the button, close this window, put a file named STOP in the hot folder, or press Esc.');
        var b = w.add('button', undefined, 'Stop watching');
        b.onClick = function () {
            self.stop = true;
            self.why = self.why || 'Stop button';
            b.text = 'Stopping...';
            b.enabled = false;
        };
        w.onClose = function () { // closing the window stops watching too
            self.stop = true;
            self.why = self.why || 'status window closed';
        };
        w.show();
        this.win = w;
    } catch (e) { this.win = null; }
}
HFWindow.prototype.show = function (ctx) {
    if (!this.win) return;
    try {
        this.stateText.text = ctx.state;
        this.todayText.text = hfTodayLine(ctx);
        this.queueText.text = hfQueueLines(ctx).join('\n');
        this.win.update();
    } catch (e) { }
};
HFWindow.prototype.pump = function () {
    if (!this.win) return;
    try { this.win.update(); } catch (e) { }
};
HFWindow.prototype.close = function () {
    if (this.win) { try { this.win.close(); } catch (e) { } }
};

// ---------------------------------------------------------------------------
// Processing one file
// ---------------------------------------------------------------------------
function hfMarkerFile(ctx) { return new File(ctx.logDir.fullName + '/' + HF.MARKER_FILE); }

function hfMarkStart(ctx, job) {
    var count = (ctx.crashCounts[hfKey(job.file)] || 0) + 1;
    try {
        writeTextFile(hfMarkerFile(ctx), [job.file.fsName, count, job.queue.name, job.rel].join('\n') + '\n');
    } catch (e) { }
}

function hfMarkEnd(ctx) {
    try {
        var m = hfMarkerFile(ctx);
        if (m.exists) m.remove();
    } catch (e) { }
}

// Moves the original into Originals/Failed (mirroring dropped folders). Returns a
// warning, or '' if all went well.
function hfMoveOriginal(ctx, job, bucket, noteText) {
    var dest = new Folder(job.queue.folder.fullName + '/' + bucket + (job.rel ? '/' + job.rel : ''));
    try {
        var r = hfMoveFile(job.file, dest, job.name);
        if (noteText) {
            try { writeTextFile(new File(r.target.fullName + '.txt'), noteText); } catch (eNote) { }
        }
        if (!r.removed) {
            hfIgnore(ctx, job.file);
            return 'copied to ' + bucket + ' but the original could not be deleted from the drop folder - ' +
                'delete it by hand (it will not be processed again)';
        }
        return '';
    } catch (e) {
        hfIgnore(ctx, job.file);
        return 'could not move the original to ' + bucket + ' (' + errMsg(e) + ') - it will not be processed again';
    }
}

function hfFailJob(ctx, job, reason, secs) {
    var note = [
        HF.NAME + ' could not process this file.',
        'File: ' + job.name,
        'Queue: ' + job.queue.name,
        'When: ' + hfDateTime(nowMs()),
        'Reason: ' + reason,
        '',
        'To try again, move the file back into the queue folder.'
    ].join('\n') + '\n';
    var warn = hfMoveOriginal(ctx, job, HF.FAILED, note);
    delete ctx.tries[hfKey(job.file)];
    delete ctx.retryAt[hfKey(job.file)];
    ctx.failedToday++;
    hfLog(ctx, 'FAILED', job.queue.name, hfJobPath(job), '', secs, reason + (warn ? '; ' + warn : ''));
}

function hfPause(ctx, reason) {
    ctx.pausedUntil = nowMs() + ctx.opts.pauseMinutes * 60000;
    ctx.pauseReason = reason;
    ctx.genFailsInARow = 0;
    hfLog(ctx, 'PAUSED', '', '', '', '', 'until ' + hfTime(ctx.pausedUntil) + ': ' + reason);
}

function hfProcessFile(ctx, job, S) {
    var fmtInfo = resolveFormat(S, getExt(job.name));
    var outRoot = S.destFolder ? new Folder(S.destFolder) : new Folder(job.queue.folder.fullName + '/' + HF.OUTPUT);
    var outDir = job.rel ? new Folder(outRoot.fullName + '/' + job.rel) : outRoot;
    ensureFolder(outDir);
    var outFile = hfUniqueFile(outDir, stripExt(job.name) + S.suffix, '.' + fmtInfo.fmt);
    var notes = [];
    if (fmtInfo.note) notes.push(fmtInfo.note);
    var doc = null;
    hfMarkStart(ctx, job);
    try {
        doc = app.open(job.file);
        processDocument(doc, S, notes);
        saveOutput(doc, outFile, fmtInfo.fmt, S);
    } finally {
        if (doc) {
            try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose) { }
        }
        hfMarkEnd(ctx);
        try { app.purge(PurgeTarget.ALLCACHES); } catch (ePurge) { }
        try { $.gc(); } catch (eGc) { }
    }
    return { outFile: outFile, notes: notes };
}

function hfRunJob(ctx, job, S, expands) {
    ctx.state = 'Processing ' + hfJobPath(job);
    hfWriteStatus(ctx);
    ctx.win.show(ctx);
    var t0 = nowMs();
    var err = null;
    var res = null;
    try {
        res = hfProcessFile(ctx, job, S);
    } catch (e) {
        if (isUserCancel(e)) throw e; // Esc: stop watching, the file just waits
        err = e;
    }
    ctx.waiting[job.queue.name] = Math.max(0, (ctx.waiting[job.queue.name] || 1) - 1);
    var k = hfKey(job.file);
    if (!err) {
        ctx.doneToday++;
        if (expands) ctx.gensToday++;
        ctx.genFailsInARow = 0;
        delete ctx.tries[k];
        delete ctx.retryAt[k];
        delete ctx.crashCounts[k];
        var warn = hfMoveOriginal(ctx, job, HF.ORIGINALS, '');
        hfLog(ctx, 'OK', job.queue.name, hfJobPath(job), File.decode(res.outFile.name), secondsSince(t0),
            res.notes.concat(warn ? [warn] : []).join('; '));
        return;
    }
    var msg = errMsg(err);
    if (!err.hecGenerative) {
        hfFailJob(ctx, job, msg, secondsSince(t0));
        return;
    }
    // Firefly problem: often not the file's fault, so keep it and try again later.
    ctx.genFailsInARow++;
    if (err.hecCredits) {
        hfLog(ctx, 'RETRY', job.queue.name, hfJobPath(job), '', secondsSince(t0), msg);
        hfPause(ctx, 'generative credits look exhausted');
        return;
    }
    ctx.tries[k] = (ctx.tries[k] || 0) + 1;
    if (ctx.tries[k] >= HF.MAX_TRIES) {
        hfFailJob(ctx, job, msg + ' (tried ' + HF.MAX_TRIES + ' times)', secondsSince(t0));
    } else {
        // Wait before trying this file again: pauseMinutes, then twice that, so an outage
        // has to last a good while before anything lands in Failed.
        var waitMin = ctx.opts.pauseMinutes * Math.pow(2, ctx.tries[k] - 1);
        ctx.retryAt[k] = nowMs() + waitMin * 60000;
        hfLog(ctx, 'RETRY', job.queue.name, hfJobPath(job), '', secondsSince(t0),
            msg + ' (try ' + ctx.tries[k] + ' of ' + HF.MAX_TRIES + ', next try in ' + waitMin + ' min)');
    }
    if (ctx.genFailsInARow >= HF.FAILS_BEFORE_PAUSE) {
        hfPause(ctx, HF.FAILS_BEFORE_PAUSE + ' Generative Fill failures in a row (last: ' + msg + ')');
    }
}

// If Photoshop died mid-file last time, count it against that file.
function hfRecoverFromCrash(ctx) {
    var m = hfMarkerFile(ctx);
    if (!m.exists) return;
    var lines;
    try {
        lines = readTextFile(m).split(/\r\n|\r|\n/);
        m.remove();
    } catch (e) { return; }
    var f = new File(lines[0]);
    if (!lines[0] || !f.exists) return;
    var count = parseInt(lines[1], 10) || 1;
    var queue = { name: lines[2] || '', folder: new Folder(ctx.root.fullName + '/' + (lines[2] || '')) };
    var job = { file: f, name: File.decode(f.name), rel: lines[3] || '', queue: queue };
    if (count >= HF.CRASH_LIMIT) {
        hfFailJob(ctx, job, 'Photoshop stopped while processing this file ' + count +
            ' times, so it was not tried again', '');
    } else {
        ctx.crashCounts[hfKey(f)] = count;
        hfLog(ctx, 'NOTE', queue.name, hfJobPath(job), '', '',
            'Photoshop stopped while processing this file last time - trying once more');
    }
}

// ---------------------------------------------------------------------------
// The watch loop
// ---------------------------------------------------------------------------
function hfCheckStop(ctx) {
    ctx.win.pump();
    if (ctx.win.stop) ctx.stopReason = ctx.win.why || 'Stop button';
    else if (hfStopRequested(ctx.root)) ctx.stopReason = 'STOP file';
    return !!ctx.stopReason;
}

// One pass: find ready files in every queue and process them, oldest first.
function hfCycle(ctx) {
    var t = nowMs();
    ctx.lastCheck = t;
    var day = hfDayKey();
    if (day !== ctx.day) {
        ctx.day = day;
        ctx.doneToday = 0;
        ctx.failedToday = 0;
        ctx.gensToday = 0;
    }
    if (!ctx.root.exists) {
        ctx.state = 'Hot folder not reachable - will keep trying';
        return;
    }
    if (hfCheckStop(ctx)) return;

    ctx.queues = hfFindQueues(ctx.root);
    ctx.waiting = {};
    var jobs = [];
    ctx.tracker.begin();
    for (var i = 0; i < ctx.queues.length; i++) {
        var q = ctx.queues[i];
        try {
            q.settings = hfReadQueueSettings(q);
            if (q.settings.destFolder) {
                // A custom output folder that can't be reached pauses this queue, not fail its files.
                try {
                    ensureFolder(new Folder(q.settings.destFolder));
                } catch (eDest) {
                    q.error = 'output folder not reachable: ' + q.settings.destFolder;
                }
            }
        } catch (eSet) {
            q.error = errMsg(eSet);
        }
        try {
            hfScanFolder(ctx, q, q.folder, '', 0, t, jobs);
        } catch (eScan) {
            q.error = 'could not read the folder: ' + errMsg(eScan);
        }
    }
    ctx.tracker.end();
    ctx.strays = hfCountStrays(ctx.root);

    if (ctx.pausedUntil) {
        if (nowMs() < ctx.pausedUntil) {
            ctx.state = 'Paused until ' + hfTime(ctx.pausedUntil) + ' - ' + ctx.pauseReason;
            return;
        }
        ctx.pausedUntil = 0;
        ctx.pauseReason = '';
        hfLog(ctx, 'RESUMED', '', '', '', '', '');
    }

    jobs.sort(function (a, b) { return a.time - b.time; });
    var limitHit = false;
    for (var j = 0; j < jobs.length; j++) {
        var job = jobs[j];
        if (!job.file.exists) continue; // taken back out
        if (job.queue.error) continue;   // settings unreadable - shown in the status
        if (!job.supported) {
            ctx.waiting[job.queue.name] = Math.max(0, (ctx.waiting[job.queue.name] || 1) - 1);
            hfFailJob(ctx, job, 'not a supported image type (' + (hfExtOf(job.name) || 'no extension') + ')', '');
            continue;
        }
        var jk = hfKey(job.file);
        if (ctx.retryAt[jk] && nowMs() < ctx.retryAt[jk]) continue; // backing off after a Firefly failure
        var S = job.queue.settings;
        var expands = S.top + S.bottom + S.left + S.right > 0;
        if (expands && ctx.opts.dailyLimit > 0 && ctx.gensToday >= ctx.opts.dailyLimit) {
            limitHit = true;
            continue;
        }
        hfRunJob(ctx, job, S, expands);
        if (ctx.pausedUntil) {
            ctx.state = 'Paused until ' + hfTime(ctx.pausedUntil) + ' - ' + ctx.pauseReason;
            return;
        }
        if (hfCheckStop(ctx)) return;
    }
    ctx.state = limitHit ?
        'Daily limit of ' + ctx.opts.dailyLimit + ' generations reached - expanding resumes tomorrow' :
        'Watching (checks every ' + ctx.opts.interval + ' s)';
}

function hfIdle(ctx, seconds) {
    var end = nowMs() + seconds * 1000;
    var nextCheck = nowMs() + 2000;
    while (nowMs() < end && !ctx.stopReason) {
        $.sleep(250);
        ctx.win.pump();
        if (ctx.win.stop) {
            ctx.stopReason = ctx.win.why || 'Stop button';
        } else if (nowMs() >= nextCheck) {
            nextCheck = nowMs() + 2000;
            if (hfStopRequested(ctx.root)) ctx.stopReason = 'STOP file';
            try { app.refresh(); } catch (e) { if (isUserCancel(e)) ctx.stopReason = 'Esc'; }
        }
    }
}

function hfWatch(opts) {
    var root = new Folder(opts.root);
    var ctx = {
        opts: opts,
        root: root,
        logDir: new Folder(root.fullName + '/' + HF.LOG_FOLDER),
        tracker: new HFTracker(HF.STABLE_SECONDS),
        queues: [],
        waiting: {},
        tries: {},
        retryAt: {},
        ignore: {},
        crashCounts: {},
        day: hfDayKey(),
        doneToday: 0,
        failedToday: 0,
        gensToday: 0,
        genFailsInARow: 0,
        pausedUntil: 0,
        pauseReason: '',
        stopReason: '',
        state: 'Starting',
        startedAt: nowMs(),
        lastCheck: 0,
        strays: 0,
        win: null
    };
    hfRemoveStopFiles(root); // an old STOP file would end the new session at once
    ensureFolder(ctx.logDir);
    hfLoadIgnoreList(ctx);
    ctx.win = new HFWindow(root);
    hfLog(ctx, 'START', '', '', '', '', 'Photoshop ' + app.version + ', checks every ' + opts.interval +
        ' s, daily limit ' + (opts.dailyLimit || 'none'));
    try {
        try {
            hfRecoverFromCrash(ctx);
        } catch (eRec) {
            hfLog(ctx, 'ERROR', '', '', '', '', 'crash recovery: ' + errMsg(eRec));
        }
        while (!ctx.stopReason) {
            try {
                hfCycle(ctx);
            } catch (e) {
                if (isUserCancel(e)) {
                    ctx.stopReason = 'Esc';
                } else {
                    ctx.state = 'Problem: ' + errMsg(e) + ' - will keep trying';
                    hfLog(ctx, 'ERROR', '', '', '', '', errMsg(e) + (e.line ? ' (line ' + e.line + ')' : ''));
                }
            }
            if (ctx.stopReason) break;
            hfWriteStatus(ctx);
            ctx.win.show(ctx);
            hfIdle(ctx, opts.interval);
        }
    } finally {
        hfRemoveStopFiles(root);
        ctx.state = 'Stopped (' + ctx.stopReason + ') at ' + hfDateTime(nowMs());
        hfWriteStatus(ctx);
        hfLog(ctx, 'STOP', '', '', '', '', ctx.stopReason);
        ctx.win.close();
    }
    return ctx;
}

// ---------------------------------------------------------------------------
// Setup dialog
// ---------------------------------------------------------------------------
function hfShowSetup(o) {
    var result = null;
    var dlg = new Window('dialog', HF.NAME + '  v' + HF.VERSION);
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
        parent.add('statictext', undefined, label);
        var et = parent.add('edittext', undefined, String(value));
        et.characters = chars;
        if (tip) et.helpTip = tip;
        return et;
    }

    var pRoot = addPanel('Hot folder');
    var gRoot = addRow(pRoot);
    var etRoot = gRoot.add('edittext', undefined, o.root);
    etRoot.preferredSize.width = 440;
    var btRoot = gRoot.add('button', undefined, 'Browse...');
    pRoot.add('statictext', undefined, 'Queues - people drop images into these folders:');
    var lb = pRoot.add('listbox', undefined, []);
    lb.preferredSize = [540, 120];
    var gQ = addRow(pRoot);
    var btAdd = gQ.add('button', undefined, 'Add queue...');
    var btEdit = gQ.add('button', undefined, 'Edit queue...');
    var btStarter = gQ.add('button', undefined, 'Create starter queues');
    btStarter.helpTip = 'Expand - JPEG, Cutout - transparent PNG and Cutout - white JPEG, all at 30% per side.';

    var pW = addPanel('Watching');
    var gW = addRow(pW);
    var etInt = addNum(gW, 'Check every (s):', o.interval, 4);
    var etLimit = addNum(gW, '   Daily generation limit (0 = none):', o.dailyLimit, 5,
        'Stops using generative credits after this many expanded images a day. Cutout-only queues keep going.');
    var gW2 = addRow(pW);
    var etPause = addNum(gW2, 'When Firefly keeps failing or credits run out, pause for (min):', o.pauseMinutes, 4);
    var cbAuto = pW.add('checkbox', undefined, 'Next time, start watching straight away (hold Shift while it starts to see this dialog)');
    cbAuto.value = o.autoStart;
    pW.add('statictext', undefined, 'Photoshop is busy while it watches. Stop: click Stop, put a file named STOP in the hot folder, or press Esc.');

    var gB = dlg.add('group');
    gB.alignment = ['right', 'top'];
    var btClose = gB.add('button', undefined, 'Close', { name: 'cancel' });
    var btStart = gB.add('button', undefined, 'Start watching', { name: 'ok' });

    var queues = [];
    function rootFolder() {
        var p = trimStr(etRoot.text);
        return p ? new Folder(p) : null;
    }
    function refreshList() {
        lb.removeAll();
        queues = [];
        var r = rootFolder();
        if (!r) {
            lb.add('item', '(choose or type the hot folder above)');
        } else if (!r.exists) {
            lb.add('item', '(folder does not exist yet - Add queue or Create starter queues will create it)');
        } else {
            queues = hfFindQueues(r);
            if (!queues.length) lb.add('item', '(no queues yet - use Add queue or Create starter queues)');
            for (var i = 0; i < queues.length; i++) {
                var line = queues[i].name + '   -   ';
                try {
                    line += hfQueueSummary(hfReadQueueSettings(queues[i]));
                } catch (e) {
                    line += 'settings unreadable: ' + errMsg(e);
                }
                lb.add('item', line);
            }
        }
        btEdit.enabled = queues.length > 0;
    }
    function ensureRoot() {
        var r = rootFolder();
        if (!r) {
            alert('Choose the hot folder first.');
            return null;
        }
        try {
            ensureFolder(r);
        } catch (e) {
            alert('Could not create the hot folder:\n' + errMsg(e));
            return null;
        }
        return r;
    }

    btRoot.onClick = function () {
        var start = rootFolder();
        var f = (start && start.exists) ? start.selectDlg('Choose the hot folder') : Folder.selectDialog('Choose the hot folder');
        if (f) etRoot.text = f.fsName;
        refreshList();
    };
    etRoot.onChange = refreshList;

    btAdd.onClick = function () {
        var r = ensureRoot();
        if (!r) return;
        var R = showDialog(hfDefaultQueueSettings(), false, {
            mode: 'queue',
            isNew: true,
            queueName: '',
            validateName: function (n) {
                return new Folder(r.fullName + '/' + n).exists ? 'There is already a folder called "' + n + '" in the hot folder.' : '';
            }
        });
        if (!R) return;
        try {
            hfCreateQueue(r, R.queueName, R);
        } catch (e) {
            alert('Could not create the queue:\n' + errMsg(e));
        }
        refreshList();
    };

    btEdit.onClick = function () {
        var idx = lb.selection ? lb.selection.index : -1;
        if (idx < 0 || idx >= queues.length) {
            alert('Select a queue in the list first.');
            return;
        }
        var q = queues[idx];
        var S;
        try {
            S = hfReadQueueSettings(q);
        } catch (e) {
            S = hfDefaultQueueSettings();
        }
        var R = showDialog(S, false, { mode: 'queue', isNew: false, queueName: q.name });
        if (!R) return;
        try {
            hfWriteQueueSettings(q.folder, R);
        } catch (e2) {
            alert('Could not save the settings:\n' + errMsg(e2));
        }
        refreshList();
    };
    lb.onDoubleClick = btEdit.onClick;

    btStarter.onClick = function () {
        var r = ensureRoot();
        if (!r) return;
        var starters = hfStarterQueues();
        var made = [];
        for (var i = 0; i < starters.length; i++) {
            if (new Folder(r.fullName + '/' + starters[i].name).exists) continue;
            try {
                hfCreateQueue(r, starters[i].name, starters[i].S);
                made.push(starters[i].name);
            } catch (e) {
                alert('Could not create "' + starters[i].name + '":\n' + errMsg(e));
            }
        }
        alert(made.length ? 'Created:\n' + made.join('\n') + '\n\nEdit them to change sizes, formats or backgrounds.' :
            'The starter queues are already there.');
        refreshList();
    };

    function readOptions(errors) {
        function num(et, label, min, max) {
            var v = parseFloat(trimStr(et.text));
            if (isNaN(v) || v < min || v > max) {
                errors.push(label + ' must be a number from ' + min + ' to ' + max + '.');
                return min;
            }
            return Math.round(v);
        }
        var R = cloneSettings(o);
        R.root = trimStr(etRoot.text);
        R.interval = num(etInt, 'Check every', 2, 3600);
        R.dailyLimit = num(etLimit, 'Daily generation limit', 0, 100000);
        R.pauseMinutes = num(etPause, 'Pause', 1, 1440);
        R.autoStart = cbAuto.value;
        return R;
    }

    btStart.onClick = function () {
        var errors = [];
        var R = readOptions(errors);
        if (!R.root || !new Folder(R.root).exists) {
            errors.push('Choose a hot folder that exists.');
        } else if (!hfFindQueues(new Folder(R.root)).length) {
            errors.push('The hot folder has no queues yet - add one or create the starter queues.');
        }
        if (errors.length) {
            alert(errors.join('\n'));
            return;
        }
        result = R;
        dlg.close(1);
    };
    btClose.onClick = function () { dlg.close(2); };
    // However the dialog is closed, remember the folder and options for next time.
    dlg.onClose = function () {
        if (!result) {
            try { hfSaveOptions(readOptions([])); } catch (e) { }
        }
    };

    refreshList();
    dlg.center();
    dlg.show();
    return result;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
function hfMain() {
    if (typeof app === 'undefined' || !/photoshop/i.test(app.name)) {
        alert('Run this script from Adobe Photoshop.');
        return;
    }
    if (HF_LIB_ERROR) {
        alert(HF.NAME + '\n\n' + HF_LIB_ERROR);
        return;
    }
    if (typeof settingsFromText !== 'function' || typeof processDocument !== 'function') {
        alert(HF.NAME + '\n\nPlease update "Headshot Expand and Cutout.jsx" to v1.1 or later.');
        return;
    }
    var o = hfLoadOptions();
    var shift = false;
    try { shift = ScriptUI.environment.keyboardState.shiftKey; } catch (e) { }
    var auto = o.autoStart && !shift && o.root && new Folder(o.root).exists &&
        hfFindQueues(new Folder(o.root)).length > 0;
    if (!auto) {
        var R = hfShowSetup(o);
        if (!R) return;
        o = R;
        hfSaveOptions(o);
    }

    if (parseInt(app.version, 10) < 25) {
        alert(HF.NAME + '\n\nGenerative Fill needs Photoshop 2024 (v25) or later. This is v' + app.version +
            ' - queues that expand will fail. Use cutout-only queues or update Photoshop.');
    }

    var saved = {
        rulerUnits: app.preferences.rulerUnits,
        typeUnits: app.preferences.typeUnits,
        dialogs: app.displayDialogs
    };
    app.preferences.rulerUnits = Units.PIXELS;
    app.preferences.typeUnits = TypeUnits.PIXELS;
    app.displayDialogs = DialogModes.NO;
    var ctx = null;
    try {
        ctx = hfWatch(o);
    } catch (e) {
        if (!isUserCancel(e)) alert(HF.NAME + ' stopped:\n' + errMsg(e) + (e.line ? '\n(line ' + e.line + ')' : ''));
    } finally {
        app.preferences.rulerUnits = saved.rulerUnits;
        app.preferences.typeUnits = saved.typeUnits;
        app.displayDialogs = saved.dialogs;
    }
    if (ctx) {
        alert(HF.NAME + ' stopped (' + ctx.stopReason + ').\n\n' + hfTodayLine(ctx) +
            '\nLogs: ' + ctx.logDir.fsName);
    }
}

if (typeof HF_NO_AUTORUN === 'undefined') {
    hfMain();
}
