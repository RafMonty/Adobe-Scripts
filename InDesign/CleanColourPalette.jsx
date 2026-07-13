/**
 * CleanColourPalette.jsx — v1.1
 * Consolidate a strictly-named swatch palette to ONE variant per brand colour.
 * InDesign ExtendScript (2023+ Mac/Win)
 *
 * v1.1 CHANGES — CRITICAL FIX
 * - v1.0 held index-based colour references (doc.colors[i]); every removal
 *   shifted the indices, so later operations resolved onto the WRONG swatches
 *   (cross-group remaps, remaps to Paper, self-replacement errors, renames
 *   landing on InDesign's locked internal ink colours). v1.1 resolves every
 *   colour BY ID (stable across removals) at the moment of each operation,
 *   verifies survivor != victim, and logs live names.
 * - Rename collisions are now MERGED, not skipped: if the final name already
 *   exists (e.g. a bare 'PANTONE 9224 C' imported by placed art, or debris
 *   from an earlier run), that stray is removed with its usages remapped to
 *   the survivor, then the survivor takes the name. Convention: same name =
 *   same colour. If the stray can't be merged (mixed ink), the suffixed name
 *   is kept and reported.
 *
 * NAMING CONVENTION EXPECTED
 *   <Friendly>_CMYK          e.g.  Linen_CMYK          (process CMYK build)
 *   <Pantone base> SPOT      e.g.  Pantone 9224 C SPOT (spot)
 *   <Pantone base> LAB       e.g.  Pantone 9224 C LAB  (process Lab)
 *
 * The Friendly <-> Pantone linkage cannot be derived from the names, so it
 * lives in the ALIAS table below. EDIT IT to match your brand palette.
 *
 * WHAT IT DOES
 *   1. Lists every named custom colour with its model/space and the PLANNED
 *      action for the chosen mode (Make LAB / Make SPOT / Make CMYK) —
 *      nothing happens until you hit Run.
 *   2. In the chosen mode, every other variant in a group is removed with
 *      InDesign's native replace: colour.remove(survivor). That remaps ALL
 *      usages — applied fills/strokes, paragraph/character/object/table
 *      styles, gradient stops — and rebases tints. This is the same
 *      "Delete Swatch and replace with" you'd do by hand, so styles are
 *      redefined correctly.
 *   3. Survivors are renamed with the suffix stripped:
 *         LAB / SPOT modes ->  Pantone 9224 C
 *         CMYK mode        ->  Linen
 *      Name collisions are reported and skipped, never forced.
 *   4. Optionally derives a MISSING target from the SPOT definition
 *      (e.g. no "Pantone 9224 C LAB" but the spot is Lab-defined: duplicate
 *      spot -> process Lab). Derivations are gated on the colour space
 *      matching, and every one is reported.
 *   5. Swatches it does not understand (no suffix, or _CMYK with no ALIAS
 *      entry) are LEFT ALONE and reported. Built-ins ([Black] etc.),
 *      unnamed local colours, tints and gradients are never touched
 *      directly.
 *
 * LIMITS
 *   - Colours inside PLACED artwork (EPS/PDF/AI links) cannot be changed by
 *     any swatch operation. If a placed logo carries the spot, the output
 *     still carries the spot.
 *   - Spots participating in Mixed Ink groups cannot be removed; these are
 *     caught and reported.
 *
 * Entire run = one undo step. Install in the Scripts Panel folder.
 */

#target "InDesign"

(function () {
    if (!app.documents.length) { alert("Open a document first."); return; }

    // =====================================================================
    // EDIT ME — Friendly (used by *_CMYK swatches) -> Pantone base name.
    // Base name = the swatch name WITHOUT the " SPOT" / " LAB" suffix.
    // =====================================================================
    var ALIAS = {
        "Linen":  "Pantone 9224 C",
        "Forest": "Pantone 3435 C",
        "Moss":   "Pantone 5777 C",
        "Sky":    "Pantone 2169 C"
    };

    var doc = app.activeDocument;

    // Reverse lookup: Pantone base -> Friendly
    var REV = {};
    for (var aKey in ALIAS) { if (ALIAS.hasOwnProperty(aKey)) REV[ALIAS[aKey]] = aKey; }

    var inv = inventory(doc);
    if (!inv.length) { alert("No named custom colours found in this document."); return; }

    var cfg = showDialog(doc, inv);
    if (!cfg) return;

    var reportText = "";
    app.doScript(function () {
        reportText = execute(doc, cfg);
    }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Clean Colour Palette");

    showReport("Clean Colour Palette \u2014 Report", reportText);
    return;

    // ---------------- Inventory & classification ----------------

    function inventory(docRef) {
        var out = [], i;
        for (i = 0; i < docRef.colors.length; i++) {
            var c = docRef.colors[i];
            var nm = "";
            try { nm = c.name; } catch (_) { continue; }
            if (nm === "" || nm.charAt(0) === "[") continue; // unnamed locals & built-ins
            var cls = classify(nm);
            var cid = null;
            try { cid = c.id; } catch (_) { continue; }
            out.push({
                color: c,
                id: cid,
                name: nm,
                model: modelName(c),
                space: spaceName(c),
                variant: cls.variant,   // "CMYK" | "SPOT" | "LAB" | "OTHER"
                base: cls.base          // name minus suffix (friendly for _CMYK)
            });
        }
        return out;
    }

    function classify(name) {
        var m;
        if ((m = name.match(/^(.*?)[\s_]+CMYK\s*$/i))) return { variant: "CMYK", base: trimStr(m[1]) };
        if ((m = name.match(/^(.*?)[\s_]+SPOT\s*$/i))) return { variant: "SPOT", base: trimStr(m[1]) };
        if ((m = name.match(/^(.*?)[\s_]+LAB\s*$/i)))  return { variant: "LAB",  base: trimStr(m[1]) };
        return { variant: "OTHER", base: name };
    }

    function modelName(c) {
        try {
            var m = c.model;
            if (m === ColorModel.PROCESS) return "Process";
            if (m === ColorModel.SPOT) return "Spot";
            if (m === ColorModel.REGISTRATION) return "Registration";
            if (m === ColorModel.MIXEDINKMODEL) return "Mixed Ink";
        } catch (_) {}
        return "?";
    }

    function spaceName(c) {
        try {
            var s = c.space;
            if (s === ColorSpace.CMYK) return "CMYK";
            if (s === ColorSpace.RGB) return "RGB";
            if (s === ColorSpace.LAB) return "Lab";
            if (s === ColorSpace.MIXEDINK) return "Mixed Ink";
        } catch (_) {}
        return "?";
    }

    function trimStr(s) { return String(s === null || s === undefined ? "" : s).replace(/^\s+|\s+$/g, ""); }

    // ---------------- Plan builder (shared by preview & execute) ----------------
    //
    // Returns:
    //   steps:   ordered operations {op:"create"|"remove"|"rename", ...}
    //   actions: name -> human-readable planned action (for the list column)
    //   notes:   general remarks for the report
    //
    function makePlan(invArr, mode, opts) {
        var groups = {}, actions = {}, steps = [], notes = [], i, k;

        for (i = 0; i < invArr.length; i++) {
            var it = invArr[i];
            if (it.variant === "OTHER") {
                actions[it.name] = "leave \u2014 no recognised suffix";
                continue;
            }
            var key = null, friendly = null;
            if (it.variant === "CMYK") {
                friendly = it.base;
                key = ALIAS[it.base] || null;
                if (!key) {
                    actions[it.name] = "leave \u2014 '" + it.base + "' not in ALIAS table";
                    continue;
                }
            } else {
                key = it.base;
                friendly = REV[key] || null;
            }
            if (!groups[key]) groups[key] = { friendly: friendly, CMYK: null, SPOT: null, LAB: null };
            if (friendly && !groups[key].friendly) groups[key].friendly = friendly;
            if (groups[key][it.variant]) {
                actions[it.name] = "leave \u2014 duplicate " + it.variant + " variant for '" + key + "'";
                notes.push("Duplicate " + it.variant + " swatch for group '" + key + "': '" + it.name + "' ignored.");
                continue;
            }
            groups[key][it.variant] = it;
        }

        for (k in groups) {
            if (!groups.hasOwnProperty(k)) continue;
            var g = groups[k];
            var target = g[mode];
            var targetLabel = null;

            if (!target && opts.createMissing) {
                var derived = deriveSource(g, mode);
                if (derived) {
                    targetLabel = "(new) " + k + " " + mode;
                    steps.push({ op: "create", from: derived.item, mode: mode, tempName: k + " " + mode, groupKey: k });
                    actions[derived.item.name] = (actions[derived.item.name] ? actions[derived.item.name] + "; " : "") +
                        "source for derived " + mode + " variant";
                    notes.push("Group '" + k + "': no " + mode + " variant \u2014 will derive from '" + derived.item.name + "' (" + derived.why + ").");
                } 
            }
            if (!target && !targetLabel) {
                var why = opts.createMissing ? "no " + mode + " variant and none derivable" : "no " + mode + " variant";
                markGroupSkipped(g, actions, why);
                notes.push("Group '" + k + "' skipped: " + why + ".");
                continue;
            }

            var finalName = (mode === "CMYK") ? (g.friendly || k) : k;
            var survivorName = target ? target.name : targetLabel;

            // removals
            var variants = ["CMYK", "SPOT", "LAB"];
            for (i = 0; i < variants.length; i++) {
                var v = variants[i];
                if (v === mode) continue;
                if (g[v]) {
                    steps.push({ op: "remove", item: g[v], groupKey: k });
                    actions[g[v].name] = "REMOVE \u2192 replace usages with " + survivorName;
                }
            }
            // rename survivor
            if (target) {
                if (opts.strip && target.name !== finalName) {
                    steps.push({ op: "rename", item: target, to: finalName, groupKey: k });
                    actions[target.name] = "KEEP \u2192 rename to '" + finalName + "'";
                } else {
                    actions[target.name] = "KEEP (survivor)";
                }
            } else if (opts.strip) {
                steps.push({ op: "renameNew", groupKey: k, to: finalName });
            }
        }

        return { steps: steps, actions: actions, notes: notes };
    }

    function markGroupSkipped(g, actions, why) {
        var vs = ["CMYK", "SPOT", "LAB"], j;
        for (j = 0; j < vs.length; j++) {
            if (g[vs[j]]) actions[g[vs[j]].name] = "leave \u2014 group skipped (" + why + ")";
        }
    }

    // Which existing swatch can a missing target be derived from?
    //   LAB target : SPOT whose space is Lab   -> duplicate, model=Process
    //   CMYK target: SPOT whose space is CMYK  -> duplicate, model=Process
    //   SPOT target: LAB, else CMYK variant    -> duplicate, model=Spot
    function deriveSource(g, mode) {
        try {
            if (mode === "LAB") {
                if (g.SPOT && g.SPOT.color.space === ColorSpace.LAB) return { item: g.SPOT, why: "spot is Lab-defined" };
                return null;
            }
            if (mode === "CMYK") {
                if (g.SPOT && g.SPOT.color.space === ColorSpace.CMYK) return { item: g.SPOT, why: "spot is CMYK-defined" };
                return null;
            }
            if (mode === "SPOT") {
                if (g.LAB) return { item: g.LAB, why: "from Lab variant" };
                if (g.CMYK) return { item: g.CMYK, why: "from CMYK variant" };
                return null;
            }
        } catch (_) {}
        return null;
    }

    // ---------------- Dialog ----------------

    function showDialog(docRef, invArr) {
        var w = new Window("dialog", "Clean Colour Palette \u2014 v1.0");
        w.orientation = "column"; w.alignChildren = "fill";

        var gMode = w.add("panel", undefined, "Consolidate every brand colour to ONE variant");
        gMode.orientation = "row"; gMode.margins = 15;
        var rbLab  = gMode.add("radiobutton", undefined, "Make LAB");
        var rbSpot = gMode.add("radiobutton", undefined, "Make SPOT");
        var rbCmyk = gMode.add("radiobutton", undefined, "Make CMYK");
        rbLab.value = true;

        var list = w.add("listbox", [0, 0, 860, 320], [], { numberOfColumns: 4, showHeaders: true,
            columnTitles: ["Swatch", "Type", "Group", "Planned action"],
            columnWidths: [220, 110, 150, 360] });

        var gOpt = w.add("group");
        var chkStrip = gOpt.add("checkbox", undefined, "Strip suffix on survivors (LAB/SPOT \u2192 Pantone base, CMYK \u2192 friendly name)");
        chkStrip.value = true;
        var chkCreate = gOpt.add("checkbox", undefined, "Derive missing targets from existing variants");
        chkCreate.value = true;

        var hint = w.add("statictext", undefined,
            "Nothing changes until Run. Removals remap styles, fills/strokes, gradient stops and tints to the survivor.");

        var gBtn = w.add("group"); gBtn.alignment = "right";
        gBtn.add("button", undefined, "Cancel", { name: "cancel" });
        var btnRun = gBtn.add("button", undefined, "Run", { name: "ok" });

        function currentMode() { return rbLab.value ? "LAB" : (rbSpot.value ? "SPOT" : "CMYK"); }

        function refreshList() {
            var plan = makePlan(invArr, currentMode(), { strip: chkStrip.value, createMissing: chkCreate.value });
            list.removeAll();
            for (var i = 0; i < invArr.length; i++) {
                var it = invArr[i];
                var li = list.add("item", it.name);
                li.subItems[0].text = it.model + " " + it.space;
                li.subItems[1].text = (it.variant === "OTHER") ? "\u2014" :
                    (it.variant === "CMYK" ? (ALIAS[it.base] || "\u2014") : it.base);
                li.subItems[2].text = plan.actions[it.name] || "leave";
            }
        }
        rbLab.onClick = refreshList; rbSpot.onClick = refreshList; rbCmyk.onClick = refreshList;
        chkStrip.onClick = refreshList; chkCreate.onClick = refreshList;
        refreshList();

        if (w.show() !== 1) return null;
        return { mode: currentMode(), strip: chkStrip.value, createMissing: chkCreate.value };
    }

    // ---------------- Execute ----------------

    function execute(docRef, cfgRef) {
        var plan = makePlan(inv, cfgRef.mode, { strip: cfgRef.strip, createMissing: cfgRef.createMissing });
        var lines = [], i;
        var nRemoved = 0, nCreated = 0, nRenamed = 0, nMerged = 0, fails = [];

        lines.push("Mode: Make " + cfgRef.mode +
            (cfgRef.strip ? "   |   strip suffixes" : "") +
            (cfgRef.createMissing ? "   |   derive missing" : ""));
        lines.push("");

        // Pass 1: create derived targets (store IDs, never references)
        var createdIds = {};
        for (i = 0; i < plan.steps.length; i++) {
            var s = plan.steps[i];
            if (s.op !== "create") continue;
            try {
                var src = resolveColorByID(docRef, s.from.id);
                if (!src) { fails.push("Derivation source '" + s.from.name + "' not found \u2014 group '" + s.groupKey + "' skipped."); continue; }
                var dup = src.duplicate();
                dup.model = (s.mode === "SPOT") ? ColorModel.SPOT : ColorModel.PROCESS;
                dup.name = uniqueName(docRef, s.tempName);
                createdIds[s.groupKey] = dup.id;
                nCreated++;
                lines.push("CREATED  '" + dup.name + "'  (from '" + src.name + "')");
            } catch (eC) {
                fails.push("Could not derive " + s.mode + " for group '" + s.groupKey + "': " + eC);
            }
        }

        // Survivor per group, as an ID (freshly created, or the existing target)
        function survivorIdFor(groupKey, modeV) {
            if (createdIds[groupKey] !== undefined) return createdIds[groupKey];
            for (var j = 0; j < inv.length; j++) {
                var it = inv[j];
                if (it.variant !== modeV) continue;
                var key = (modeV === "CMYK") ? (ALIAS[it.base] || null) : it.base;
                if (key === groupKey) return it.id;
            }
            return null;
        }

        // Pass 2: removals with replacement — victim AND survivor resolved by
        // ID at the moment of the operation; identity guarded; live names logged.
        for (i = 0; i < plan.steps.length; i++) {
            var s2 = plan.steps[i];
            if (s2.op !== "remove") continue;
            var victim = resolveColorByID(docRef, s2.item.id);
            if (!victim) { fails.push("'" + s2.item.name + "' not found at removal time \u2014 skipped."); continue; }
            var surv = resolveColorByID(docRef, survivorIdFor(s2.groupKey, cfgRef.mode));
            if (!surv) { fails.push("No survivor for group '" + s2.groupKey + "' \u2014 '" + victim.name + "' left in place."); continue; }
            if (surv.id === victim.id) { fails.push("Survivor and victim identical for '" + victim.name + "' \u2014 skipped (plan error)."); continue; }
            try {
                var vName = victim.name;
                victim.remove(surv);
                nRemoved++;
                lines.push("REMOVED  '" + vName + "'  \u2192 usages remapped to '" + surv.name + "'");
            } catch (eR) {
                fails.push("Could not remove '" + victim.name + "': " + eR + " (mixed ink group? in use by an ink alias?)");
            }
        }

        // Pass 3: renames — after removals. If the final name is already taken
        // by a stray swatch (imported by placed art, or debris from a previous
        // run), MERGE the stray into the survivor, then take the name.
        for (i = 0; i < plan.steps.length; i++) {
            var s3 = plan.steps[i];
            if (s3.op !== "rename" && s3.op !== "renameNew") continue;
            var tgtId = (s3.op === "rename") ? s3.item.id : createdIds[s3.groupKey];
            var tgt = resolveColorByID(docRef, tgtId);
            if (!tgt) continue;
            try {
                var stray = null;
                try { var hit = docRef.colors.itemByName(s3.to); if (hit.isValid && hit.id !== tgt.id) stray = hit; } catch (_) {}
                if (stray) {
                    try {
                        var strayName = stray.name;
                        stray.remove(tgt);
                        nMerged++;
                        lines.push("MERGED   stray '" + strayName + "' into survivor (usages remapped)");
                        stray = null;
                    } catch (eM) {
                        fails.push("Rename skipped: '" + s3.to + "' exists and could not be merged (" + eM + "). Kept '" + tgt.name + "'.");
                    }
                }
                if (!stray) {
                    var old = tgt.name;
                    tgt.name = s3.to;
                    nRenamed++;
                    lines.push("RENAMED  '" + old + "'  \u2192  '" + s3.to + "'");
                }
            } catch (eN) {
                fails.push("Could not rename '" + (tgt.isValid ? tgt.name : "?") + "' to '" + s3.to + "': " + eN);
            }
        }

        // Notes & summary
        if (plan.notes.length) {
            lines.push("");
            lines.push("NOTES:");
            for (i = 0; i < plan.notes.length; i++) lines.push("  " + plan.notes[i]);
        }
        var leftAlone = [];
        for (i = 0; i < inv.length; i++) {
            var a = plan.actions[inv[i].name] || "";
            if (a.indexOf("leave") === 0) leftAlone.push(inv[i].name + "  (" + a + ")");
        }
        if (leftAlone.length) {
            lines.push("");
            lines.push("LEFT UNTOUCHED:");
            for (i = 0; i < leftAlone.length; i++) lines.push("  " + leftAlone[i]);
        }
        if (fails.length) {
            lines.push("");
            lines.push("ISSUES (" + fails.length + "):");
            for (i = 0; i < fails.length; i++) lines.push("  " + fails[i]);
        }
        lines.push("");
        lines.push("Summary: " + nCreated + " created, " + nRemoved + " removed & remapped, " + nMerged + " merged, " + nRenamed + " renamed.");
        lines.push("Reminder: colours inside PLACED artwork (EPS/PDF/AI) are not affected by swatch operations.");
        lines.push("One undo step reverts the whole run.");
        return lines.join("\n");
    }

    function uniqueName(docRef, base) {
        var nm = base, n = 1;
        while (true) {
            var hit = false;
            try { hit = docRef.colors.itemByName(nm).isValid; } catch (_) { hit = false; }
            if (!hit) return nm;
            nm = base + " " + (++n);
        }
    }

    // Resolve a colour by its stable ID — never by index, never from a cached
    // reference. IDs survive removals; indices do not (the v1.0 bug).
    function resolveColorByID(docRef, id) {
        if (id === null || id === undefined) return null;
        try { var c = docRef.colors.itemByID(id); if (c && c.isValid) return c; } catch (_) {}
        try {
            for (var i = 0; i < docRef.colors.length; i++) {
                if (docRef.colors[i].id === id) return docRef.colors[i];
            }
        } catch (_) {}
        return null;
    }

    // ---------------- Report window ----------------

    function showReport(title, text) {
        var w = new Window("dialog", title);
        w.orientation = "column"; w.alignChildren = "fill";
        var et = w.add("edittext", [0, 0, 720, 420], text, { multiline: true, readonly: true, scrolling: true });
        var g = w.add("group"); g.alignment = "right";
        g.add("button", undefined, "Close", { name: "ok" });
        w.show();
    }

})();
