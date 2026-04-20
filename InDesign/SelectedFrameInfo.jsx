﻿﻿/**
 * Report & Edit Selected Frames (Label • Type • Object Style • X • Y • Page)
 * InDesign ExtendScript (2023+ Mac/Win)
 *
 * NEW: Optional locator checkbox to add a **red 2pt border** to the selected frame
 * (useful for quickly finding an object on the page). Non-destructive and applied only
 * when you click **Apply to Document** for the currently selected row.
 *
 * What it does
 * - Scans the current selection for frames (TextFrames + Rect/Polygon/Oval containers) and builds rows:
 *   Script Label | Type | Object Style | X | Y | Page
 * - Dialog preview with:
 *     • Copy to Clipboard (TSV)  • Save CSV  • Place report on page (styled table)
 *     • **Edit pane** to update the selected row's Script Label and Object Style
 *     • **Add red 2pt border** tick box for quick visual identification
 * - Optional on-page table with styles; any "n/a" cells are formatted **bold red**.
 *
 * Safety / Notes
 * - Non-destructive defaults; single undo step. Clipboard uses a temp doc (closed without saving).
 * - Frame-type switching has been **removed** per request.
 * - X/Y are taken from geometricBounds top-left [Y1, X1] and shown in current ruler units.
 *
 * Install: Save as `Report_Selected_Frames_Edit_NoType_XY_Locator.jsx` in your Scripts Panel.
 */

#target "InDesign"
#targetengine "session"

(function () {
    if (!app.documents.length) { alert("Open a document and select some frames first."); return; }

    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Report & Edit Selected Frames (No Type, XY, Locator)");

    function main() {
        var doc = app.activeDocument;
        var items = collectFramesFromSelection(app.selection);
        if (!items.length) { alert("No valid frames found in selection."); return; }

        var rows = buildRows(doc, items);
        rows.sort(rowSorter);

        var placeOnPageRef = { value: false };
        showDialog(doc, rows, placeOnPageRef);

        if (placeOnPageRef.value) {
            placeReportTable(doc, rows);
            alert("Report placed on a new last page.");
        }
    }

    // ---------------- Collect & Rows ----------------

    function collectFramesFromSelection(selArray) {
        var out = [], seen = {};
        function addUnique(it){ if (it && it.isValid){ var k = it.id+":"+(it.constructor?it.constructor.name:""); if(!seen[k]){seen[k]=true; out.push(it);} } }
        function descend(item){
            try{
                if(!item || !item.isValid) return; var cn = item.constructor?item.constructor.name:"";
                if (cn==="InsertionPoint"||cn==="Character"||cn==="Text"){ if(item.parentTextFrames&&item.parentTextFrames.length)addUnique(item.parentTextFrames[0]); return; }
                if (cn==="TextFrame") addUnique(item);
                if (cn==="Rectangle"||cn==="Polygon"||cn==="Oval"||cn==="Ellipse") addUnique(item);
                if (item.hasOwnProperty("pageItems")) { var kids=item.pageItems; for (var i=0;i<kids.length;i++) descend(kids[i]); }
            }catch(_){ }
        }
        for (var i=0;i<selArray.length;i++) descend(selArray[i]);
        return out;
    }

    function buildRows(doc, items){
        var rows=[]; for (var i=0;i<items.length;i++) rows.push(rowFromItem(doc, items[i])); return rows;
    }

    function rowFromItem(doc, item){
        var label = safeString(item.label); label = (label==="")?"n/a":label;
        var type = detectType(item);
        var style = detectObjectStyleName(item);
        var pg = detectPage(item);
        var xy = detectXY(doc, item);
        return { scriptLabel:label, type:type, objectStyle:style, x:xy.x, y:xy.y, page:pg.display, _pageIndex:pg.index, _ref:item };
    }

    function rowSorter(a,b){ if(a._pageIndex!==b._pageIndex) return a._pageIndex-b._pageIndex; if(a.type!==b.type) return (a.type==="Text")?-1:1; return toLowerSafe(a.scriptLabel).localeCompare(toLowerSafe(b.scriptLabel)); }

    function detectType(item){ try{ var cn=item.constructor?item.constructor.name:""; if(cn==="TextFrame") return "Text"; if(item.allGraphics&&item.allGraphics.length>0) return "Image"; if((cn==="Rectangle"||cn==="Polygon"||cn==="Oval"||cn==="Ellipse") && item.images && item.images.length>0) return "Image"; }catch(_){ } return "Text"; }
    function detectObjectStyleName(item){ try{ if(item.appliedObjectStyle&&item.appliedObjectStyle.isValid){ var nm=item.appliedObjectStyle.name; return (nm&&nm.length)?nm:"[None]"; } }catch(_){ } return "n/a"; }
    function detectPage(item){ var display="n/a", idx=999999; try{ if(item.parentPage&&item.parentPage.isValid){ display=item.parentPage.name; idx=item.parentPage.documentOffset; return {display:display,index:idx}; } if(item.parent&&item.parent.constructor&&item.parent.constructor.name==="Spread"){ var spr=item.parent; if(spr.pages&&spr.pages.length){ display=spr.pages[0].name; idx=spr.pages[0].documentOffset; return {display:display,index:idx}; } } }catch(_){ } return {display:display,index:idx}; }
    function detectXY(doc, item){
        try {
            var gb = item.geometricBounds; // [y1, x1, y2, x2]
            var y = gb[0]; var x = gb[1];
            return { x: formatUnits(doc, x), y: formatUnits(doc, y) };
        } catch(_) { return { x:"n/a", y:"n/a" }; }
    }

    function formatUnits(doc, pts){
        // Convert points to current ruler units (string with unit suffix)
        var u = doc.viewPreferences.horizontalMeasurementUnits; // use horizontal for both
        function round(n, p){ var f=Math.pow(10,p); return Math.round(n*f)/f; }
        switch (u) {
            case MeasurementUnits.MILLIMETERS: return round(pts*0.3527777778, 2) + " mm";
            case MeasurementUnits.CENTIMETERS: return round(pts*0.03527777778, 2) + " cm";
            case MeasurementUnits.INCHES:      return round(pts/72, 3) + " in";
            case MeasurementUnits.POINTS:      return round(pts, 2) + " pt";
            case MeasurementUnits.PICAS:       return round(pts/12, 2) + " p"; // decimal picas
            case MeasurementUnits.Q:           return round((pts*0.3527777778)/0.25, 1) + " Q"; // 1Q = 0.25 mm
            default:                           return round(pts, 2) + " pt";
        }
    }

    // ---------------- Dialog (Label + Object Style only + locator checkbox) ----------------

    function showDialog(doc, rows, placeOnPageRef){
        var dlg = new Window("dialog", "Selected Frames – Report & Edit (Locator)");
        dlg.orientation = "column"; dlg.alignChildren = "fill";

        var list = dlg.add("listbox", [0,0,880,360], [], {numberOfColumns:7, showHeaders:true,
            columnTitles:["#","Script Label","Type","Object Style","X","Y","Page"],
            columnWidths:[40,260,90,220,90,90,90]});
        function populateList(){ list.removeAll(); for(var i=0;i<rows.length;i++){ var r=rows[i]; var it=list.add("item", (i+1)); it.subItems[0].text=r.scriptLabel; it.subItems[1].text=r.type; it.subItems[2].text=r.objectStyle; it.subItems[3].text=r.x; it.subItems[4].text=r.y; it.subItems[5].text=r.page; it._rowIndex=i; } }
        populateList();

        // --- Edit pane ---
        var editPanel = dlg.add("panel", undefined, "Edit selection"); editPanel.alignChildren = "left"; editPanel.orientation="column";
        var g1 = editPanel.add("group"); g1.add("statictext", undefined, "Script Label:"); var etLabel = g1.add("edittext", [0,0,360,24], "");
        var g2 = editPanel.add("group"); g2.add("statictext", undefined, "Object Style:"); var ddStyle = g2.add("dropdownlist", [0,0,360,24], allObjectStyleNames(doc)); ddStyle.selection = 0;
        var g3 = editPanel.add("group"); var chkBorder = g3.add("checkbox", undefined, "Add red 2pt border (locator)"); chkBorder.value = false;
        var gBtns = editPanel.add("group"); var btnApply = gBtns.add("button", undefined, "Apply to Document");

        // Options & main buttons
        var opts = dlg.add("group"); var chkPlace = opts.add("checkbox", undefined, "Place report on page (new last page) – 'n/a' will be bold red"); chkPlace.value=false;
        var btnRow = dlg.add("group"); btnRow.alignment="right"; var btnCopy = btnRow.add("button", undefined, "Copy to Clipboard"); var btnSave = btnRow.add("button", undefined, "Save CSV…"); var btnClose = btnRow.add("button", undefined, "Close", {name:"ok"});

        // Selection change → load fields
        list.onChange = function(){ var it=list.selection; if(!it) return; var r=rows[it._rowIndex]; etLabel.text = (r.scriptLabel==="n/a")?"":r.scriptLabel; setStyleDropdown(ddStyle, r.objectStyle, doc); };

        btnApply.onClick = function(){
            var it = list.selection; if(!it){ alert("Select a row to edit."); return; }
            var idx = it._rowIndex; var row = rows[idx]; var item = row._ref; if(!item || !item.isValid){ alert("The underlying item is no longer valid."); return; }
            // Script Label
            try{ item.label = etLabel.text; row.scriptLabel = etLabel.text && etLabel.text.length ? etLabel.text : "n/a"; }catch(e){ alert("Could not set Script Label: \n"+e); }
            // Object Style (from full list)
            try{
                var chosenName = ddStyle.selection ? ddStyle.selection.text : "[None]";
                var os = findObjectStyleByName(doc, chosenName);
                if (os && os.isValid) item.appliedObjectStyle = os;
                row.objectStyle = detectObjectStyleName(item);
            }catch(e2){ alert("Could not assign Object Style: \n"+e2); }
            // Optional locator border
            try{
                if (chkBorder.value) applyLocatorBorder(doc, item);
            }catch(e3){ alert("Could not apply locator border: \n"+e3); }

            // Refresh XY & page and list
            var pg = detectPage(item); var xy = detectXY(doc, item); row.x = xy.x; row.y = xy.y; row.page=pg.display; row._pageIndex=pg.index; populateList(); list.selection = list.items[idx];
        };

        btnCopy.onClick = function(){ try{ copyStringToClipboard(rowsToTSV(rows)); alert("Copied as TSV. Paste into Excel/Sheets/Notes."); }catch(e){ alert("Copy failed:\n"+e);} };
        btnSave.onClick = function(){ try{ var f=File.saveDialog("Save report as CSV","CSV:*.csv"); if(!f) return; var csv=rowsToCSV(rows); f.encoding="UTF-8"; f.open("w"); f.write("\uFEFF"+csv); f.close(); alert("CSV saved: \n"+f.fsName);}catch(e){ alert("Save failed:\n"+e);} };
        btnClose.onClick = function(){ placeOnPageRef.value = chkPlace.value===true; dlg.close(1); };

        dlg.show();
    }

    // ---------------- Locator border ----------------

    function applyLocatorBorder(doc, item){
        if (item.locked || (item.itemLayer && item.itemLayer.locked)){
            alert("Item or its layer is locked; cannot add border.");
            return;
        }
        var red = getOrCreateColor(doc, "Locator Red", 0, 100, 100, 0); // CMYK red (works in RGB docs too)
        try { item.strokeWeight = 2; } catch(_){ }
        try { item.strokeTint = 100; } catch(_){ }
        try { item.strokeColor = red; } catch(_){ }
        try { item.transparencySettings.blendingSettings.opacity = item.transparencySettings.blendingSettings.opacity; } catch(_){ }
    }

    // ---------------- Object Style helpers ----------------

    function allObjectStyleNames(doc){
        var names = [];
        try { if (doc.objectStyles.itemByName("[None]").isValid) names.push("[None]"); } catch(_){ }
        function walk(container){
            try {
                var i;
                for (i=0; i<container.objectStyles.length; i++){
                    var n = container.objectStyles[i].name;
                    if (n !== "[None]") names.push(n);
                }
                for (i=0; i<container.objectStyleGroups.length; i++){
                    walk(container.objectStyleGroups[i]);
                }
            } catch(_){ }
        }
        walk(doc);
        var seen = {}, out = [], j;
        for (j=0; j<names.length; j++){ if (!seen[names[j]]){ seen[names[j]] = 1; out.push(names[j]); } }
        if (out.length && out[0] === "[None]"){
            var head = out.shift(); out.sort(); out.unshift(head);
        } else {
            out.sort();
        }
        return out;
    }

    function setStyleDropdown(dd, currentName, doc){
        dd.removeAll(); var arr = allObjectStyleNames(doc);
        for (var i=0; i<arr.length; i++) dd.add("item", arr[i]);
        var sel = 0; for (var j=0; j<arr.length; j++){ if (arr[j] === currentName){ sel = j; break; } }
        dd.selection = sel;
    }

    function findObjectStyleByName(doc, name){
        try { var os = doc.objectStyles.itemByName(name); if (os && os.isValid) return os; } catch(_){ }
        function walk(container){
            var i, s;
            for (i=0; i<container.objectStyles.length; i++){ s = container.objectStyles[i]; if (s.name === name) return s; }
            for (i=0; i<container.objectStyleGroups.length; i++){ s = walk(container.objectStyleGroups[i]); if (s) return s; }
            return null;
        }
        return walk(doc);
    }

    // ---------------- Clipboard & CSV/TSV ----------------

    function rowsToTSV(rows){ var out=[]; out.push(["Script Label","Type","Object Style","X","Y","Page"].join("\t")); for(var i=0;i<rows.length;i++){ var r=rows[i]; out.push([r.scriptLabel,r.type,r.objectStyle,r.x,r.y,r.page].join("\t")); } return out.join("\r"); }
    function rowsToCSV(rows){ var lines=[]; lines.push(csvLine(["Script Label","Type","Object Style","X","Y","Page"])); for(var i=0;i<rows.length;i++){ var r=rows[i]; lines.push(csvLine([r.scriptLabel,r.type,r.objectStyle,r.x,r.y,r.page])); } return lines.join("\r\n"); function csvLine(arr){ var cells=[]; for (var j=0; j<arr.length; j++){ cells.push(csvEscape(arr[j])); } return cells.join(","); } function csvEscape(val){ var s=(val===null||val===undefined)?"":String(val); if(/[",\r\n]/.test(s)) s='"'+s.replace(/"/g,'""')+'"'; return s; } }
    function copyStringToClipboard(str){ var tmpDoc=app.documents.add(); try{ var tf=tmpDoc.textFrames.add(); tf.geometricBounds=[20,20,200,400]; tf.contents=str; tf.texts[0].select(); app.copy(); } finally { tmpDoc.close(SaveOptions.NO); } }

    function toLowerSafe(s){ return (typeof s==="string")?s.toLowerCase():""; }
    function safeString(v){ return (typeof v==="string")?v:"n/a"; }

    // ---------------- Place on page (styled; bold red for n/a) ----------------

    function placeReportTable(doc, rows){
        var pg = doc.pages.add(LocationOptions.AT_END);
        var mp = pg.marginPreferences; var gb=[mp.top, mp.left, doc.documentPreferences.pageHeight-mp.bottom, doc.documentPreferences.pageWidth-mp.right];
        var tf = pg.textFrames.add(); tf.geometricBounds=gb; tf.contents = rowsToTSV(rows);
        var tab = tf.parentStory.texts[0].convertToTable("\t","\r");

        var pBody = getOrCreateParagraphStyle(doc, "Report/Body", { pointSize:10, leading:12 });
        var pHead = getOrCreateParagraphStyle(doc, "Report/Header", { pointSize:10.5, leading:12.5, capitalization:Capitalization.ALL_CAPS });
        var pNA   = getOrCreateParagraphStyle(doc, "Report/n-a", { pointSize:10, leading:12, fillColor:getOrCreateColor(doc,"Report Red", 0,100,100,0), fontStyle:"Bold" });
        var cTH   = getOrCreateCellStyle(doc, "Report/TH", { appliedParagraphStyle:pHead, topEdgeStrokeWeight:0.5, bottomEdgeStrokeWeight:1 });
        var cTD   = getOrCreateCellStyle(doc, "Report/TD", { appliedParagraphStyle:pBody, topEdgeStrokeWeight:0.5, bottomEdgeStrokeWeight:0.5 });
        var tStyle= getOrCreateTableStyle(doc, "Report/Table");
        try{ tStyle.headerRegionCellStyle=cTH; tStyle.bodyRegionCellStyle=cTD; tStyle.tableStrokeWeight=0.5; }catch(_){ }
        tab.appliedTableStyle=tStyle; tab.headerRowCount=1; tab.rows[0].rowType=RowTypes.HEADER_ROW;

        // Column widths for 7 columns: #, Label, Type, Style, X, Y, Page
        try{ var colPerc=[8,32,12,20,10,10,8]; var frameW=tf.geometricBounds[3]-tf.geometricBounds[1]; for(var c=0;c<tab.columns.length;c++) tab.columns[c].width=(colPerc[c]/100)*frameW; }catch(_){ }
        // Insets
        for (var r=0;r<tab.rows.length;r++){ for (var cc=0; cc<tab.columns.length; cc++){ var cell=tab.rows[r].cells[cc]; cell.leftInset=cell.rightInset=3; cell.topInset=cell.bottomInset=2; } }
        // Bold red for n/a cells in body
        for (var rr=1; rr<tab.rows.length; rr++){
            var row = tab.rows[rr];
            for (var cc2=1; cc2<row.cells.length; cc2++){
                var t = row.cells[cc2].contents;
                if (String(t).toLowerCase()==="n/a") row.cells[cc2].texts[0].appliedParagraphStyle = pNA;
            }
        }
    }

    // ---------------- Style helpers ----------------

    function getOrCreateParagraphStyle(doc, name, props){ var s=doc.paragraphStyles.itemByName(name); if(!s.isValid) s=doc.paragraphStyles.add({name:name}); applyProps(s,props); return s; }
    function getOrCreateCellStyle(doc, name, props){ var s=doc.cellStyles.itemByName(name); if(!s.isValid) s=doc.cellStyles.add({name:name}); applyProps(s,props); return s; }
    function getOrCreateTableStyle(doc, name){ var s=doc.tableStyles.itemByName(name); if(!s.isValid) s=doc.tableStyles.add({name:name}); return s; }
    function getOrCreateColor(doc, name, c,m,y,k){ try{ var s=doc.colors.itemByName(name); if(s.isValid) return s; }catch(_){ } return doc.colors.add({name:name, model:ColorModel.PROCESS, space:ColorSpace.CMYK, colorValue:[c,m,y,k]}); }
    function applyProps(obj, props){ if(!props) return; for (var k in props) if(props.hasOwnProperty(k)){ try{ obj[k]=props[k]; }catch(_){ } } }

})();
