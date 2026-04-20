# Module Frame Renamer
### Adobe InDesign ExtendScript — `.jsx`

Automatically detects and renames frames on the active page (or a selection) by their physical dimensions, following a structured module naming convention.

---

## Output Format

```
Module001_FPFP_001
Module002_FPHP_001
Module003_FPHP_002
Module004_FPQP_001
```

| Segment | Example | Description |
|---|---|---|
| `Module` | `Module001` | Global module counter, zero-padded to 3 digits |
| `PageCode` | `FP` | User-defined page size code (e.g. FP = Full Page) |
| `SizeCode` | `HP` | User-defined module size code (e.g. HP = Half Page) |
| `SizeCount` | `001` | Per-size counter, zero-padded to 3 digits |

---

## Installation

1. Copy `ModuleFrameRenamer.jsx` into your InDesign scripts folder:

   | Platform | Path |
   |---|---|
   | **Mac** | `~/Library/Preferences/Adobe InDesign/[Version]/[Locale]/Scripts/Scripts Panel/` |
   | **Windows** | `%AppData%\Adobe\InDesign\[Version]\[Locale]\Scripts\Scripts Panel\` |

2. In InDesign, open the Scripts panel via **Window → Utilities → Scripts**.
3. Double-click `ModuleFrameRenamer.jsx` to run.

---

## How It Works

The script runs in two passes before any renaming occurs.

**Pass 1 — Frame Detection**
Before the dialog opens, the script scans all frames in the source (selection or active page) and groups them by their physical dimensions, using a 0.5 mm tolerance to catch minor placement variation. Groups are sorted largest to smallest by area.

**Pass 2 — User Dialog**
The dialog presents each detected size group as a pre-filled row showing the width, height, and number of matching frames. The user types a short code next to each size they want renamed. Rows left blank are skipped.

Once confirmed, frames are sorted top → bottom, left → right, then renamed sequentially.

---

## The Dialog

```
┌─ Module Frame Renamer ──────────────────────────────────┐
│                                                          │
│  Source:  active page (1)                                │
│                                                          │
│  Page size code:  [ FP ]                                 │
│  Match tolerance (mm):  [ 0.5 ]                          │
│  ☑ Also apply name to Script Label                       │
│                                                          │
│ ┌─ Detected Frame Sizes ─────────────────────────────┐  │
│ │  Code      Width (mm)   Height (mm)   Frames        │  │
│ │  [     ]   210.0        297.0         1 frame        │  │
│ │  [     ]   210.0        148.5         2 frames       │  │
│ │  [     ]   105.0        148.5         4 frames       │  │
│ │  [     ]   70.0         99.0          8 frames       │  │
│ └─────────────────────────────────────────────────────┘  │
│                                                          │
│                          [ Cancel ]  [ Run ]             │
└──────────────────────────────────────────────────────────┘
```

### Fields

| Field | Description |
|---|---|
| **Page size code** | Short code representing the overall page format. Becomes the first code segment in every name (e.g. `FP` for Full Page broadsheet). |
| **Match tolerance** | How many mm of size variation is allowed when matching frames to a size group. Default `0.5`. |
| **Apply name to Script Label** | When ticked, sets `item.label` to the new name as well as `item.name`. Useful for scripts and data merge workflows that read the Script Label. On by default. |
| **Code column** | Type a short code next to each detected size to include it. Leave blank to skip that size entirely. |

---

## Source Priority

| Condition | Scope |
|---|---|
| One or more items selected | Runs on the **selection only** |
| Nothing selected | Runs on all items on the **active page** |

If a Group is selected, the script flattens and processes all child items within it.

---

## Naming Example

Given a page code of `FP` and the following codes assigned in the dialog:

| Code | Size |
|---|---|
| `FP` | 210 × 297 mm |
| `HP` | 210 × 148.5 mm |
| `QP` | 105 × 148.5 mm |

Frames are sorted top → bottom, left → right, then named:

```
Module001_FPFP_001    ← first Full Page frame
Module002_FPHP_001    ← first Half Page frame
Module003_FPHP_002    ← second Half Page frame
Module004_FPQP_001    ← first Quarter Page frame
Module005_FPQP_002    ← second Quarter Page frame
```

---

## Summary Report

After running, an alert displays a full report:

```
── Module Rename Complete ──────────────────
Source: active page (1)
Script label: ✅ applied

✅  Renamed: 5 element(s)

  Module001_FPFP_001   (210.0 × 297.0 mm)
  Module002_FPHP_001   (210.0 × 148.5 mm)
  Module003_FPHP_002   (210.0 × 148.5 mm)
  Module004_FPQP_001   (105.0 × 148.5 mm)
  Module005_FPQP_002   (105.0 × 148.5 mm)

⚠  Skipped: 1 element(s)
  "rectangle"  34.0 × 12.0 mm  — no code assigned

── Size breakdown ──
  FPFP:  1 frame(s)
  FPHP:  2 frame(s)
  FPQP:  2 frame(s)
```

Frames are only listed as skipped if their size did not match any code-assigned group. Items with no code assigned are silently ignored.

---

## Re-running the Script

Running the script a second time on an already-named page is safe. All frames are renumbered from scratch — existing names are overwritten regardless of their current value.

---

## Supported InDesign Versions

- InDesign CC 2018 and later (tested on CC 2021+)
- ExtendScript engine only — not UXP
- Units: document ruler units are auto-detected and converted to mm internally

---

## Known Limitations

- Processes one page at a time — multi-page batch runs are not supported
- Frame detection uses `geometricBounds` — rotated frames are measured by their bounding box, not their visual size
- Frames smaller than 1 × 1 mm are ignored during size detection

---

## Notes for Developers

The script uses `item.label` (the InDesign Script Label, accessible via **Window → Utilities → Scripts → Script Label panel**) as an optional parallel identifier. This is useful when downstream scripts or data merge workflows need to locate frames by a stable key that survives layer reordering.

All counters use zero-padded 3-digit formatting (`001`, `002` … `999`). If you need wider padding for very large layouts, change the second argument in the `pad()` calls at the top of the script.
