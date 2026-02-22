# Mark's Tools — pyRevit Extension

Cloud model dependency discovery and **bottom-up sync** for ACC / BIM Collaborate Pro.

## What it does

| Button | Description |
|---|---|
| **Get Model ID** | Extracts Region, Project GUID, and Model GUID from the active cloud document. Copies a CSV row to the clipboard. |
| **Sync Tree** | Opens a WPF wizard that: **(1)** reads the active document's linked models, **(2)** recursively discovers the full dependency tree (opens each model detached for speed), **(3)** topologically sorts them leaf → root, and **(4)** syncs every selected model bottom-up so children are always up-to-date before their parents. |

## Install

### Option A — pyRevit CLI (recommended)

```bash
# 1. Register the custom extension source (one-time)
pyrevit extensions sources add MarksTools "https://raw.githubusercontent.com/CHANGEME/MarksTools/main/extensions.json"

# 2. Install
pyrevit extensions install MarksTools
```

Or install directly from the Git URL:

```bash
pyrevit extensions install MarksTools --git "https://github.com/CHANGEME/MarksTools.git"
```

After installing, **reload pyRevit** (pyRevit tab → Reload) or restart Revit.

### Option B — Shared network folder

1. Copy the `MarksTools.extension` folder to a shared drive, e.g.:  
   `\\server\share\pyRevitExtensions\MarksTools.extension`
2. In Revit → **pyRevit tab → Settings → Custom Extension Directories** → add:  
   `\\server\share\pyRevitExtensions`
3. Reload pyRevit.

### Option C — Manual clone

```bash
cd %appdata%\pyRevit\Extensions
git clone https://github.com/CHANGEME/MarksTools.git
```

## Requirements

- Revit **2024** or later (tested on 2026)
- **pyRevit 4.8+**
- ACC / BIM Collaborate Pro cloud models

## Extension structure

```
MarksTools.extension/
├── extension.json
├── lib/
│   ├── __init__.py
│   ├── cloud_helpers.py      ← open / sync / close helpers
│   ├── dependency_tree.py    ← BFS discovery + topological sort
│   └── guid_extractor.py     ← extract GUIDs from a Document
│
└── Mark's Tools.tab/
    └── Cloud Models.panel/
        ├── Get Model ID.pushbutton/
        │   ├── script.py
        │   └── bundle.yaml
        └── Sync Tree.pushbutton/
            ├── script.py
            ├── SyncWizard.xaml
            └── bundle.yaml
```

## Customisation

- **Sync options** (workset opening, link reload delay, etc.) are configurable in the wizard's Step 3 panel — no code edits needed.
- **Icons**: drop a `icon.png` (32×32 or 96×96) into each `.pushbutton` folder for a custom ribbon icon.

## Licence

MIT
