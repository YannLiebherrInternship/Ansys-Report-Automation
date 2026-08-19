# Ansys Mechanical – PowerPoint Report Generator

IronPython 2.7 script executed directly in the **Ansys Mechanical 2025 R2 scripting console**. It opens a WPF window in which the engineer selects the model elements (geometry, mesh, boundary conditions, contacts, results...) to include in the report, then automatically generates a PowerPoint presentation from a corporate template, archiving the extracted data as CSV along the way.

> **Language note.** This is the English translation of `README.md`. The graphical interface (buttons, tabs, fields, messages) has been in English since 2026-08-19; code comments and the primary documentation (`README.md`) remain in French. This file is kept in sync with `README.md` but is the reference for English-speaking readers.

## Table of contents

1. [Prerequisites](#1-prerequisites)
2. [Installing the tool in an Ansys project](#2-installing-the-tool-in-an-ansys-project)
3. [Repository structure](#3-repository-structure)
4. [Data pipeline](#4-data-pipeline)
5. [Modules `00_constants.py` → `05_interactive_slides.py`](#5-modules-00_constantspy--05_interactive_slidespy)
6. [WPF interface (`AnsysReportGenerator_WPF.py` / `.xaml`)](#6-wpf-interface)
7. [Ansys domain concepts used in the code](#7-ansys-domain-concepts-used-in-the-code)
8. [Ansys Mechanical APIs used](#8-ansys-mechanical-apis-used)
9. [How the code drives PowerPoint (COM Interop)](#9-how-the-code-drives-powerpoint-com-interop)
10. [Python design choices used in the project](#10-python-design-choices-used-in-the-project)
11. [Python fundamentals illustrated by the project's code](#11-python-fundamentals-illustrated-by-the-projects-code)
12. [Creating a new custom slide in the Master Template](#12-creating-a-new-custom-slide-in-the-master-template)
13. [Known pitfalls / technical choices](#13-known-pitfalls--technical-choices)

---

## 1. Prerequisites

| Item | Detail |
|---|---|
| Ansys Mechanical | 2025 R2 (provides the embedded IronPython 2.7 + the `ExtAPI`/`DataModel` API) |
| Microsoft Office | PowerPoint installed (COM Interop `Microsoft.Office.Interop.PowerPoint`) |
| PowerPoint template | A corporate `.pptx` file with the expected custom layouts (see §5, `00_constants.py`) |
| System | Windows (Windows Forms + WPF via .NET, COM Interop) |

No external Python dependency is required: everything goes through the IronPython standard library (`os`, `csv`, `re`, `datetime`, `shutil`, `xml.etree.ElementTree`) and .NET assemblies loaded via `clr.AddReference()`.

## 2. Installing the tool in an Ansys project

The generator is not installed once and for all: every Ansys project must receive its own copy of the application folder, because the script locates itself relative to the currently open project. The steps are as follows.

First open the relevant Ansys project and save it at least once. Saving creates a folder called `<ProjectName>_files`, which itself contains a standard `user_files` subfolder. It is next to this `user_files` folder — not inside it — that a new folder named exactly `Report Generator` must be created. Inside this `Report Generator` folder, copy flat, with no subfolder, all of the following files: the entry point `AnsysReportGenerator_WPF.py`, its layout `AnsysReportGenerator_WPF.xaml`, the six modules `00_constants.py` through `05_interactive_slides.py`, and the corporate PowerPoint template under the exact name `Master Template_def.pptx`. Nothing else is needed: the data subfolders (`data/image_export`, `data/csv_export`, `data/reports`, `data/legend`, `data/export_3D`) don't exist yet at this stage and will be created automatically by the script on its first run.

Once this folder is in place, open the scripting console in Ansys Mechanical (**File > Scripting > Run Script** menu, or the equivalent depending on the version), and select `AnsysReportGenerator_WPF.py` inside the `Report Generator` folder you just created. This is the only file to run manually: as soon as it starts, it locates the `Report Generator` folder on its own by querying the Ansys API (`ExtAPI.DataModel.Project.ProjectDirectory`, which always returns the `<ProjectName>_files` folder of the currently open project, regardless of the machine or where the project was saved), then loads the six modules `00_constants.py` through `05_interactive_slides.py` in order with `execfile()`, builds the window from the `.xaml`, and displays it. If the `Report Generator` folder is missing, incomplete, or if the Ansys project has never been saved, the script stops immediately with an explicit error message rather than crashing further down without explanation.

No path needs to be changed in the code for this to work on a new project or a new machine: all working paths (images folder, CSV folder, reports folder, legends folder, 3D export folder) are automatically recalculated from the location of `Report Generator`. Only the PowerPoint template must be provided manually, since it's a content file that cannot be generated automatically; if it's missing at load time, a warning is shown in the console, and report generation will fail cleanly (clear message, no crash) until a valid template has been placed in the right spot. To deploy the tool on a new project, simply copy the entire `Report Generator` folder (`.py`/`.xaml` files + the `.pptx`) next to `user_files` in the new `<ProjectName>_files`.

These paths can also be changed without restarting the script, once the window is open, from the "Files" tab of the interface (with a "Reset" button to go back to the automatically computed values).

> **Other installation modes.** The mode described above (manual script execution via the scripting console) is currently the only one available. Two other launch modes — a button pinned to Ansys Mechanical's "Automation" toolbar, and a published Ansys extension installable via the Extension Manager — are under development and will be documented here once available.

## 3. Repository structure

The folder of this repository corresponds to the contents of the `Report Generator` folder to deploy as described in §2: everything else must be copied next to `user_files` in `<ProjectName>_files`.

The entry point is `AnsysReportGenerator_WPF.py`, the only file to run directly in Mechanical. It relies on `AnsysReportGenerator_WPF.xaml`, which describes the layout of the main window (tabs, cards, styles). Next come the six modules loaded in order by `execfile()`: `00_constants.py` groups paths, layout constants, and generic utility functions (files, CSV); `01_data_export.py` extracts data from the model and the Tabular Data pane to CSV; `02_image_export.py` captures viewport images and rebuilds certain charts; `03_ppt_utils.py` defines the `PPTReportBuilder` class, which owns the PowerPoint COM session and exposes slide-adding methods; `04_slides.py` contains the "legacy" slide builders, which systematically process every object of a category; `05_interactive_slides.py` is the largest module and provides all the interactive-selection logic used by the interface. The `Master Template_def.pptx` template must exist at the root of the folder, next to these files; it's the only path that isn't created automatically.

A `data/` folder is created automatically on first launch (it's absent at first on a new project) and contains five subfolders: `image_export/` for exported PNG images (viewport and rebuilt charts), `csv_export/` for CSVs archived independently of PowerPoint, `reports/` for the template's working copies and the generated `.pptx` reports, `legend/` for `.xml` legend files the user can drop in manually (empty at first), and `export_3D/` for the `.avz` files (interactive 3D views) generated by the "Export to 3D" button.

`04_slides.py` and `05_interactive_slides.py` deliberately coexist: `04_slides.py` provides the original `create_..._slide` functions, which always export everything with no possible configuration, and `05_interactive_slides.py` reuses them as building blocks to construct versions filtered by the user's selection (`build_..._slides`), without duplicating the CSV/image extraction logic already written. The WPF application only calls functions from `05_interactive_slides.py`, except for `create_geometry_slide` and `create_analysis_parameters_slide` from `04_slides.py`, reused as-is.

## 4. Data pipeline

Generating a report always follows the same path, whatever the slide category involved. Everything starts from the Mechanical model: either the object tree and the Tabular Data pane for tabular data, or the 3D viewport for images. On the data side, export functions such as `export_active_tabular_data` or the various `export_*_csv` functions in `01_data_export.py` read this data and write it as CSV files in `CSV_EXPORT_FOLDER`. On the visual side, the functions `export_current_view_image`, `export_object_image`, and `export_chart_image_from_csv` in `02_image_export.py` capture or rebuild an image and write it as PNG in `IMAGE_EXPORT_FOLDER`.

These CSV and PNG files then serve as raw material for `PPTReportBuilder`, in `03_ppt_utils.py`: on opening, it starts by making a working copy of the template in `REPORT_OUTPUT_FOLDER` (never the original template), opens this copy via COM Interop, then each call to an `add_..._slide` method adds a slide to that same presentation, inserting the image and/or table read from the CSV/PNG files produced in the previous step. Once all slides have been added, the presentation is saved under its final name in `REPORT_OUTPUT_FOLDER`, which constitutes the report delivered to the user.

The CSV is always kept on disk, regardless of whether it was successfully inserted into the PowerPoint: it remains viewable and downloadable from the "Files" tab of the interface, and constitutes an archive usable separately from the report. Its insertion as a PowerPoint table is simply skipped if the table exceeds `MAX_TABLE_ROWS`/`MAX_TABLE_COLUMNS` (`00_constants.py`), a table that large becoming unreadable once inserted into a slide.

## 5. Modules `00_constants.py` → `05_interactive_slides.py`

### `00_constants.py`
Root paths, indexes of the template's custom layouts (`LAYOUT_IMAGE_TABLE`, `LAYOUT_TABLE_ONLY`, `LAYOUT_MESH_MULTI`), table display limits, and generic helpers independent of Ansys: `ensure_folder_exists`, `safe_file_name`, `get_unique_file_path`, `clean_cell_text`, `to_csv_cell`. Must be run first — all the constants it defines (`IMAGE_EXPORT_FOLDER`, `CSV_EXPORT_FOLDER`, etc.) are used as-is (global variables, no `import`) by every other module.

### `01_data_export.py`
Everything that reads the **Tabular Data pane** or the model and writes a CSV: tabular data of an active object (`export_active_tabular_data`), contact summary, mesh report, materials used (via the Ansys `materials` module), step-parameter and solution-info tables (`export_analysis_settings_csv`/`export_solution_info_csv`, used in the "Analysis Parameters" slide).

### `02_image_export.py`
Image capture of the Mechanical viewport (`export_current_view_image`, based on `ExtAPI.Graphics.ExportImage`), and "high-level" export per object type (geometry, mesh, analysis overview, any object via a `Figure` snapshot). Also contains a minimal 2D chart-drawing engine in `System.Drawing` (`export_chart_image_from_csv`): "Solution Information" trackers have no 3D representation, so their chart is redrawn from the exported CSV rather than captured from the viewport.

### `03_ppt_utils.py`
**`PPTReportBuilder`** class: encapsulates the single COM PowerPoint session opened on the template's working copy, and exposes high-level methods to add a slide (`add_image_table_slide`, `add_table_slide`, `add_analysis_context_slide`, `add_csv_table`, `save_as`, `close`). See §9 for the details of its internal workings.

### `04_slides.py`
"Legacy" `create_..._slide(report)` functions: each processes **every** object of a model category (no selection/configuration possible). Used by the UI for Geometry and Analysis Context (standalone checkboxes, no checklist).

### `05_interactive_slides.py`
The largest module (~1500 lines). Provides all the support logic for the interface:
- **Cleanup**: `remove_stale_figures()` (deletes leftover `Figure` objects from a previous generation).
- **3D export (.avz)**: `export_all_3d_views()` — for each analysis in the project, exports every simple result (`collect_all_results`) and every child of a Contact Tool / Bolt Tool under the *Solution* branch (`collect_contact_tool_results`, `collect_bolt_tool_results`) to `.avz` (`ExtAPI.Graphics.ModelViewManager.Capture3DImage`) in `EXPORT_3D_FOLDER`.
- **Per-row view / section / scale / legend**: `apply_view_if_exists`, `apply_section_plane`, `apply_scale_factor`, `apply_legend_if_exists` — applied right before capturing an image of a given object, then reset right after.
- **Steps and combined slides**: a result can be exported step by step (`evaluate_result_for_step`, driven by `SetDriverStyle.ResultSet` + `SetNumber`) either as one slide per step, or as a single "combined" multi-image slide (`add_multi_step_image_slide`) if a dedicated template exists for that exact number of steps (`MULTI_STEP_SLIDE_TEMPLATES`: 2, 3, 4, 6, or 8 steps — any other count automatically falls back to individual mode).
- **`*RowConfig` classes** (`SlideRowConfig`, `GeometryPartRowConfig`, `MeshPartRowConfig`, `ContactRowConfig`, `SolutionInfoRowConfig`, `AnalysisContextRowConfig`): one instance per selection row in the UI, carrying the relevant Mechanical object and its display settings.
- **Collectors** (`collect_views`, `collect_section_planes`, `collect_bodies`, `collect_boundary_conditions[_multi]`, `collect_bolt_pretensions[_multi]`, `collect_contact_tool_results[_multi]`, `collect_bolt_tool_results[_multi]`, `collect_all_results[_multi]`, `collect_solution_information_trackers[_multi]`, `collect_analyses`...): query `ExtAPI.DataModel` to populate the UI's selection lists. The `_multi` variants compile objects from **all** analyses in the project (multi-analysis support) as `(object, analysis)` tuples.
- **"Selection-aware" builders** (`build_bc_slides`, `build_bp_slides`, `build_result_slides`, `build_geometry_part_slides`, `build_mesh_part_slides`, `build_contact_summary_slide`, `build_solution_info_slides`, `build_analysis_context_slides`, `build_mesh_slide`): equivalents of `04_slides.py` but limited to the list of checked objects, with view/section/steps/legend/scale applied per row.
- **Geometry per isolated part**: `isolate_body_by_transparency` makes one part opaque and the others semi-transparent (context visible in the background) — one slide per part.
- **Mesh per isolated part**: `show_only_body` fully hides the other parts; up to 4 parts grouped on a single slide (`LAYOUT_MESH_MULTI` layout), beyond that a new slide starts automatically.

## 6. WPF interface

`AnsysReportGenerator_WPF.py` defines the **`ReportGeneratorApp`** class, which loads `AnsysReportGenerator_WPF.xaml` via `XamlReader` and drives a utility toolbar (above the tabs) and 6 tabs (vertical tabs, on the left side of the window).

**Utility toolbar** — 4 global actions, independent of the current selection:
- **Delete figures** — cleans up leftover `Figure` objects from a previous generation (`remove_stale_figures`)
- **Reset legends** — resets the viewport's legend to automatic (`reset_legend`)
- **Create basic views** — creates 7 views (X+/X-/Y+/Y-/Z+/Z-/ISO) in the View Manager, reusable afterward in the "..." side panel (`create_basic_views`)
- **Export to 3D (.avz)** — for each analysis in the project, exports an interactive `.avz` 3D view of every simple result and every child of a Contact Tool / Bolt Tool under the *Solution* branch (`export_all_3d_views`, see §8), into `data/export_3D/`

| Tab | Content |
|---|---|
| **General slides** | "Overview slides": two distinct Geometry / Mesh cards (checkbox + status + view-selection "Settings" button), "Parts to isolate (geometry)", "Mesh part to isolate", "Analysis context" (one row per project analysis, with view selection) |
| **Conditions and contacts** | Boundary Conditions, Bolt Pretension, Contacts to display, Connection: Contact Tool (*Connections* branch, no step), Solution Information |
| **Result categories** | Contact Tool Results (*Solution* branch, with steps), Results, Bolt Tool |
| **Combined slide** | Building a "different results" combined slide — see below |
| **Report preview** | One card per checked category (or added combined slide), reorderable by drag and drop — the chosen order is the report generation order |
| **Files** | Editable paths (template, images, CSV, legends, reports), data folder cleanup (see below), list of already-generated CSVs (Open/Show in folder), progress + access to the last generated report |

**"Combined slide (different results)" tab.** This flow used to live in 3 successive modal dialog boxes (template choice, then grid, then result choice); it is now fully integrated into this tab, with no separate window at all. At the top: choosing a multi-image template (2/3/4/6/8 results, same `MULTI_STEP_SLIDE_TEMPLATES` as the multi-step combined slides) and an "Add to report" button. On the left: a 2×4 grid where only the first N cells of the chosen template are active; clicking an empty cell shows, on the right, the (filterable) list of available results, clicking a result switches the right-hand panel to its full graphic configuration (same fields as a normal row — view/section/legend/appearance/scoping/scale factor, via `_build_row_config_fields`/`_apply_row_config_fields` — but with no notion of step, a different, fixed result per cell); the "Apply" button confirms the cell. "Add to report" requires all active cells to be configured, then adds the configuration (`MultiResultSlideConfig`) to `self._multi_result_slides` and resets the grid to build another one — nothing is generated immediately: like the other categories, the slide appears as a card in the "Preview" tab (dedicated "Delete" button, no checkbox) and is only built when clicking "Generate report" (`ReportGeneratorApp._build_multi_result_tab` and the `_on_multi_result_*`/`_show_multi_result_*` methods, `build_multi_result_slide`/`capture_multi_result_cell_image` in `05_interactive_slides.py`).

**Global side configuration panel ("...").** Every selection row (BC, result, isolated part, Solution Information tracker...) has a **"..."** button that no longer opens a separate window: it displays, to the right of the main window, a "SETTINGS" panel shared by all tabs (`borderConfigPanel`/`panelConfigPanel` in the XAML, an `Auto` column next to the `TabControl` — width 0 and `Visibility="Collapsed"` while no row is being configured). The panel's content depends on the "kind" of the clicked row (`ReportGeneratorApp._open_config_panel`):
- `"result"` — view / section / legend (file + orientation) / color display mode (Contour View) / scoping display (ScopingDisplay) / deformation scale (manual, or Auto Scale x1/x2) / step selection (BC, Bolt Pretension, Contact Tool, Bolt Tool, Results) — via `_build_row_config_fields` + `_build_steps_section_fields`
- `"geometry_part"` — view / section / context opacity (isolated part in geometry) — via `_build_geometry_part_fields`
- `"mesh_part"` — view only (isolated part in mesh, but also Geometry/Mesh/Analysis Context below) — via `_build_mesh_part_fields`
- `"solution_info"` — title / axes / color of the rebuilt chart (Solution Information trackers) — via `_build_solution_info_fields`

Each field category is a pair of shared functions `_build_*_fields`/`_apply_*_fields` that set/read their controls on a generic `target` (`_ConfigFieldsHolder`, a simple attribute bag) rather than on `self` of a dedicated window class — this decoupling is what lets the same code serve both the global side panel and the cell panel of the "Combined slide" tab (`_build_row_config_fields` alone, without steps). "Apply" confirms (`row_config.configured = True`) and closes the panel; "Cancel"/the "x" button close without confirming. `SectionRow.panel_kind` (formerly `dialog_factory`) carries the "kind" to use for each section (`None` for "Contacts to display", which has nothing to configure).

**View selection for Geometry / Mesh / Analysis Context.** These three checkboxes used to be captured only with the viewport's current view, with no configuration possible. They now also have a selectable (View Manager) view: for Geometry and Mesh (standalone checkboxes, no list), a "..." button appears next to each checkbox and opens the global side panel in `"mesh_part"` mode on a `MeshPartRowConfig` created for the occasion (`self._geometry_view_config`/`self._mesh_view_config` in `ReportGeneratorApp.__init__`, carrying `ExtAPI.DataModel.Project.Model.Geometry`/`.Mesh` respectively). For Analysis Context (one row per analysis), `AnalysisContextRowConfig` now carries a `view_name`, with its own "..." button (same `"mesh_part"`). In all three cases, the chosen view is applied right before capture (`apply_view_if_exists`), with no reset afterward (same convention as for BC/results: the applied view stays active for the next capture until explicitly changed).

**Result appearance (Contour View / legend / scoping / deformation scale).** The global side panel in `"result"` mode (BC, Bolt Pretension, Contact Tool, Bolt Tool, Results) exposes four settings per row: the result color display mode (`ResultPreference.ContourView`, among `ContourBands`, `Isolines`, `SmoothContours`, `SolidFill` — the .NET names are kept as-is in the UI, more explicit than a translation), legend orientation (`GlobalLegendSettings.LegendOrientation`, Vertical or Horizontal), scoping display mode (`ResultPreference.ScopingDisplay`, among `ScopedBodies` (default), `ResultOnly`, `AllBodies`), and deformation scale (manual via a numeric factor, or one of the two native "Auto Scale x1"/"Auto Scale x2" presets — `ResultPreference.DeformationScaling`/`DeformationScaleMultiplier`). By default, an unconfigured result is captured with `ContourBands`, vertical legend, `ScopedBodies` scoping, manual x1 scale. These settings are carried by `SlideRowConfig` (`contour_view`, `legend_orientation`, `scoping_display`, `deformation_scale_mode`/`scale_factor`, see `05_interactive_slides.py`) and applied only while capturing the relevant row (`apply_contour_view`/`apply_legend_orientation`/`apply_scoping_display`/`apply_scale_factor`), then systematically reset to their default value right after (`reset_contour_view`/`reset_legend_orientation`/`reset_scoping_display`/`reset_scale_factor`), so that a setting chosen for one row never "leaks" onto the next one. The `apply_*` functions call `ExtAPI.Graphics.Redraw()` right after changing their property: changing a display property via script doesn't refresh the viewport on its own, and without this explicit call the image exported right after would keep reflecting the old state. (The "show/hide legend" option — `ViewOptions.ShowLegend` — has been removed: it didn't produce the expected effect on export.)

Two settings, on the other hand, apply globally, to every image export without exception: `ExtAPI.Graphics.ViewOptions.ModelColoring = ModelColoring.ByMaterial` (already in place via `set_material_display`, `02_image_export.py`, applied before every geometry/mesh export), and `ExtAPI.Graphics.ViewOptions.ShowLogo = False` (forced in `export_current_view_image`, the common entry point for every image export, so that no report image shows the Ansys logo).

**Camera framing: the user's responsibility.** The image-export functions (`export_geometry_image`, `export_mesh_image`, `export_analysis_overview_image`, `export_object_image` in `02_image_export.py`, as well as `export_geometry_part_image`/`export_body_mesh_image` in `05_interactive_slides.py`) no longer call `ExtAPI.Graphics.Camera.SetFit()` before capture: a `SetFit()` right before export would silently overwrite any view chosen by the user (via the "View (View Manager)" field of the global side panel, or simply the current camera position). It is therefore up to the user to frame the view (manually or via a named view) before generating the report. Only `create_basic_views()` (the "Create basic views" button) still uses `SetFit()`, since its purpose is precisely to define the framing of the 7 standard views, not to apply an existing one.

Generation (`ReportGeneratorApp._on_generate`) walks through the order of the "Report preview" tab, opens a single `PPTReportBuilder` session, calls the `build_..._slides` function matching each category, updates the progress bar (with `SWF.Application.DoEvents()` to keep the window responsive), then cleanly closes the PowerPoint session and enables the report's "Open"/"Show in folder" buttons.

**Files tab: 4-quadrant layout.** Top-left: file paths, compact (font size 11, one line per path) — the Template is highlighted in light red (`DangerLightBrush`, a "sensitive" path, the whole generation depends on it), the Legends line in gray with a "to check" note (see below). Top-right: data folder cleanup. Bottom-left: CSV list. Bottom-right: progress + access to the last report. The "Generate report"/"Close" buttons stay at the window level (always visible, across every tab), unchanged.

**Legends folder: moved out of `DATA_ROOT`.** `LEGEND_FOLDER` (`00_constants.py`) now points to `<project>/user_files/legend` (a sibling of `PROJECT_DIR`, computed via `PROJECT_ROOT = os.path.dirname(PROJECT_DIR)`) instead of `data/legend`. This folder is no longer created automatically by the script (removed from the startup `ensure_folder_exists()` calls) — it is maintained manually by the engineer in the Ansys project's files, this script only *reads* it. A console warning flags its absence at load time, just like for the template. Direct consequence: no longer being inside `DATA_ROOT`, it no longer appears at all in the cleanup tiles (see below) — no need for an explicit exclusion anymore.

**Data folder cleanup (top-right quadrant).** One tile per `DATA_ROOT` subfolder (`list_data_cleanup_folders`, `00_constants.py`), dynamic, laid out in a 2×2 `UniformGrid` (`panelDataCleanup`). Each tile shows the total size and file count (`get_folder_stats`/`format_folder_size`) and a "Clear" button (light red, `DangerButtonLight`) that deletes all the folder's contents without deleting the folder itself (`clear_folder_contents`). A global "Delete all" button (bright red, `DangerButtonStrong`, full width below the grid) clears all these folders at once. Both actions require confirmation (`MessageBoxButton.YesNo`) before this irreversible deletion. If the reports folder is cleared, the "report result" tile goes back to its neutral state (`_reset_report_status_tile`). Refreshed when the app opens and continuously during generation (`_update_generation_progress`), like the CSV list.

**CSV and report lists: "Show in folder" rather than a download.** The CSV tile grid (`panelCsvFiles`) has become a tabular list (one row per file, name on the left + buttons on the right). The download buttons (copy to another location, `SaveFileDialog`) have been removed everywhere — CSV and PPTX report alike — in favor of a **"Show in folder"** button (`_on_show_in_folder`, `Process.Start("explorer.exe", "/select,\"<path>\"")`), which opens Windows Explorer with the file already selected. The viewing button is called "Open" (same behavior, `Process.Start(path)`).

> `AnsysReportGenerator_WPF.xaml` only contains declarative layout (styles, brushes, `x:Name`-named controls): check that file directly for the exact appearance, or the `.py` file (`_find_controls`) for the mapping between each `x:Name` and its use.

## 7. Ansys domain concepts used in the code

| Term | Meaning | Where in the code |
|---|---|---|
| **Step / Load Case** | A loading step in an analysis (e.g. Step 1 = preload, Step 2 = service load) | `get_step_count`, `selected_steps`, `evaluate_result_for_step` |
| **Boundary Condition (BC)** | An imposed constraint/load (fixed support, pressure, force...) | `collect_boundary_conditions[_multi]`, `build_bc_slides` |
| **Bolt Pretension** | Bolt preload | `collect_bolt_pretensions[_multi]`, `build_bp_slides` |
| **Contact Tool** | Analysis of a contact's quality (gap, pressure, slip...) — exists twice: one under *Connections* (definition, no step) and one under *Solution* (results, with steps), distinguished by their position in the tree (`_is_descendant_of`) | `collect_contact_tool_results` vs `collect_connection_contact_tool_results` |
| **Bolt Tool** | Forces in bolted connections (axial, shear...) | `collect_bolt_tool_results[_multi]` |
| **Solution Information** | Solver convergence data; its children ("trackers") only have a 2D chart, no 3D view | `collect_solution_information_trackers`, `export_chart_image_from_csv` |
| **Named View** | A camera view saved in the View Manager | `collect_views`, `apply_view_if_exists` |
| **Section Plane** | A cutting plane to reveal the model's interior | `collect_section_planes`, `apply_section_plane` |
| **Focus** | An aggregated result filtered by a selection (not yet integrated into the active UI) | — |

## 8. Ansys Mechanical APIs used

Access to the model and its objects:
```python
ExtAPI.DataModel.Project.Model            # model root
ExtAPI.DataModel.Project.Model.Analyses   # list of analyses
ExtAPI.DataModel.AnalysisList             # same, shortcut
ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.XXX)  # search by category
ExtAPI.DataModel.Project.Model.GetChildren(DataModelObjectCategory.Body, True)  # every body
ExtAPI.DataModel.Project.ProjectDirectory # "<ProjectName>_files" folder of the current project (used to locate PROJECT_DIR, see §2)
```

Display / image capture:
```python
ExtAPI.Graphics.Camera.SetFit()
ExtAPI.Graphics.ExportImage(path, GraphicsImageExportFormat.PNG, settings)  # settings = Ansys.Mechanical.Graphics.GraphicsImageExportSettings()
ExtAPI.Graphics.ViewOptions.ModelColoring / ShowMesh / ShowLogo / ResultPreference.DeformationScaleMultiplier / ResultPreference.ContourView / ResultPreference.ScopingDisplay (MechanicalEnums.Graphics.ScopingDisplay)
ExtAPI.Graphics.GlobalLegendSettings.LegendOrientation  # LegendOrientationType.Vertical / .Horizontal
ExtAPI.Graphics.ModelViewManager          # named views (ExportModelViews to XML to list them, ApplyModelView to activate one)
ExtAPI.Graphics.ModelViewManager.Capture3DImage(path)  # export of an interactive 3D .avz view of the active object ("Export to 3D" button)
ExtAPI.Graphics.SectionPlanes
ExtAPI.Graphics.ImportLegend(path, unit) / Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings()
ExtAPI.Graphics.Redraw()
```

Individual objects: most expose `.Activate()`, `.Name`, `.Children`, `.Parent`, `.DataModelObjectCategory`. To capture an image reliably, the code prefers a `Figure` snapshot (`obj.AddFigure()` then `figure.Activate()`) over a direct "live" capture.

Other notable APIs:
```python
Ansys.ACT.Mechanical.Transaction   # used as "with Transaction(True): ..." to defer UI refresh during bulk operations (deleting figures, looping over every body...)
materials.GetMaterialPropertyByName(material, group)   # Ansys module to read material properties
```

.NET / COM side (outside the Ansys API):
```python
clr.AddReference("Microsoft.Office.Interop.PowerPoint")  # + "Office"
clr.AddReference("System.Windows.Forms") / "System.Drawing"
clr.AddReference("PresentationFramework") / "PresentationCore" / "WindowsBase"  # WPF
```

## 9. How the code drives PowerPoint (COM Interop)

The project doesn't depend on any Python library to manipulate PowerPoint: `python-pptx` (like `pandas` or `openpyxl`) is incompatible with IronPython 2.7, the Python engine embedded in Ansys Mechanical, and is therefore never used here. It drives the PowerPoint application installed on the machine directly via **COM Interop**: Microsoft Office exposes a COM API, and .NET provides "Interop" assemblies (`Microsoft.Office.Interop.PowerPoint`, `Office`) that translate this COM API into .NET classes usable from any .NET language — hence from IronPython, which itself runs on the .NET CLR. This is what `03_ppt_utils.py` does right at the top of the file:

```python
clr.AddReference("Microsoft.Office.Interop.PowerPoint")
clr.AddReference("Office")
import Microsoft.Office.Interop.PowerPoint as PPT
import Microsoft.Office.Core as Office
```

`clr.AddReference` loads the matching .NET assembly (it's installed with Office, independently of the project), after which `PPT` and `Office` are used like ordinary Python modules — except that every object handled (`Presentation`, `Slide`, `Shape`...) is actually a remote COM object: every property access or method call actually goes and queries the running PowerPoint process, which has a cost (hence several optimizations described further below).

All the logic is concentrated in the `PPTReportBuilder` class, which owns a single PowerPoint session for the entire report generation (one open/close, not one per slide). Its constructor illustrates the module's central principle:

```python
def __init__(self, template_path):
    self.working_copy_path = get_unique_file_path(
        REPORT_OUTPUT_FOLDER, _build_working_copy_base_name(), ".pptx")
    shutil.copyfile(template_path, self.working_copy_path)

    self.app = PPT.ApplicationClass()
    self.app.Visible = True
    self.presentation = self.app.Presentations.Open(self.working_copy_path, WithWindow=True)
```

`PPT.ApplicationClass()` starts (or retrieves) an instance of the PowerPoint application itself, exactly as if the user had double-clicked its icon; `self.app.Presentations.Open(...)` then opens a file in it, returning a `Presentation` object which every subsequent operation acts on. The original template is never opened directly: a copy (`working_copy_path`) is created right before via `shutil.copyfile`, and it's this copy that gets opened — an accidental `Ctrl+S` in the PowerPoint window during generation therefore overwrites the copy, never the corporate template. `self.app.Visible = True` is not cosmetic: a session left invisible turned out to be unstable on a report with many slides (the `SlideMaster` object would eventually become inaccessible mid-generation), so the PowerPoint window stays visible throughout generation and closes normally at the end, in `close()`.

Adding a slide always consists of requesting a custom layout from the template by its index, then filling in the zones (`Shapes`) of the newly created slide:

```python
def _add_slide(self, layout_index):
    layout = self.presentation.SlideMaster.CustomLayouts[layout_index]
    return self.presentation.Slides.AddSlide(self.presentation.Slides.Count + 1, layout)
```

`SlideMaster.CustomLayouts` is the list of custom layouts defined in the template (visible in PowerPoint via View > Slide Master); their index (`LAYOUT_IMAGE_TABLE = 10`, etc., in `00_constants.py`) is determined once and for all by listing the template's layouts (see §12) and doesn't change again as long as the template isn't modified. `add_image_table_slide` then illustrates how a slide's zone is filled once it's been created:

```python
slide.Shapes[8].TextFrame.TextRange.Text = comment
slide.Shapes[2].TextFrame.TextRange.Text = title
...
coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[3]
slide.Shapes.AddPicture(img_path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                         coord.Left, coord.Top, coord.Width, coord.Height)
```

Each `Shapes[n]` corresponds to a precise zone defined in the layout at the time it was created in PowerPoint (a title zone, an image zone, a table zone...); the order and index of these zones are fixed by the template, not by the code, hence the importance of never rearranging the zones of an existing layout without updating the indexes used in `03_ppt_utils.py` (see §12). Text is always assigned on the newly created slide, never on the layout itself: modifying the layout would modify the master template for every future slide. To position the image, the code looks up the coordinates (`Left`, `Top`, `Width`, `Height`) of the image zone as defined *in the layout*, rather than hard-coding these coordinates: the image's position and size therefore stay consistent with what was drawn in the template, even if it evolves.

`add_csv_table` is the most performance-sensitive part, since every statement in the following block is a COM round-trip:

```python
for r in range(1, rows + 1):
    row_cells = table.Rows(r).Cells
    for border_index in range(1, 5):
        row_cells.Borders(border_index).ForeColor.RGB = 0x000000
        row_cells.Borders(border_index).Weight = 1
```

Borders are applied once per whole row (`table.Rows(r).Cells` accepts a range of cells) rather than cell by cell × side by side, which divided a table's formatting time by roughly the number of columns (up to 45 seconds for 8 rows before this optimization, versus a fraction of a second after). Text and font, on the other hand, have no "per-range" equivalent in PowerPoint's COM API and therefore necessarily remain applied cell by cell in the following loop. One last quirk: after filling the table, the code forces `table.Rows(r).Height = 1` on every row — a deliberately absurd value, but PowerPoint automatically brings it back to the minimum height actually needed to fit the text, which is the only way to tighten a table already created (`AddTable` allocates a much larger height than needed for size-7 text by default, which would overflow the slide without this fix).

Finally, `close()` illustrates the rule to systematically follow with COM objects: release them explicitly rather than counting on Python's garbage collector, so as to never leave an invisible PowerPoint process running in the background after an error:

```python
def close(self):
    self.presentation.Save()
    self.presentation.Close()
    self.app.Quit()
```

## 10. Python design choices used in the project

Several recurring choices in the code answer constraints specific to IronPython 2.7 and to execution in the Mechanical scripting console; understanding them helps in reading (and extending) any module of the project.

**Loading via `execfile()` rather than `import`.** `AnsysReportGenerator_WPF.py` does not do `import constants` or `from data_export import ...`: it calls `execfile(module_path)` for each of the six modules, in order. `execfile()` executes a file's contents as if it had been typed directly next in the same console, in the same global namespace — unlike `import`, which would create a separate namespace (`data_export.export_active_tabular_data` instead of `export_active_tabular_data`). It's this deliberate sharing of a single global namespace that lets `05_interactive_slides.py` call `export_active_tabular_data` (defined in `01_data_export.py`) directly with no prefix, exactly as the Mechanical scripting console itself does with `ExtAPI`/`DataModel`. It's also what lets a function defined earlier reference a name defined later in another module: `00_constants.py` uses `PROJECT_DIR`, which is actually only defined in `AnsysReportGenerator_WPF.py`, *before* the `execfile()` of `00_constants.py` — the loading order (`00` → `05`, then the main script last in the console) is therefore significant and must never be changed.

**Accessing .NET enums via `getattr()` rather than an explicit import.** Several places in the code, for instance `apply_contour_view` in `05_interactive_slides.py`, write:

```python
vo.ResultPreference.ContourView = getattr(vo.ResultPreference.ContourView, contour_view)
```

instead of importing the matching `.NET` enumeration and writing a long `if/elif` chain to convert the string chosen in the UI (`"ContourBands"`, `"Isolines"`...) into an enum value. `getattr(obj, "MemberName")` looks up the attribute named `"MemberName"` on the *type* of `obj` (here the type of the `ContourView` enum already present on the current instance): since the strings used in the UI's dropdowns (`CONTOUR_VIEW_OPTIONS`) carry exactly the same names as the .NET enum's members, `getattr` performs the string → enum-value conversion directly in one line, without having to explicitly import each enum type or keep it up to date if Ansys adds a member in a future version.

**Local failure, never an exception bubbling up to the UI.** Almost every export or setting-application function follows the same pattern:

```python
try:
    ...
except Exception as e:
    print "Error: " + str(e)
    return False  # or None
```

This choice is deliberate: a report generation can span dozens of slides, and a single misconfigured Boundary Condition row (or an image that fails to export) must not interrupt the entire generation and lose the work already done on previous slides. The error is therefore absorbed locally, logged to the scripting console (visible to the engineer), and the function returns a "neutral" value (`False`, `None`, or simply does nothing) that the caller can test to decide whether to continue.

**"collect_" functions always return a plain Python list.** Whether the source is `ExtAPI.DataModel.GetObjectsByType(...)`, a recursive tree walk, or the compilation of several analyses (`_multi` variants), every collector returns an ordinary Python `list`, never the original .NET/COM object. This fully decouples the WPF interface (which builds its dropdowns and checklists from these lists) from the details of how each object category is found in the Mechanical tree — a new collector can entirely change its internal logic without the UI code consuming it having to change.

**Constants and `(label, value)` options.** The options meant to appear in the UI (`CONTOUR_VIEW_OPTIONS`, `LEGEND_ORIENTATION_OPTIONS`, `DEFORMATION_SCALE_MODE_OPTIONS`, `BASIC_VIEW_ORIENTATIONS`...) are systematically lists of tuples `(label shown in the UI, technical value used in the code/API)`, with symmetrical `xxx_label`/`xxx_from_label` functions to convert one way or the other. This cleanly separates what is shown to the engineer (safely editable) from what must stay identical to the exact name expected by the .NET API.

## 11. Python fundamentals illustrated by the project's code

This section revisits the basic building blocks of the Python language (IronPython 2.7-compatible) using real examples from the project, for a reader discovering Python through this code.

**Variables and types.** A variable has no declared type, it takes the type of whatever is assigned to it: `DATA_ROOT = os.path.join(PROJECT_DIR, "data")` creates a `DATA_ROOT` variable of type `str` (a string). `MAX_TABLE_ROWS = 50` creates an integer. A list is written between square brackets and can grow dynamically: `_MODULE_FILES = ["00_constants.py", "01_data_export.py", ...]`. A dictionary maps keys to values between curly braces; the project makes little direct use of one in the code shown here, but `_DEFAULT_FILE_PATHS = dict((name, globals()[name]) for name, _, _, _ in FILE_PATH_SETTINGS)` builds one on the fly from a list of tuples.

**Functions.** `def` defines a function, its parameters go between parentheses, and `return` gives back its output value (a function with no `return` implicitly returns `None`):

```python
def safe_file_name(name):
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "object"
```

A parameter can have a default value, used if the caller doesn't supply one: `def add_image_table_slide(self, title, subtitle, img_path=None, csv_path=None, comment=" ")` can therefore be called with only `title` and `subtitle`, `img_path` then being `None`. The project systematically uses strings formatted with `.format()`, never Python 3.6+ f-strings (unavailable in IronPython 2.7): `"result_{}.csv".format(step_id)` rather than `f"result_{step_id}.csv"`.

**Classes.** `class ClassName(object):` defines a class (the explicit `(object)` is required in Python 2 to get a "new-style" class, with every modern object-oriented feature). `__init__` is the constructor, called automatically when an instance is created; `self` (the first parameter of every method) refers to the instance itself and must be used explicitly to read or write an attribute:

```python
class PPTReportBuilder(object):
    def __init__(self, template_path):
        self.working_copy_path = get_unique_file_path(...)
        self.app = PPT.ApplicationClass()

    def save_as(self, output_path):
        self.presentation.SaveAs(output_path)
```

`self.app` and `self.working_copy_path` are attributes belonging to each `PPTReportBuilder` instance: two instances created separately would each have their own PowerPoint session, without interference. A method is then called on an instance: `builder = PPTReportBuilder(TEMPLATE_PATH)` then `builder.save_as(output_path)`.

**Loops and conditions.** `for` iterates over any sequence (list, range of numbers, result of an API query); `range(1, rows + 1)` produces the integers from 1 to `rows` inclusive (Python always excludes `range`'s upper bound). `if`/`elif`/`else` tests conditions; indentation (always 4 spaces in this project, never mixed tabs) delimits blocks, there are no curly braces in Python:

```python
for step_id in steps:
    if step_id in selected_steps:
        rows.append(evaluate_result_for_step(obj, step_id))
    else:
        print "Step skipped: {}".format(step_id)
```

**Error handling (`try`/`except`).** A `try` block runs potentially risky code (COM access, disk access, Ansys API call); if an exception is raised, execution jumps directly to the matching `except` block rather than crashing the whole script:

```python
try:
    graphics.ExportImage(output_path, export_settings)
    return True
except Exception as e:
    print "Error exporting view: {}".format(e)
    return False
```

`except Exception as e` catches any standard error and makes it available in the `e` variable (usually converted to text via `str(e)` to display it). This pattern is everywhere in the project (see §10).

**Context managers (`with`).** `with open(path, "rb") as f:` opens a file and guarantees it's closed automatically on exiting the block, even if an error occurs inside — a safer, shorter equivalent of a manual `try`/`finally` with `f.close()`. Used for every CSV file access in the project, and repurposed for a different use with `with Transaction(True): ...` (`Ansys.ACT.Mechanical.Transaction`), which doesn't manage a file but defers Mechanical's UI refresh until the block exits, to speed up bulk operations (deleting several figures, looping over every body...).

**List comprehensions.** A comprehension builds a new list in a single expression, more concise than a classic `for` loop with `append`: `[cell.decode("utf-8") for cell in row]` (in `add_csv_table`) reads back each cell of a CSV row and decodes it from UTF-8, directly producing the decoded list.

**`import` vs `execfile`.** The project uses `import` for standard modules (`import csv`, `import os`) and .NET assemblies (`import Microsoft.Office.Interop.PowerPoint as PPT`, after `clr.AddReference`), but `execfile()` to load its own `00` through `05` modules — see §10 for the explanation of this unusual choice, specific to the execution context in the Mechanical console.

## 12. Creating a new custom slide in the Master Template

Adding a new slide type to the report first requires creating the matching layout in the PowerPoint template itself, and only then writing the Python code that fills it in. The PowerPoint-side procedure is strict on one point: **the new slide/layout must always be inserted at the end of the slide master, never in the middle**. Every index used in the code (`LAYOUT_IMAGE_TABLE = 10`, `LAYOUT_TABLE_ONLY = 8`, `LAYOUT_MESH_MULTI = 11`, in `00_constants.py`) corresponds to the layout's position in the template's `CustomLayouts` list; inserting a new layout in the middle of that list shifts the index of every existing layout after it, and silently breaks every slide already generated by the current code.

To create the layout, open the Master Template in PowerPoint (View > Slide Master), insert a new layout after the existing ones, and build its content either by drawing new zones (text box, image zone, table), or by copying elements from an existing layout close to what's needed. Once the layout is finished, save this new version of the template under a **different name** from the original (for instance by adding a suffix), to keep a backup copy of the template currently used in production in case the change turns out to be incompatible with the existing code.

The index of the new layout must then be identified, along with the index of each of its zones (`Shapes`), since it's by these indexes that the Python code refers to them (see `Shapes[n]` in §9). This is done by running, in the Mechanical scripting console, a small script that opens the template via COM Interop exactly like `PPTReportBuilder` does, and lists the available layouts:

```python
import clr
import os
import System

clr.AddReference("Microsoft.Office.Interop.PowerPoint")
clr.AddReference("Office")
import Microsoft.Office.Interop.PowerPoint as PPT
import Microsoft.Office.Core as Office
from Microsoft.Office.Core import MsoTriState

app = PPT.ApplicationClass()
app.Visible = True
template_path = r"PATH_TO_THE_TEMPLATE.pptx"  # adjust as needed
presentation = app.Presentations.Open(template_path, WithWindow=True)
custom_layouts = presentation.SlideMaster.CustomLayouts

for design in presentation.Designs:
    for i in range(1, design.SlideMaster.CustomLayouts.Count + 1):
        layout = design.SlideMaster.CustomLayouts[i]
        print(i, layout.Name)
```

This first block prints the full list of existing layouts with their index and name (for example `(1, "Title Page")`, `(10, "Image + Table")`...): that's where the index assigned to the newly added layout can be spotted. Once that index is identified, select that layout then list its zones in the order PowerPoint knows them:

```python
slide = custom_layouts[10]  # replace with the new layout's index

index = 0
for shape in slide.Shapes:
    index += 1
    print shape.Name
```

This second block gives, for each zone of the layout, its name and its position in the `Shapes` collection (the first item listed corresponds to `Shapes[1]`): it's this mapping between position and the zone's visual role (title, image, table, comment...) that must then be carried over into the Python code, exactly as `LAYOUT_IMAGE_TABLE` is documented today as a comment in `00_constants.py` (`# title[2] / subtitle[4] / image[3] / table[1] / comment[8]`). A new `add_..._slide` function can then be added to `03_ppt_utils.py` following the model of `add_image_table_slide`, using the new layout's index and the zone indexes identified this way.

## 13. Known pitfalls / technical choices

- **IronPython 2.7 constraints**: the scripting engine embedded in Ansys Mechanical runs Python 2.7 via .NET, not Python 3. Any code change must therefore stay compatible with these restrictions:
  - `.format()` instead of f-strings: `"result_{}.csv".format(step_id)`, never `f"result_{step_id}.csv"` (syntax error in IronPython 2.7).
  - `print "text"` as a statement, never `print("text")` as a function.
  - `os.path.join(...)` instead of the `pathlib` module, absent from IronPython 2.7.
  - No type annotations (`variable: str = ""`), no `async`/`await`.
  - The `pandas`, `openpyxl`, and `python-pptx` libraries are incompatible and must never be imported — this is why every piece of tabular data in the project goes through plain CSV files (standard `csv` module) rather than these libraries.
- **PowerPoint session always visible** (`self.app.Visible = True` in `PPTReportBuilder.__init__`): a session left invisible turned out to be unstable on a report with many slides (`COMException` on `SlideMaster` mid-generation). The PowerPoint window closes normally at the end (`close()`).
- **The original template is never opened directly** — always a working copy (see §4), to never risk overwriting it via an accidental `Ctrl+S` during generation.
- **Table borders applied per whole row**, not cell by cell × side by side: every COM round-trip is expensive, this optimization divided a table's formatting time by roughly N (N = number of columns).
- **Legend unit always deduced dynamically** (`get_result_display_unit`, reads the text shown in `VisibleProperties`, not `result_obj.Maximum.Unit` which was found unreliable): `ImportLegend()` compares the requested unit to that of the object **currently active** in the viewport, hence a systematic explicit `Activate()` right before, to avoid a one-row offset with the object actually being processed.
- **CSV always read/written in explicit UTF-8** (`open(path, "rb")` + manual decoding): units returned by Mechanical sometimes contain special characters (degree, micro...) that crash a read/write without explicit encoding.
- **Table display limit** (`MAX_TABLE_ROWS` / `MAX_TABLE_COLUMNS`, 50×50 by default): beyond that, the CSV is still generated but not inserted as a PowerPoint table (unreadable once inserted).
- **Inserting a layout into the template**: always at the end of the slide master, never in the middle — see §12 for the full procedure and the reason (index shift of the `LAYOUT_*` constants used throughout the code).
