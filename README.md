# AutoCAD Engineering Toolkit (AET)

A small, original collection of AutoCAD VBA utilities for civil and structural drafting. The modules are written for this repository and do not include LL Tools source, branding, UI, or implementation.

## Compatibility and assumptions

- Windows desktop AutoCAD with the VBA engine / Autodesk VBA Enabler installed and enabled. The code uses the AutoCAD ActiveX/VBA object model, not Civil 3D APIs.
- Intended for full AutoCAD 2018 and later. AutoCAD LT and AutoCAD for Mac are not supported. Autodesk has changed VBA availability across releases; check the VBA Enabler for your exact AutoCAD release.
- Work in a copy of the drawing first. All distances and coordinate values use the current drawing's units; the toolkit does not infer or convert units. Volume is in cubic drawing units.
- 2D setting-out, grid, and boundary tools operate on world-coordinate plan geometry; AutoCAD converts picked points from the active UCS. Keep a consistent UCS and drawing elevation. Verify engineering results independently before construction or issue.

## Import and run

1. Open a test drawing and save a backup.
2. In AutoCAD, run `VBALOAD` to load an existing DVB, or open `VBAIDE` and create/open a VBA project. This repository intentionally contains source modules, **not a compiled `.dvb`**.
3. In the Visual Basic Editor, choose **File > Import File…** and import every `.bas` file from `src/vba/`. Import all modules into the same project; no external references or class/form modules are needed.
4. In the editor, run **Debug > Compile**. If AutoCAD does not expose the VBA commands, install/enable the matching Autodesk VBA Enabler and restart AutoCAD.
5. Run a public macro from **Tools > Macro > Macros…** (or the VBA editor). `AET_Menu` displays the command list. Save the project as a `.dvb` locally if desired; that binary is not committed here.

The modules are also usable from an existing VBA project by importing the `.bas` files in `src/vba/` in the same session. Do not paste the `Attribute VB_Name` line into a code pane; it is module export metadata.

## Public macros

| Macro | Purpose |
| --- | --- |
| `AET_Menu` | Displays a simple in-drawing macro index. |
| `AET_RunChecks` | Creates and removes a temporary 4-by-3 closed polyline and checks area 12 and perimeter 14. |
| `AET_SetupLayers` | Creates or updates AET point, level, grid, note, and helper layers. |
| `AET_SetCurrentLayer` | Creates (if needed) and activates a named layer. |
| `AET_IsolateSelectedLayer` | Runs AutoCAD's layer-isolation command for a selected object's layer; restore with `LAYUNISO`. |
| `AET_ThawAllLayers` | Thaws all layers except reserved `0` and `Defpoints`. |
| `AET_GeometryReport` | Reports selected-object length/perimeter, supported area, and XY extents at the command line. |
| `AET_CalculateArea` | Reports area and perimeter/length of one supported closed entity. |
| `AET_NumberPoints` | Places numbered circle markers at user-picked 3D points. |
| `AET_LabelCoordinates` | Places X/Y/Z coordinate labels at user-picked points. |
| `AET_DrawGrid` | Draws axis-aligned gridlines between two picked corners at a chosen spacing. |
| `AET_ChainageOffset` | Reports chainage and signed left/right offset from a picked point to a line or 2D polyline. |
| `AET_LabelLevels` | Places spot-level text using each picked point's Z coordinate. |
| `AET_ExportPointsCsv` | Exports selected point entities and block insertion points. |
| `AET_ImportPointsCsv` | Imports point entities from the documented CSV format. |
| `AET_CalculateVolume` | Estimates cut/fill or stockpile volume using a closed straight-segment lightweight-polyline boundary and a grid approximation. |
| `AET_DrawConstructionLine` | Adds an infinite construction line through two picked points. |
| `AET_OffsetSelectedGeometry` | Creates a positive-distance offset of a selected line or 2D polyline. |
| `AET_DrawFoundationOutline` | Draws a rectangular foundation outline from a picked lower-left corner and dimensions. |
| `AET_DrawDrainageNote` | Adds plain-text drainage annotation. |
| `AET_CleanupDrawing` | Starts AutoCAD's native `PURGE` after warning that purging is not generally undoable. |
| `AET_AuditDrawing` | Offers to run AutoCAD `AUDIT` with repairs after confirmation. |

Layer names are `AET-POINTS`, `AET-LEVELS`, `AET-GRID`, `AET-NOTES`, and `AET-HELPERS`. Colors use AutoCAD indexed colors; linetypes are loaded from `acad.lin` where needed.

## CSV format

ANSI plain text with one header followed by `id,x,y,z` rows. Use commas as field separators and a period as the decimal mark. IDs are informational and are not used as AutoCAD handles. Coordinates are in drawing units. See [`samples/points.csv`](samples/points.csv).

Import makes AutoCAD point entities on `AET-POINTS`; export supports `POINT` objects and block-reference insertion points. Other selected entity types are skipped. Keep coordinate values numeric and do not include thousands separators or embedded commas.

## Volume method and limitations

`AET_CalculateVolume` requires a closed lightweight 2D polyline with 3–1,000 straight segments in world XY and a positive Z normal. It rejects bulged/curved edges and checks for proper nonadjacent edge crossings. A regular grid samples each cell at 4 × 4 midpoint locations to estimate plan area, then multiplies that area by the absolute difference between the two user-supplied average levels. The sign of the level difference labels the result cut or stockpile/fill. This is a constant-average-thickness estimate, not a triangulated survey/design surface comparison; it does not account for slopes, local level variation, overhangs, or compaction factors. Decrease grid spacing for finer boundary sampling, subject to the 500,000-sample limit. Review the reported sampled area and independently verify the result.

Other limitations: `AET_GeometryReport` only computes area for closed lightweight/2D polylines, circles, and regions; 3D polylines are not supported by chainage/offset; grid and foundation helpers create axis-aligned world-coordinate geometry; the drainage helper creates plain text rather than a structured pipe object. The toolkit does not automatically detect every invalid or self-intersecting shape.

## Manual verification

Use a disposable drawing with units set to metres and a consistent UCS:

1. Run `AET_SetupLayers`; confirm all five AET layers appear with the expected colors.
2. Create a 3-unit line and a closed 4-by-3 rectangle. `AET_GeometryReport` / `AET_CalculateArea` should report line length 3 and rectangle area 12 square drawing units, perimeter 14 drawing units. `AET_RunChecks` repeats the rectangle calculation and deletes its temporary test geometry.
3. Run `AET_NumberPoints` for two points and verify sequential labels and point markers. Pick a point at Z=12.345, run `AET_LabelLevels` at precision 2, and confirm `EL 12.35`.
4. Draw a straight alignment from (0,0) to (10,0), then use `AET_ChainageOffset` near (4,2). Expect chainage about 4 and offset about +2; the right side should be negative.
5. Draw a closed 10-by-10 boundary. For base level 0 and second level 2, `AET_CalculateVolume` should return approximately 200 cubic drawing units (grid boundary estimate can vary slightly).
6. Import `samples/points.csv`, export the created points, and compare coordinate rows. Test malformed CSV on a disposable drawing and confirm an explanatory error.
7. Try grid, construction-line, foundation, drainage-note, offset, layer isolation/restoration (`LAYUNISO`), and the audit/cleanup confirmation and decline paths. Confirm generated geometry lands on its documented layer and review the undo stack before retaining changes.

## Project layout

- `src/vba/` — importable VBA standard modules (`.bas`).
- `samples/` — example point CSV.
- `CHANGELOG.md` — release history.
- `CONTRIBUTING.md` — contribution and review guidance.

Automated VBA execution requires AutoCAD on Windows and is not available in this repository's current environment. Compile the imported project in the VBA editor and use the manual checks above for release verification.
