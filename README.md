# VFX Script Breakdown

Break a screenplay down into VFX shots — scope, methodology, shot counts, artist-days
and cost — entirely in your browser. One HTML file, no install, no server: the PDF is
parsed locally and your work autosaves to the browser.

## Quick start

1. Open **`index.html`** in Chrome (double-click, or serve the folder).
2. Drop a screenplay PDF onto the upload screen. The parser reads standard
   screenplay formatting (numbered shooting scripts included) and shows you the
   scene count/numbering before you commit — check it matches your script.
3. Break it down. Everything autosaves to that browser; use
   **Export → Session backup** to move work between machines (the backup carries
   the script with it).

## The three panes

| | |
|---|---|
| **Left — outliner** | every scene. Green dot = has VFX, yellow = TBD, tint depth = shot density, L/M/H difficulty badges. Clicking a scene jumps the script **and opens that scene's full shot list in the right pane**. |
| **Middle — script** | the screenplay. Click a line to tag it (VFX / TBD / Practical), click a slugline for scene-level work. Drag a script line into a shot description to link it. |
| **Right — inspector** | scope of work from the element library, shooting methodology, LOD, shot count, **est. days/shot**, cost, notes, storyboard drops. |

Keyboard: `V` vfx · `T` tbd · `P` practical · `1–9` shot count · `⏎` next untagged ·
`↑↓` move · `E` element search · `⌘Z` undo · `⌘S` export.

## Element library

A starter rate card ships built in (screens & comps, cleanup, blood & gore, fire,
extensions, character work, opticals…). Everything is editable in **Library** —
rename, reprice, add your own elements and categories; renames propagate to every
shot already tagged. **Reset to default library** brings the starter card back.

## Days

Every shot (line-tagged or scene-level) carries an **estimated artist-days** value.
Days roll up per scene, in the header, in Overview/Summary, and export as their own
column everywhere.

## Exports

Open **Export** (or `⌘S`). Set your **project code** and **sequence code** there —
they flow into the workbook and CSV (placeholder `PRJ` until you change them).

- **Breakdown JSON → xlsx** — the full bid workbook. Export the JSON, then:

  ```bash
  python3 pipeline/build_xlsx.py MY_FILM_breakdown.json
  ```

  Nothing to install — the spreadsheet library ships bundled in
  `pipeline/vendor/`; any stock Python 3 works (macOS includes one).

  Produces `<PROJ>_VFX_Breakdown.xlsx` from `pipeline/template.xlsx`: Shot
  Breakdown sheet with live formulas, Days column, storyboards embedded in
  column R, allowance row, linked Sequence Breakdown.
- **CSV** — the same columns for quick pasting.
- **Session backup** — script + all work in one JSON; re-import on any machine
  (also from the upload screen).

An **Include prices** toggle strips cost columns from every export when you need a
scope-only document.

## Notes

- Fully offline: the PDF engine ships in `lib/`, and the script never leaves
  your machine. (If `lib/` is deleted, the tool falls back to loading the engine
  from a CDN — the only case where it touches the network.)
- Scene numbering: the parser handles numbered shooting scripts, doubled margin
  numbers ("66  66"), and numberless cold-open beats. Always eyeball the scene
  list against your script before breaking down — the upload screen shows the
  first scenes for exactly this reason.
- Storage lives in the browser (localStorage). Export a session backup before
  clearing site data or switching browsers.
