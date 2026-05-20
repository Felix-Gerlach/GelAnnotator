# Gel Annotator - Operator Manual
 
Author: Felix Gerlach  
Bug reports: felixgerlach@yahoo.de

---

## 1. Purpose

This application is for:
- loading one or multiple gel/blot images,
- optionally extracting multiple panels,
- composing panels into one figure,
- annotating lanes, headers, brackets, highlights, and marker ticks,
- running optional lane intensity / band-size analysis,
- exporting publication-ready output.

The workflow is step-based and stateful.

---

## 2. Quick Start

1. Open the app.
2. In **Step 1**, choose mode and image(s).
3. If you use multi-panel or multi-image workflows, complete **Step 2a/2b**.
4. Edit image in **Step 3**.
5. Set lanes and gel region in **Step 4**.
6. Optionally add highlights in **Step 4b**.
7. Add headers/brackets in **Step 5**.
8. Assign marker ticks in **Step 6**.
9. Optionally run analysis in **Step 6b**.
10. Finalize/export in **Step 7**.

---

## 3. Workflow Reference

## Step 1 - Session Settings

Main controls:
- Experiment type
- Panel mode (`single` or `multi_panel`)
- Source mode (`single_image` or `multi_image`)
- Default marker unit (`bp` / `kDa`)
- Invert preference
- Keep settings for next image
- UI theme color
- Image selector buttons
- `Readme / Help` button

## Step 2a - Panel Selection from Source Images (mode-dependent)

Used when selecting multiple panels before composition.

Functions:
- select source image,
- rotate source (`Apply`, `+/-90`, `+/-1`),
- draw panel rectangles,
- assign panel metadata:
  - marker panel flag,
  - scaling group ID.

Result:
- selected crops + metadata are passed to compose step.

## Step 2b - Compose Multiple Panels/Images (mode-dependent)

Functions:
- drag panels,
- rotate (`+/-90`, `+/-1`),
- scale X/Y,
- color preprocessing per panel (B/C/G, WB, invert, B/W),
- align/scaling using marker-based workflow,
- optional final crop with undo.

## Step 3 - Edit

Functions:
- brightness / contrast / gamma,
- WB / invert / B/W,
- rotation (`Apply`, `+/-90`, `+/-1`),
- crop and reset edits.

## Step 3b - Panel Selection from Single Edited Image (mode-dependent)

Functions:
- draw one or more crop rectangles from the edited image.

## Step 4 - Panel Layout

Functions:
- set lane count,
- marker lane index,
- define gel region(s) and lane counts per region,
- choose whether optional highlight step is enabled.

Display:
- lane boundaries + lane numbers in circles are shown over preview.

## Step 4b - Band Highlights (optional)

Enabled only when Step 4 checkbox is active.

Highlight shapes:
- `box`
- `arrow` (filled arrow head)
- `asterisk`

Controls:
- color,
- line width,
- apply style to existing highlights,
- remove last / clear all.

## Step 5 - Header / Column Annotation

Functions:
- add/remove header rows,
- per-row:
  - row name,
  - top/bottom position,
  - angle,
  - heading font size,
  - value font size,
- bracket groups (lane ranges + labels),
- bracket text distance,
- per-element offsets,
- top-annotation order editing.

## Step 6 - Marker / Ladder

Functions:
- choose marker definition (library/new/edit),
- assign tick positions by selecting a tick then clicking panel,
- override / clear / hide / show marker ticks,
- marker style controls:
  - X/Y offsets,
  - tick length,
  - label gap,
  - marker font size.

Option:
- `Run lane intensity/band analysis before review`.

## Step 6b - Lane Intensity and Band Analysis (optional)

Functions:
- lane profile/histogram view,
- lane overlay + marker fit view,
- peak detection controls:
  - threshold,
  - prominence,
  - min distance,
  - smoothing window,
  - polarity,
- optional background correction using selected empty lane,
- fit model select:
  - monotone interpolation,
  - log-linear regression,
  - log-quadratic regression,
- save analysis report.

## Step 7 - Review and Export

Functions:
- inspect final rendered image,
- interactive element editing on selected items:
  - position,
  - color,
  - font,
  - size,
  - angle/orientation,
  - line/tick/bracket styling,
- save final output

---

## 4. Important Defaults (current)

Selected core defaults in current version:
- panel lanes default: `10`
- marker label font size default: `15`
- bracket text distance default: `10`
- highlights default: disabled
- highlight style default: `box`
- analysis defaults:
  - threshold `20`
  - prominence `7`
  - min distance `10`
  - smoothing `10`
  - fit mode `Monotone interpolation`

When background correction is enabled in analysis, workflow-specific reduced defaults are applied (`threshold=3`, `smoothing=5`).

---
