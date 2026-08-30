# Excel Worksheet Boosters

**Excel Worksheet Boosters** is a collection of VBA tools that enhance Microsoft Excel worksheets.

The goal of this project is to provide practical worksheet utilities that improve productivity and simplify repetitive tasks. Each toolkit is designed to work independently while sharing a common design philosophy.

---

# MAPBST – Map Booster for Excel

MAPBST is a worksheet-based tile map editor for Microsoft Excel.

Instead of manually entering bitmap numbers, simply paint your map using cell background colors.

MAPBST automatically converts colored cells into bitmap IDs, supports region painting, generates color palettes, restores bitmap numbers, and expands map layouts.

---

## Block Fill

The following example shows the first 45 × 45 map area.

On the left is a worksheet containing only colored cells.

After running **RunBlockFill**, bitmap numbers are automatically generated.

![Block Fill Demo](MAPBST/images/blockfill-before-after.png)

---

## Features

* Automatic 2×2 bitmap placement (Block Fill)
* Region painting based on cell colors (Domino Paint)
* RGB color viewer
* Automatic palette generation
* Bitmap number restoration
* Map layout expansion
* Configurable worksheet-based workflow

---

## Included Sample

The repository includes:

```text
MAPBST/MAPBST_Sample.xlsm
```

The sample workbook intentionally contains only colored cells without bitmap numbers.

This allows you to experience the complete MAPBST workflow from scratch.

---

## Quick Start

Open **MAPBST_Sample.xlsm** and enable VBA macros.

### 1. View RGB Colors

Run:

```text
ShowActiveCellColor
```

Select any colored cell to display its RGB value.

---

### 2. Generate Bitmap Numbers

Run:

```text
RunBlockFill
```

MAPBST searches for valid 2×2 blocks and automatically places bitmap numbers.

---

### 3. Paint Connected Areas

1. Enter a bitmap number into a colored cell.
2. Select that cell.
3. Run:

```text
RunDominoPaint
```

MAPBST fills all connected cells with the same background color.

---

### 4. Generate a Palette

Run:

```text
RunPalette
```

A Palette worksheet is created containing:

* Bitmap IDs
* Color samples
* RGB color values

---

## SHTBST – Sheet Store Booster

SHTBST is a simple worksheet data storage and restore tool.

It temporarily stores the values of a specified worksheet range and restores them to the original location when needed.

This can be useful when worksheet data needs to be modified temporarily, tested, or experimented with without losing the original values.

The application determines the range to be stored, while SHTBST handles the Store and Restore operations.

---

## Store and Restore

The basic workflow is:

```text
Worksheet Range
      ↓
    Store
      ↓
Modify / Test / Experiment
      ↓
   Restore
      ↓
Original Values
```

SHTBST stores the values of the entire specified range, including blank cells.

This makes it possible to restore the complete state of the stored area rather than restoring only cells that contained values.

---

## Included Sample

The repository includes:

```text
SHTBST/SHTBST_Sample.xlsm
```

The sample workbook provides a simple environment for testing the Store and Restore functions.

---

## Quick Start

Open **SHTBST_Sample.xlsm** and enable VBA macros.

Specify the Store ID in the sample worksheet and use the test routines to verify the Store and Restore operations.

The sample uses a worksheet area defined by the application:

```vb
Public Const AREA_INFO As String = "E5:W20"
```

The range itself is determined by the application. SHTBST does not require a fixed worksheet area.

---

## Project Structure

```text
Excel Worksheet Boosters
│
├── MAPBST
│   ├── MAPBST_Sample.xlsm
│   ├── images
│   │   └── blockfill-before-after.png
│   │
│   ├── MAPBST_00_Config.bas
│   ├── MAPBST_10_RGBViewer.bas
│   ├── MAPBST_40_BlockFill.bas
│   ├── MAPBST_50_DominoPaint.bas
│   ├── MAPBST_60_Palette.bas
│   ├── MAPBST_70_MapRestore.bas
│   ├── MAPBST_80_Reset.bas
│   └── MAPBST_90_Expand.bas
│
└── SHTBST
    ├── SHTBST_Sample.xlsm
    ├── SHTBST_00_Config.bas
    ├── SHTBST_10_Store.bas
    ├── SHTBST_20_Restore.bas
    └── SHTBST_90_Test.bas
```

---

## Design Philosophy

Traditional worksheet tools often require users to manually manage data and repeatedly perform the same operations.

**Excel Worksheet Boosters** follows a different approach.

The worksheet itself becomes part of the tool.

Each Booster provides a small set of focused operations that can be combined with the worksheet-based applications users already create.

MAPBST allows users to create maps by painting cells.

SHTBST allows users to temporarily store worksheet data and restore it when needed.

The goal is not to replace Excel, but to make the worksheet itself more useful as a development and working environment.

> **Turn the Excel worksheet into a useful tool.**

---

## Future Toolkits

MAPBST is the first toolkit included in the **Excel Worksheet Boosters** project.

SHTBST is the second toolkit.

Additional worksheet tools and utilities may be added in the future.

## License

This project is licensed under the MIT License.
See the LICENSE file for details.
