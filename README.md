# Excel Games

Four playable games built entirely inside Microsoft Excel — no VBA, no macros, just spreadsheet formulas, conditional formatting, and cell grids used as a display. Open a file in Excel and play.

## Overview

Each workbook is a self-contained game where the spreadsheet itself is the engine and the screen. State lives in cells, logic is pure worksheet formulas, and the visible grid is the render target. Because there are no macros, the files are plain `.xlsx` and safe to open without enabling anything.

## The games

| File | Game | Notes |
|---|---|---|
| `ExcelCraft.xlsx` | Minecraft in Excel | A voxel/sandbox take on Minecraft rendered in the cell grid. Largest workbook (~4.5 MB). |
| `raycaster_v10.xlsx` | Raycaster maze | A first-person raycasting maze engine — walls drawn by casting rays across the grid, tenth iteration. ~4 MB. |
| `Chess2.2.3.xlsx` | Chess | Playable chess board with move logic in formulas. |
| `TicTacToe.xlsx` | Tic-tac-toe | The small one — classic tic-tac-toe. |

## How to play

Open any `.xlsx` in Microsoft Excel (desktop recommended for performance on the larger workbooks). Follow the in-sheet instructions; input is typically done by typing into designated cells or using the arrow-key/navigation scheme described in the workbook.

## Status

Personal experiments in how far pure Excel formulas can be pushed. Playable.
