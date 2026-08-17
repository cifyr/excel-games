# Excel Games

Four playable games built entirely inside Microsoft Excel — no VBA, no macros, no add-ins. Just worksheet formulas, form controls, and conditional formatting on a dense cell grid used as a display. Open a file in Excel and play.

Because there are no macros, every file is a plain `.xlsx`. Nothing to enable, nothing to trust.

| File | Game | Size |
|---|---|---|
| [`ExcelCraft.xlsx`](ExcelCraft.xlsx) | Minecraft — first-person voxel sandbox with mining, placing, and a hotbar | 4.5 MB |
| [`raycaster_v10.xlsx`](raycaster_v10.xlsx) | First-person raycaster maze at 400x240 | 4.0 MB |
| [`Chess2.2.3.xlsx`](Chess2.2.3.xlsx) | Full chess with legal-move validation, castling, and a move log | 135 KB |
| [`TicTacToe.xlsx`](TicTacToe.xlsx) | Tic-tac-toe with win and draw detection | 42 KB |

## How these work

All four share the same three tricks.

**Iterative calculation is the game loop.** Every game keeps mutable state — player position, board contents, whose turn it is — in cells that refer to themselves, e.g. `PLAYER X = IF(moveIsLegal, candidateX, PLAYER X)`. Excel normally rejects that as a circular reference. With iterative calculation enabled it instead resolves it a fixed number of passes per recalc, which turns one recalculation into exactly one tick of game state. Every workbook here ships with the setting already saved on.

**Form controls are the input device.** Spin buttons, scroll bars, and checkboxes each write a number into one linked cell. A spinner click nudges an odometer by one; the formulas read the delta against the previous frame's value to know a click happened, then reset. Nothing listens for keystrokes — a click on a control is the only event Excel gives you without macros.

**The cell grid is the framebuffer.** Columns are narrowed to square pixels and conditional formatting colors each cell from the value the render formulas write into it. The raycaster's `3D View` sheet is a 400-column by 240-row image, one cell per pixel, recomputed on every recalc.

## Setup

Excel for Windows or Mac, desktop. Excel Online and Google Sheets will not run these — neither supports the form controls, and their circular-reference handling differs.

Iterative calculation is saved on inside each file, but if Excel warns about circular references, turn it on manually:

- **Windows** — File > Options > Formulas > Enable iterative calculation
- **Mac** — Excel > Settings > Calculation > Use iterative calculation

Set **Maximum Iterations to 1** for ExcelCraft, the raycaster, and chess. One iteration per recalc is what keeps a single click equal to a single game tick; higher values make the state advance several steps per click.

Leave calculation on Automatic. If a frame ever looks half-drawn or stale, press **F9** once to force a clean recalc.

## The games

### ExcelCraft.xlsx — Minecraft

A 12x12 floating island rendered in first person, with a height map, block types, textures, and a working mine/place loop.

Play on the **GAME** sheet. All controls sit next to the viewport:

| Control | Action |
|---|---|
| MOVE, spinner | Walk forward / back along your facing |
| TURN, scroll bar | Rotate the view, 3 degrees per click |
| PITCH, spinner | Look up and down |
| MINE | Swing at the block under the crosshair. Cracks accumulate over repeated clicks, then the block breaks and drops into your hotbar |
| PLACE | Put one of the selected block on the ground in front of you |
| SLOT | Choose hotbar slot 1-9 |
| REGEN | Rebuild the world and respawn |

Grass drops dirt. Bedrock is unbreakable. Reach is 4.3 blocks. You auto-climb one-block steps, and the void surrounds the island — walk off and that is that. A top-down minimap on the `Setup` sheet shows your position in white.

REGEN takes effect over the next two recalcs; click it, then press F9 or any arrow control.

Field of view is 72 degrees, move speed 0.22 blocks per click. Sheets: `GAME` is the viewport, `Setup` holds all engine state and the minimap, and `Gen` / `World` / `Hmap` / `Rays` / `Floor` / `Tex` hold world generation, block IDs, the height map, the per-column ray march, floor casting, and textures.

### raycaster_v10.xlsx — Maze

A 16x16 maze drawn in first person by DDA raycasting, at 400x240 with a 60-degree field of view.

Play on the **Setup** sheet, which holds the controls and the full engine readout; watch the **3D View** sheet for the image.

| Control | Action |
|---|---|
| Forward / Back, spinner | Move along your facing, 0.15 units per click |
| Turn, scroll bar | Rotate, wrapping continuously in either direction |
| Pitch, spinner | Look up and down |

The `Rays` sheet does the work: one column per screen column, marching a ray through the grid until it hits a wall, then computing hit distance, applying fisheye correction, and deriving the wall slice's top and bottom rows. Walls hit on their north-south face get a darker shade so corners read correctly. `Maze` holds the 16x16 layout — edit it and the view updates.

### Chess2.2.3.xlsx — Chess

Full chess against a second person at the same keyboard, with per-piece legal-move validation, castling, and a running log.

Every square carries a checkbox. **Click a square to select your piece, then click the destination.** Illegal moves are rejected and the selection clears.

The click detection is the interesting part. There is no macro to tell the sheet which square you clicked, so the `input handling` sheet keeps two 8x8 grids: the live checkbox states, and a snapshot of what they were on the previous recalc. XOR the two and exactly one cell differs — that is your click. `COUNTIF(...)=1` confirms a single unambiguous toggle before anything is committed, then the snapshot is copied forward for the next turn.

Sheets: `Sheet1` is the visible board, `base` is the starting position, `livegame` is the current position and move resolution, `piecetype` maps pieces to movement rules, `input handling` and `inputsimplified` decode clicks, and `Logs` records the game.

### TicTacToe.xlsx — Tic-tac-toe

The small one, and a good place to read the pattern before opening the others. A single sheet holds separate X and O boards, the turn counter, win detection across all eight lines, and a draw check labeled `Cats` for cat's game. Type into the board cells to play.

## Status

Personal experiments in how far pure worksheet formulas can be pushed. All four are playable. The two 3D workbooks are heavy by design — a recalc is a rendered frame, so expect a beat between click and picture on older machines.

## License

MIT — see [LICENSE](LICENSE).
