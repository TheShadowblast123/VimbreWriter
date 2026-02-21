# Usage Guide

This document explains how to use VimbreWriter and the provided development tools.

## Using the Script

VimbreWriter attempts to bring modal editing to LibreOffice. When enabled, you start in **Normal Mode**.

### Modes
- **Normal Mode**: Navigate the document and manipulate text using commands.
- **Insert Mode**: Type text normally. Press `Esc` or `Ctrl-[` to return to Normal Mode.
- **Visual Mode** (`v`): Select text character-by-character.
- **Visual Line Mode** (`V`): Select text line-by-line.

### Common Workflows
- **Navigation**: Use `hjkl` for basic movement. Use `w` and `b` to jump words.
- **Editing**: 
    - `ciw`: Change inside word (deletes word and enters Insert Mode).
    - `dd`: Delete current line.
    - `yy`: Copy current line (uses system clipboard).
    - `p`: Paste after cursor.
- **Search**: 
    - Press `/` to focus the LibreOffice Find Bar.
    - Press `\` to open the Find & Replace dialog.

## Makefile Usage

The project includes a `Makefile` to automate building and testing.

### 1. Testing (`make testing`)
This is the primary command for development.
```bash
make testing
```
**What it does:**
1. Checks for `unopkg`.
2. **Kills** any running LibreOffice instances (`soffice.bin`).
3. Uninstalls the previous version of the extension.
4. Compiles `src/vimbrewriter.vbs` into an `.oxt` package.
5. Installs the new extension.
6. Opens `docs/testing.odt` in LibreOffice Writer.

### 2. Building (`make extension`)
Builds the distributable `.oxt` file.
```bash
VIMBREWRITER_VERSION="1.0.0" make extension
```
- Requires `VIMBREWRITER_VERSION` to be set.
- Output is placed in the `dist/` directory.

### 3. Installing (`make install`)
Builds and installs the extension to your current LibreOffice user profile.
```bash
VIMBREWRITER_VERSION="1.0.0" make install
```
- Ensure LibreOffice is closed before running this.

### 4. Cleaning (`make clean`)
Removes the `build/` directory and temporary artifacts.
```bash
make clean
```

## detailed Bindings

### Movement
| Key | Description |
| :--- | :--- |
| `h` | Move left |
| `j` | Move down |
| `k` | Move up |
| `l` | Move right |
| `w` / `W` | Move forward one word |
| `b` / `B` | Move backward one word |
| `e` | Move to end of word |
| `0` | Move to start of line |
| `^` | Move to first non-blank character of line |
| `$` | Move to end of line |
| `gg` | Move to start of document |
| `G` | Move to end of document |
| `}` | Move forward one paragraph |
| `{` | Move backward one paragraph |
| `)` | Move forward one sentence |
| `(` | Move backward one sentence |
| `C-d` | Scroll down (half screen) |
| `C-u` | Scroll up (half screen) |

### Insert Mode Entry
| Key | Description |
| :--- | :--- |
| `i` | Insert before cursor |
| `I` | Insert at start of line |
| `a` | Append after cursor |
| `A` | Append at end of line |
| `o` | Open new line below |
| `O` | Open new line above |

### Editing (Operators)
Operators can be combined with motions (e.g., `d` + `w` = delete word).

| Key | Description |
| :--- | :--- |
| `d` | Delete |
| `c` | Change (Delete + Insert Mode) |
| `y` | Yank (Copy to clipboard) |
| `x` | Delete character under cursor |
| `r` | Replace single character |
| `p` | Paste after cursor |
| `P` | Paste before cursor |
| `u` | Undo |
| `C-r` | Redo |

### Search
- `f{char}`: Move to next occurrence of `{char}` on line.
- `F{char}`: Move to previous occurrence of `{char}` on line.
- `t{char}`: Move till next occurrence of `{char}` on line.
- `T{char}`: Move till previous occurrence of `{char}` on line.
- `/`: Focus Find Bar.
- `\`: Open Find & Replace Dialog.

### Visual Mode
- `v`: Enter Visual Mode (character selection).
- `V`: Enter Visual Line Mode (line selection).
- Standard movement keys extend the selection.
- Operators (`d`, `c`, `y`) work on the selection.
