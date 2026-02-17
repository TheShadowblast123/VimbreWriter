# VimbreWriter

*(Unfortunately, vimbrewriter was in an experimental stage, and the creator [seanyeh](https://github.com/seanyeh) no longer has much time to work on it. So I guess I'm learning Visual basic to pick up the mantel)*

VimbreWriter is an extension for Libreoffice and OpenOffice that brings some of
your favorite key bindings from vi/vim to your favorite office suite, with a heavy focus on librewriter.

Most of the rest of this readme is just going to be the original readme until I actually know what I'm doing.
The project is now in a working state with the most up to date version of libre office

### Installation/Usage

The easiest way to install is to download the
[latest extension file](https://raw.github.com/TheShadowblast123/vimbrewriter/master/dist/vimbrewriter-1.0.0.oxt)
and open it with LibreOffice/OpenOffice.

To enable/disable vimbrewriter, simply select Tools -> Add-Ons -> vimbrewriter.

If you really want to, you can build the .oxt file yourself by running
```shell
# replace 0.0.0 with your desired version number
VIMBREWRITER_VERSION="0.0.0" make extension
```
This will simply build the extension file from the template files in
`extension/template`. These template files were auto-generated using
[Extension Compiler](https://wiki.openoffice.org/wiki/Extensions_Packager#Download).


### Features

vimbrewriter currently supports:
- Insert (`i`, `I`, `a`, `A`, `o`, `O`), Visual (`v`), Normal modes
- Movement keys: `hjkl`, `w`, `W`, `b`, `B`, `e`, `$`, `^`, `{}`, `()`, `C-d`, `C-u`
    - Search movement: `f`, `F`, `t`, `T`
- Number modifiers: e.g. `5w`, `4fa`
- Replace: `r`
- Deletion: `x`, `d`, `c`, `s`, `D`, `C`, `S`, `dd`, `cc`
    - Plus movement and number modifiers: e.g. `5dw`, `c3j`, `2dfe`
    - Delete a/inner block: e.g. `di(`, `da{`, `ci[`, `ci"`, `ca'`
- Undo/redo: `u`, `C-r`
- Copy/paste: `y`, `p`, `P` (using system clipboard, not vim-like registers)
### Added functionality
- (`d`/`c`)(`a`/`i`)(`,`/`.`) Delete/Change Around/Inside Commas/Periods for enhanced word editing


### Known differences/issues

If you are familiar with vi/vim, then vimbrewriter should give very few
surprises. However, there are some differences, primarily due to word
processor-text editor differences or limitations of the LibreOffice API and/or
my patience.
- I'm pretty sure I broke `dit`
- Currently, I am using LibreOffice's built-in word/sentence movement which
  differs from vi's. It's sort of broken now but I plan to fix it eventually.
- The concept of lines in a text editor is not quite analogous to that of a
  word processor. I made my best effort to incorporate the line analogy while keeping
  the spirit of word processing.
    - Unlike vi/vim, movement keys will wrap to the next line
    - Due to line wrapping, you may find your cursor move up/down a line for
      commands that would otherwise leave you in the same position (such as `dd`)
      this has mostly been  mitigated but a side effect is `dd` will not deleted
      entire paragraphs

vimbrewriter is new, so it is bound to have plenty of bugs. Please let me know
if you run into anything!


### License
vimbrewriter is released under the MIT License.
