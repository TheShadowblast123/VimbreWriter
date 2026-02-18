# VimbreWriter

*(Unfortunately, vimbrewriter, formerly vibreoffice, was in an experimental stage, and the creator [seanyeh](https://github.com/seanyeh) no longer has much time to work on it. So I guess I'm learning Visual basic to pick up the mantel)*

VimbreWriter is an extension for Libreoffice and OpenOffice that brings some of
your favorite key bindings from vi/vim to your favorite office suite, with a heavy focus on librewriter.

### Installation/Usage

The easiest way to install is to download the
[latest extension file](https://raw.github.com/TheShadowblast123/vimbrewriter/master/dist/vimbrewriter-1.1.0.oxt)
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
- Movement keys: `hjkl`, `w`, `W`, `b`, `B`, `e`, `$`, `^`, `{}`, `()`, `C-d`, `C-u`, `gg`, `G`, `0`
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
    - example: `ca,` Change Around Commas
- (`d`/`c`)(`a`/`i`)`u`(`any`/`any`) Delete/Change Around/Inside almost any two symbols example
    - example: `diu:.` Delete Inside Unmatched Colon and Period


### Known differences/issues
- tag block deletion doesn't work, but I don't believe it's necessary,
  especially given the `iu` and `au` functionality.
- Movement will differ between Vi and LibreOffice. The goal is close enough.
- Unlike vi/vim, movement keys will wrap to the next line
- Due to line wrapping, you may find your cursor move a bit for commands that
  that would otherwise leave you in the same position (such as `dd`) this has
  mostly been mitigated but a side effect is `dd` will not deleted entire
  paragraphs

### I found a bug/How do I help
- If you find a bug or have a suggestion follow these simple steps:
    - check if there's a pre-existing issue
        - If there is and the issue is closed, I didn't want to deal with it
            - either create a pull request or create an issue
        - If there is and it's open, give information on that issue!
    - No issue? Create one!
    - If you know how to program, create a pull request
        - It is best to create an issue first, or explain what the code is doing
    - profit   


### License
vimbrewriter is released under the MIT License.
