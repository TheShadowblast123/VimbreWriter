' vibreoffice - Vi Mode for LibreOffice/OpenOffice
'
' The MIT License (MIT)
'
' Copyright (c) 2014 Sean Yeh
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

Option Explicit

' --------
' Globals
' --------
global VIBREOFFICE_STARTED as boolean ' Defaults to False
global VIBREOFFICE_ENABLED as boolean ' Defaults to False
global oXKeyHandler as object
global MODE as String
global VIEW_CURSOR as object
global MULTIPLIER as integer
global VISUAL_BASE as object ' Position of first selected line in VISUAL_LINE
global LAST_PAGE as integer
global oSelectionListener as object
global PREV_MODE as String
global SPECIAL_MODE as String
global SPECIAL_COUNT as integer
global MOVEMENT_MODIFIER as String

' -----------
' Main entry point (toggle plugin)
' -----------
Sub Main
    If Not VIBREOFFICE_STARTED Then
        initVibreoffice()
    End If

    ' Toggle enable/disable
    VIBREOFFICE_ENABLED = Not VIBREOFFICE_ENABLED

    ' Restore statusbar
    If Not VIBREOFFICE_ENABLED Then restoreStatus()
End Sub

' -----------
' Initialisation
' -----------
Sub initVibreoffice
    Dim oTextCursor As Object
    ' Initializing
    VIBREOFFICE_STARTED = True
    VIEW_CURSOR = thisComponent.getCurrentController.getViewCursor

    resetMultiplier()
    gotoMode("NORMAL")

    ' Show terminal cursor
    oTextCursor = getTextCursor()

    If oTextCursor Is Nothing Then
        ' Do nothing
    Else
        cursorReset(oTextCursor)
    End If

    sStartXKeyHandler()

    ' selection change listener to track page changes
    LAST_PAGE = getPageNum()
    oSelectionListener = CreateUnoListener("SelectionChange_", "com.sun.star.view.XSelectionChangeListener")
    ThisComponent.CurrentController.addSelectionChangeListener(oSelectionListener)
End Sub

' -------------
' Key Handling
' -------------
Sub sStartXKeyHandler
    sStopXKeyHandler()

    oXKeyHandler = CreateUnoListener("KeyHandler_", "com.sun.star.awt.XKeyHandler")
    thisComponent.CurrentController.AddKeyHandler(oXKeyHandler)
End Sub

Sub sStopXKeyHandler
    thisComponent.CurrentController.removeKeyHandler(oXKeyHandler)
    If Not IsNull(oSelectionListener) Then
        ThisComponent.CurrentController.removeSelectionChangeListener(oSelectionListener)
    End If
End Sub

Sub XKeyHandler_Disposing(oEvent)
End Sub

' --------------------
' Main Key Processing
' --------------------
Function KeyHandler_KeyPressed(oEvent) as boolean
    Dim oTextCursor As Object

    ' Exit if plugin is not enabled
    If IsMissing(VIBREOFFICE_ENABLED) Or Not VIBREOFFICE_ENABLED Then
        KeyHandler_KeyPressed = False
        Exit Function
    End If

    ' Exit if TextCursor does not work (as in Annotations)
    oTextCursor = getTextCursor()
    If oTextCursor Is Nothing Then
        KeyHandler_KeyPressed = False
        Exit Function
    End If

    Dim bConsumeInput, bIsMultiplier, bIsModified, bIsSpecial As Boolean
    bConsumeInput = True ' Block all inputs by default
    bIsMultiplier = False ' reset multiplier by default
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    bIsSpecial = getSpecial() <> ""


    ' --------------------------
    ' Process global shortcuts, exit if matched (like ESC)
    If ProcessGlobalKey(oEvent) Then
        ' Pass

    ElseIf (MODE = "NORMAL" Or MODE = "VISUAL" Or MODE = "VISUAL_LINE") And oEvent.KeyChar = 58 And getSpecial() = "" And getMovementModifier() = "" Then
        ' Colon pressed – enter command mode
        PREV_MODE = MODE
        gotoMode("CMD")
        bConsumeInput = True

        ' If INSERT mode, allow all inputs
    ElseIf MODE = "INSERT" Then
        bConsumeInput = False

    ElseIf MODE = "CMD" Then
        ' In command mode: the next key is the command (e.g., 'b' for bold)
        HandleCommand(oEvent.KeyChar)

        If PREV_MODE <> "" Then
            setMode(PREV_MODE)
        Else
            gotoMode("NORMAL")
        End If
        bConsumeInput = True

        ' If Change Mode
        ' ElseIf MODE = "NORMAL" And Not bIsSpecial And getMovementModifier() = "" And ProcessModeKey(oEvent) Then
    ElseIf ProcessModeKey(oEvent) Then
        ' Pass

        ' Replace Key
    ElseIf getSpecial() = "r" And Not bIsModified Then
        Dim iLen as Integer
        iLen = Len(getCursor().getString())
        getCursor().setString(Chr(oEvent.KeyChar))

        ' Multiplier Key
    ElseIf ProcessNumberKey(oEvent) Then
        bIsMultiplier = True
        delaySpecialReset()

        ' Normal Key
    ElseIf ProcessNormalKey(oEvent.KeyChar, oEvent.Modifiers) Then
        ' Pass

        ' If is modified but doesn't match a normal command, allow input
        '   (Useful for built-in shortcuts like Ctrl+s, Ctrl+w)
    ElseIf bIsModified Then
        bConsumeInput = False

        ' Movement modifier here?
    ElseIf ProcessMovementModifierKey(oEvent.KeyChar) Then
        delaySpecialReset()

        ' If standard movement key (in VISUAL or VISUAL_LINE mode) like arrow keys, home, end
    ElseIf (MODE = "VISUAL" Or MODE = "VISUAL_LINE") And ProcessStandardMovementKey(oEvent) Then
        ' Pass

        ' If bIsSpecial but nothing matched, return to normal mode
    ElseIf bIsSpecial Then
        gotoMode("NORMAL")

        ' Allow non-letter keys if unmatched
    ElseIf oEvent.KeyChar = 0 Then
        bConsumeInput = False
    End If
    ' --------------------------

    ' Reset Special
    resetSpecial()

    ' Reset multiplier if last input was not number and not in special mode
    If Not bIsMultiplier And getSpecial() = "" And getMovementModifier() = "" Then
        resetMultiplier()
    End If
    setStatus(getMultiplier())

    ' Update the terminal-style cursor appearance after every keypress.
    ' Done here in KeyPressed rather than KeyReleased because KeyReleased
    ' must return True (consume the event) to prevent the key firing twice.
    oTextCursor = getTextCursor()
    If Not (oTextCursor Is Nothing) Then
        If MODE = "NORMAL" Then
            cursorReset(oTextCursor)
        ElseIf MODE = "INSERT" Then
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            thisComponent.getCurrentController.Select(oTextCursor)
        End If
    End If

    KeyHandler_KeyPressed = bConsumeInput
End Function

Function KeyHandler_KeyReleased(oEvent) As boolean
    ' Always consume the KeyReleased event in NORMAL and VISUAL mode.
    ' Returning False lets the release propagate to LibreOffice, which
    ' re-processes it as input — causing every command to fire twice.
    ' INSERT mode passes releases through so the application can handle
    ' auto-complete, IME, etc. normally.
    KeyHandler_KeyReleased = (MODE <> "INSERT")
End Function

' ----------------
' Processing Keys
' ----------------
Function ProcessGlobalKey(oEvent)
    Dim bMatched, bIsControl as Boolean
    bMatched = True
    bIsControl = (oEvent.Modifiers = 2) Or (oEvent.Modifiers = 8)

    ' PRESSED ESCAPE (or ctrl+[)
    ' KeyCode 1281 = Escape, KeyCode 1315 = '[', ascii 91
    If oEvent.KeyCode = 1281 Or (oEvent.KeyCode = 1315 And bIsControl) Then
        ' Move cursor back if was in INSERT (but stay on same line)
        If MODE <> "NORMAL" And Not getCursor().isAtStartOfLine() Then
            getCursor().goLeft(1, False)
        End If

        resetSpecial(True)
        gotoMode("NORMAL")
    Else
        bMatched = False
    End If
    ProcessGlobalKey = bMatched
End Function

Function ProcessModeKey(oEvent)
    Dim bIsModified, key as Integer
    Dim bMatched as Boolean

    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    ' Don't change modes in these circumstances
    If MODE <> "NORMAL" Or bIsModified Or getSpecial() <> "" Or getMovementModifier() <> "" Then
        ProcessModeKey = False
        Exit Function
    End If

    key = oEvent.KeyChar

    ' Mode matching
    bMatched = True
    Select Case key
            ' Insert modes
            ' 105='i', 97='a', 73='I', 65='A', 111='o', 79='O'
        Case 105, 97, 73, 65, 111, 79 : ' i,a,I,A,o,O
            If key = 97 Then getCursor().goRight(1, False)
            If key = 73 Then ProcessMovementKey(94, 1, 0)
            If key = 65 Then ProcessMovementKey(36, 1, 0)

            If key = 111 Then ' 'o': open line below
                ProcessMovementKey(36, 1, 0) ' '$'
                ProcessMovementKey(108, 1, 0) ' 'l'
                getCursor().setString(Chr(13))
                If Not getCursor().isAtStartOfLine() Then
                    getCursor().setString(Chr(13) & Chr(13))
                    ProcessMovementKey(108, 1, 0)
                End If
            End If

            If key = 79 Then ' 'O': open line above
                ProcessMovementKey(94, 1, 0) ' '^'
                getCursor().setString(Chr(13))
                If Not getCursor().isAtStartOfLine() Then
                    ProcessMovementKey(104, 1, 0) ' 'h'
                    getCursor().setString(Chr(13))
                    ProcessMovementKey(108, 1, 0) ' 'l'
                End If
            End If

            gotoMode("INSERT")
        Case 118 : ' 118='v'
            gotoMode("VISUAL")

        Case 86 : ' 86='V'
            gotoMode("VISUAL_LINE")
        Case Else :
            bMatched = False
    End Select
    ProcessModeKey = bMatched
End Function

Sub HandleCommand(keyChar As Integer)
    Dim oTextCursor, oViewCursor, oStartPos, dispatcher As Object
    Dim iMultiplier As Integer


    oViewCursor = getCursor()
    oTextCursor = getTextCursor()
    iMultiplier = getMultiplier()

    ' If no selection, select the current word
    If oTextCursor.isCollapsed() Then
        Set oTextCursor = oViewCursor.getText().createTextCursorByRange(oViewCursor.getStart())
        oTextCursor.gotoStartOfWord(False)
        oTextCursor.gotoEndOfWord(True)
        If oTextCursor.isCollapsed() Then
            ' No word under cursor (e.g., whitespace) – do nothing
            Exit Sub
        End If
        Set oStartPos = oTextCursor.getStart()
        ' Set the controller’s selection to this word
        ThisComponent.getCurrentController().Select(oTextCursor)
    Else
        Set oStartPos = oTextCursor.getStart()
        ' Selection already active, no need to change controller selection
    End If

    Set dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    Select Case keyChar
        Case 98 '  b -> Bold
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Bold", "", 0, Array())
        Case 101 ' e -> Align center
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:CommonAlignHorizontalCenter", "", 0, Array())
        Case 104 ' h -> Highlight (yellow background)
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:CharBackColor", "", 0, Array())
        Case 105 ' i -> Italic
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Italic", "", 0, Array())
        Case 106 ' j -> Justify
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:CommonAlignJustified", "", 0, Array())
        Case 108 ' l -> Align left
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:CommonAlignLeft", "", 0, Array())
        Case 112 ' p -> paste 
            pasteSelection False, True, iMultiplier
        Case 80 ' P -> unformattd paste
            pasteSelection True, False, iMultiplier
        Case 114 ' r -> Align right
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:CommonAlignRight", "", 0, Array())
        Case 115 ' s -> Subscript
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:SubScript", "", 0, Array())
        Case 83 '  S -> Superscript
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:SuperScript", "", 0, Array())
        Case 116 ' t -> Strike through
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Strikeout", "", 0, Array())
        Case 117 ' u -> Underline
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Underline", "", 0, Array())
        Case 119 ' w -> save Document
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Save", "", 0, Array())
        Case 93 ' ] -> Indent increase
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:IncrementIndent", "", 0, Array())
        Case 91 ' [ -> Indent decrease
            dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:DecrementIndent", "", 0, Array())
        Case Else
            ' Unknown command – do nothing and exit
            Exit Sub
    End Select

    ' Return to Normal mode and place cursor at the start of the formatted range
End Sub

Function ProcessNormalKey(keyChar, modifiers)

    Dim i, bMatched, bIsVisual, iMultiplier, iRawMultiplier, bIsControl As Boolean

    ' Set dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    bIsControl = (modifiers = 2) Or (modifiers = 8)

    bIsVisual = (MODE = "VISUAL" Or MODE = "VISUAL_LINE")

    iMultiplier = getMultiplier()
    iRawMultiplier = getRawMultiplier()

    ' ----------------------
    ' 1. Check Movement Key
    ' ----------------------

    If keyChar = 48 Or keyChar = 94 Then ' 48='0', 94='^'
        iMultiplier = 1
        iRawMultiplier = 0 ' optional, not strictly needed
    End If

    bMatched = ProcessMovementKey(keyChar, iMultiplier, iRawMultiplier, bIsVisual, modifiers)

    ' If Special: d/c + movement
    If bMatched And (getSpecial() = "d" Or getSpecial() = "c" Or getSpecial() = "y") Then
        yankSelection((getSpecial() <> "y"))
    End If

    ' Reset Movement Modifier
    setMovementModifier("")

    ' Exit already if movement key was matched
    If bMatched Then
        ' If Special: d/c : change mode
        If getSpecial() = "d" Or getSpecial() = "y" Then gotoMode("NORMAL")
        If getSpecial() = "c" Then gotoMode("INSERT")

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 2. Undo/Redo, USE DEFAULT ctrl+z and ctrl+y
    ' --------------------
    ' 117='u', 114='r' (Ctrl+r = redo)
    ' If keyChar = 117 Or (bIsControl And keyChar = 114) Then ' u or Ctrl+r
    '     For i = 1 To iMultiplier
    '         Undo(keyChar = 117)
    '     Next i
    '
    '     ProcessNormalKey = True
    '     Exit Function
    ' End If


    ' --------------------
    ' 3. Paste
    '   Note: in vim, paste will result in cursor being over the last character of the pasted content. Here, the cursor will be the next character after that. Fix?
    ' --------------------
    ' 112='p', 80='P'
    If keyChar = 112 Then ' p
        pasteSelection False, True, iMultiplier
        ProcessNormalKey = True

        Exit Function
    End If

    If keyChar = 80 Then ' P
        pasteSelection True, False, iMultiplier
        ProcessNormalKey = True
        Exit Function
    End If

    '---------------------------------------------------
    '               Find and Replace
    '--------------==-----------------------------------
    If keyChar = 47 Then
        Dim frameFind, dispatcherFind as Object

        frameFind = thisComponent.CurrentController.Frame
        dispatcherFind = createUnoService("com.sun.star.frame.DispatchHelper")

        ' Uses the specific findbar protocol to focus the search bar
        dispatcherFind.executeDispatch(frameFind, "vnd.sun.star.findbar:FocusToFindbar", "", 0, Array())

        ProcessNormalKey = True
        Exit Function
    End If

    ' 92 is the key code for '\'
    If keyChar = 92 Then ' \
        Dim frame as Object
        Dim dispatcherSearchDialog as Object
        frame = thisComponent.CurrentController.Frame
        dispatcherSearchDialog = createUnoService("com.sun.star.frame.DispatchHelper")

        ' Launches the built-in Find & Replace dialog
        dispatcherSearchDialog.executeDispatch(frame, ".uno:SearchDialog", "", 0, Array())

        ProcessNormalKey = True
        Exit Function
    End If

    ' --------------------
    ' 4. Check Special/Delete Key
    ' --------------------

    ' There are no special/delete keys with modifier keys, so exit early
    If modifiers > 1 Then
        ProcessNormalKey = False
        Exit Function
    End If

    ' Only 'x' or Special (dd, cc) can be done more than once
    ' 120='x'
    If keyChar <> 120 And getSpecial() = "" Then ' 120='x'
        iMultiplier = 1
    End If
    For i = 1 To iMultiplier
        Dim bMatchedSpecial as Boolean

        ' Special/Delete Key
        bMatchedSpecial = ProcessSpecialKey(keyChar)

        bMatched = bMatched Or bMatchedSpecial
    Next i


    ProcessNormalKey = bMatched
End Function

Function ProcessSpecialKey(keyChar)
    Dim oTextCursor, bMatched, bIsSpecial, bIsDelete as Boolean
    Dim savedCol As Integer
    bMatched = True
    bIsSpecial = getSpecial() <> ""

    Select Case keyChar
        Case 100, 99, 115, 121 ' d,c,s,y
            bIsDelete = (keyChar <> 121)

            If bIsSpecial Then
                Dim bIsSpecialCase
                bIsSpecialCase = (keyChar = 100 And getSpecial() = "d") Or (keyChar = 99 And getSpecial() = "c")
                If bIsSpecialCase Then
                    savedCol = GetCursorColumn()
                    Dim oLineCursor as Object
                    Set oLineCursor = GetLineTextCursor(True)
                    thisComponent.getCurrentController.Select(oLineCursor)
                    yankSelection(bIsDelete)
                    If bIsDelete Then
                        SetCursorColumn(savedCol)
                    End If
                Else
                    bMatched = False
                End If
                If bIsSpecialCase And keyChar = 99 Then
                    gotoMode("INSERT")
                Else
                    gotoMode("NORMAL")
                End If
            ElseIf MODE = "VISUAL" Or MODE = "VISUAL_LINE" Then
                oTextCursor = getTextCursor()
                thisComponent.getCurrentController.Select(oTextCursor)
                yankSelection(bIsDelete)
                If keyChar = 99 Or keyChar = 115 Then gotoMode("INSERT")
                If keyChar = 100 Or keyChar = 121 Then gotoMode("NORMAL")
            ElseIf MODE = "NORMAL" Then
                If keyChar = 115 Then
                    setSpecial("c")
                    gotoMode("VISUAL")
                    ProcessNormalKey(108, 0) ' 108='l'
                Else
                    setSpecial(Chr(keyChar))
                    gotoMode("VISUAL")
                End If
            End If

        Case 114 ' r
            setSpecial("r")

        Case Else
            If bIsSpecial Then
                bMatched = False
            Else
                Select Case keyChar
                    Case 120 ' x
                        oTextCursor = getTextCursor()
                        thisComponent.getCurrentController.Select(oTextCursor)
                        yankSelection(True)
                        cursorReset(oTextCursor)
                        gotoMode("NORMAL")

                    Case 68, 67 ' D, C
                        If MODE = "VISUAL" Or MODE = "VISUAL_LINE" Then
                            ProcessMovementKey(94, 1, 0, False)
                            ProcessMovementKey(36, 1, 0, True)
                            ProcessMovementKey(108, 1, 0, True)
                        Else
                            oTextCursor = getTextCursor()
                            oTextCursor.gotoRange(oTextCursor.getStart(), False)
                            thisComponent.getCurrentController.Select(oTextCursor)
                            ProcessMovementKey(36, 1, 0, True)
                        End If
                        yankSelection(True)
                        If keyChar = 68 Then
                            gotoMode("NORMAL")
                        Else
                            gotoMode("INSERT")
                        End If

                    Case 83 ' S
                        If MODE = "NORMAL" Then
                            ProcessMovementKey(94, 1, 0, False)
                            ProcessMovementKey(36, 1, 0, True)
                            yankSelection(True)
                            gotoMode("INSERT")
                        Else
                            bMatched = False
                        End If

                    Case Else
                        bMatched = False
                End Select
            End If
    End Select

    ProcessSpecialKey = bMatched
End Function

Function ProcessMovementKey(keyChar, iMultiplier, iRawMultiplier, Optional bExpand, Optional keyModifiers)
    Dim oTextCursor, bSetCursor, bMatched, i, dispatcher
    oTextCursor = getTextCursor()
    bMatched = True
    If IsMissing(bExpand) Then bExpand = False
    If IsMissing(keyModifiers) Then keyModifiers = 0


    ' Check for modified keys (Ctrl, Alt, not Shift)
    If keyModifiers > 1 Then
        Dim bIsControl, bIsCtrlShift
        bIsControl = (keyModifiers = 2) Or (keyModifiers = 8)
        bIsCtrlShift = (keyModifiers = 3) Or (keyModifiers = 9)

        ' Ctrl+Shift+> (Page Down)
        If bIsCtrlShift And (keyChar = 62 Or keyChar = 46) Then ' > or .
            dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
            For i = 1 To iMultiplier
                If bExpand Then
                    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:PageDownSel", "", 0, Array())
                Else
                    getCursor().ScreenDown(False)
                End If
            Next i
            ' Ctrl+Shift+< (Page Up)
        ElseIf bIsCtrlShift And (keyChar = 60 Or keyChar = 44) Then ' < or ,
            dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
            For i = 1 To iMultiplier
                If bExpand Then
                    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:PageUpSel", "", 0, Array())
                Else
                    getCursor().ScreenUp(False)
                End If
            Next i
        Else
            bMatched = False
        End If
        ProcessMovementKey = bMatched
        Exit Function
    End If

    ' Set global cursor to oTextCursor's new position if moved
    bSetCursor = True


    ' ------------------
    ' Movement matching
    ' ------------------

    ' ---------------------------------
    ' Special Case: Modified movements (f/t/F/T/i/a + target key)
    If getMovementModifier() <> "" Then
        Select Case getMovementModifier()
                ' f,F,t,T searching — convert int keyChar to string for search
            Case "f", "t", "F", "T" :
                bMatched = False
                For i = 1 To iMultiplier
                    If ProcessSearchKey(oTextCursor, getMovementModifier(), Chr(keyChar), bExpand) Then
                        bMatched = True
                    Else
                        Exit For
                    End If
                Next i
            Case "a", "i" :
                bMatched = GetSymbol(keyChar, getMovementModifier())
                bSetCursor = False
            Case Else :
                bSetCursor = False
                bMatched = False
        End Select

        If Not bMatched Then
            bSetCursor = False
        End If
        ' ---------------------------------

    Else
        Select Case keyChar
            Case 108 ' l
                For i = 1 To iMultiplier : oTextCursor.goRight(1, bExpand) : Next i

            Case 104 ' h
                For i = 1 To iMultiplier : oTextCursor.goLeft(1, bExpand) : Next i

            Case 107 ' k
                If MODE = "VISUAL_LINE" Then
                    For i = 1 To iMultiplier
                        Dim lastSelected
                        If getCursor().getPosition().Y() <= VISUAL_BASE.Y() Then
                            lastSelected = getCursor().getPosition().Y()
                            If VISUAL_BASE.Y() = getCursor().getPosition().Y() Then
                                getCursor().gotoEndOfLine(False)
                                If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                                    getCursor().goRight(1, False)
                                End If
                            End If
                            Do Until getCursor().getPosition().Y() < lastSelected
                                If Not getCursor().goUp(1, bExpand) Then Exit Do
                            Loop
                            getCursor().gotoStartOfLine(bExpand)
                        ElseIf getCursor().getPosition().Y() > VISUAL_BASE.Y() Then
                            getCursor().goUp(1, bExpand)
                            lastSelected = getCursor().getPosition().Y()
                            getCursor().goUp(1, bExpand)
                            If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                                formatVisualBase()
                            Else
                                getCursor().gotoEndOfLine(bExpand)
                                If getCursor().getPosition().Y() < lastSelected Then
                                    getCursor().goRight(1, bExpand)
                                End If
                            End If
                        End If
                    Next i
                    bSetCursor = False
                Else
                    For i = 1 To iMultiplier : getCursor().goUp(1, bExpand) : Next i
                    bSetCursor = False
                End If

            Case 106 ' j
                If MODE = "VISUAL_LINE" Then
                    For i = 1 To iMultiplier
                        If getCursor().getPosition().Y() >= VISUAL_BASE.Y() Then
                            If VISUAL_BASE.Y() = getCursor().getPosition().Y() Then
                                getCursor().gotoStartOfLine(False)
                                getCursor().gotoEndOfLine(bExpand)
                                If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                                    getCursor().goRight(1, bExpand)
                                End If
                            End If
                            If getCursor().goDown(1, bExpand) Then
                                getCursor().gotoStartOfLine(bExpand)
                            Else
                                getCursor().gotoEndOfLine(bExpand)
                            End If
                        ElseIf getCursor().getPosition().Y() < VISUAL_BASE.Y() Then
                            getCursor().goDown(1, bExpand)
                            getCursor().gotoStartOfLine(bExpand)
                        End If
                    Next i
                    bSetCursor = False
                Else
                    For i = 1 To iMultiplier : getCursor().goDown(1, bExpand) : Next i
                    bSetCursor = False
                End If

            Case 48 ' 0
                getCursor().gotoStartOfLine(bExpand)
                bSetCursor = False

            Case 94 ' ^
                ' --- Move cursor to the first non‑whitespace character on the current line ---
                Dim originalLineY
                originalLineY = getCursor().getPosition().Y()

                ' Temporarily select the entire line to extract its text.
                getCursor().gotoEndOfLine(False)
                If getCursor().getPosition().Y() > originalLineY Then
                    ' If we landed on the next line (line was exactly one character long), move back.
                    getCursor().goLeft(1, False)
                End If
                getCursor().gotoStartOfLine(True)
                Dim lineText As String
                lineText = getCursor().String

                ' Restore the original cursor (and any existing selection) before moving.
                getCursor().gotoRange(oTextCursor, False)
                getCursor().gotoStartOfLine(bExpand) ' Start from column 0

                ' Find the index (1‑based) of the first character that is not a space or tab.
                Dim firstNonBlankPos As Integer
                firstNonBlankPos = 1
                Do While firstNonBlankPos <= Len(lineText)
                    Dim ch As String
                    ch = Mid(lineText, firstNonBlankPos, 1)
                    If ch <> " " And ch <> Chr(9) Then ' space or tab
                        Exit Do
                    End If
                    firstNonBlankPos = firstNonBlankPos + 1
                Loop

                ' If the line is empty or all whitespace, stay at column 0.
                If firstNonBlankPos > Len(lineText) Then
                    firstNonBlankPos = 1
                End If

                ' Move right to the first non‑blank character.
                getCursor().goRight(firstNonBlankPos - 1, bExpand)

                bSetCursor = False ' We already moved the view cursor directly

            Case 36 ' $
                Dim oldPos, newPos as Object
                oldPos = getCursor().getPosition()
                getCursor().gotoEndOfLine(bExpand)
                newPos = getCursor().getPosition()

                ' If the result is at the start of the line, then it must have
                ' jumped down a line; goLeft to return to the previous line.
                '   Except for: Empty lines (check for oldPos = newPos)
                If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
                    getCursor().goLeft(1, bExpand)
                End If

                bSetCursor = False

            Case 119, 87 ' w, W
                For i = 1 To iMultiplier : oTextCursor.gotoNextWord(bExpand) : Next i

            Case 98, 66 ' b, B
                For i = 1 To iMultiplier
                    If Not MoveWordBackward(oTextCursor, bExpand) Then
                        Exit For
                    End If
                Next i
                bSetCursor = True

            Case 103 ' g
                If iRawMultiplier > 0 Then
                    Dim targetPage, itotalPages As Integer
                    targetPage = iMultiplier
                    itotalPages = getPageCount()
                    If targetPage > itotalPages Then targetPage = itotalPages
                    getCursor().jumpToPage(targetPage, bExpand)
                    oTextCursor.gotoRange(getCursor().getStart(), bExpand)
                ElseIf getSpecial() = "g" Then
                    ' Handle gg with visual modes
                    If MODE = "VISUAL" Then
                        Dim oldPosG
                        oldPosG = getCursor().getPosition()
                        getCursor().gotoRange(getCursor().getStart(), True)
                        If Not samePos(getCursor().getPosition(), oldPosG) Then
                            getCursor().gotoRange(getCursor().getEnd(), False)
                        End If
                    ElseIf MODE = "VISUAL_LINE" Then
                        Do Until getCursor().getPosition().Y() <= VISUAL_BASE.Y()
                            getCursor().goUp(1, False)
                        Loop
                        If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                            formatVisualBase()
                        End If
                    End If
                    Dim bExpandLocal
                    bExpandLocal = (MODE = "VISUAL" Or MODE = "VISUAL_LINE")
                    getCursor().gotoStart(bExpandLocal)
                    bSetCursor = False
                    resetSpecial(True)
                Else
                    ' Set special 'g' and wait for the next key
                    setSpecial("g")
                    bMatched = True
                    bSetCursor = False
                End If

            Case 71 ' G
                oTextCursor.gotoEnd(bExpand)

            Case 101 ' e
                For i = 1 To iMultiplier
                    oTextCursor.goRight(1, bExpand)
                    oTextCursor.gotoEndOfWord(bExpand)
                Next i

            Case 41 ' )
                For i = 1 To iMultiplier : oTextCursor.gotoNextSentence(bExpand) : Next i

            Case 40 ' (
                For i = 1 To iMultiplier : oTextCursor.gotoPreviousSentence(bExpand) : Next i

            Case 125 ' }
                For i = 1 To iMultiplier : oTextCursor.gotoNextParagraph(bExpand) : Next i

            Case 123 ' {
                For i = 1 To iMultiplier : oTextCursor.gotoPreviousParagraph(bExpand) : Next i

            Case Else
                bSetCursor = False
                bMatched = False
        End Select
    End If

    ' If oTextCursor was moved, set global cursor to its position
    If bSetCursor Then
        getCursor().gotoRange(oTextCursor.getStart(), False)
    End If

    ' If oTextCursor was moved and is in VISUAL mode, update selection
    If bSetCursor And bExpand Then
        thisComponent.getCurrentController.Select(oTextCursor)
    End If

    ProcessMovementKey = bMatched
End Function

Function ProcessMovementModifierKey(keyChar)
    Dim bMatched as Boolean

    bMatched = True
    ' 102='f', 116='t', 70='F', 84='T', 105='i', 97='a',

    Select Case keyChar
        Case 102 : setMovementModifier("f") ' 'f'
        Case 116 : setMovementModifier("t") ' 't'
        Case 70 : setMovementModifier("F") ' 'F'
        Case 84 : setMovementModifier("T") ' 'T'
        Case 105 : setMovementModifier("i") ' 'i'
        Case 97 : setMovementModifier("a") ' 'a'

        Case Else :
            bMatched = False
    End Select

    ProcessMovementModifierKey = bMatched
End Function

Function ProcessNumberKey(oEvent)
    ' If a movement modifier (f/F/t/T) is active, digits are target chars, not multipliers
    If getMovementModifier() <> "" Then
        ProcessNumberKey = False
        Exit Function
    End If

    Dim key as Integer
    key = oEvent.KeyChar

    ' 48='0' through 57='9'
    If key >= 48 And key <= 57 Then
        Dim digit As Integer
        digit = key - 48
        ' If multiplier is 0 and the digit is 0, this is the '0' motion, not a multiplier
        If getRawMultiplier() = 0 And digit = 0 Then
            ProcessNumberKey = False
            Exit Function
        End If
        ' Otherwise add to multiplier
        addToMultiplier(digit)
        ProcessNumberKey = True
    Else
        ProcessNumberKey = False
    End If
End Function

Function ProcessStandardMovementKey(oEvent)
    Dim c, bMatched, dispatcher, i, nMult
    c = oEvent.KeyCode

    bMatched = True
    If (MODE <> "VISUAL" And MODE <> "VISUAL_LINE") Then
        bMatched = False
    Else
        Select Case c
            Case 1024 ' Down arrow
                ProcessMovementKey(106, 1, 0, True) ' 106='j'
            Case 1025 ' Up arrow
                ProcessMovementKey(107, 1, 0, True) ' 107='k'
            Case 1026 ' Left arrow
                ProcessMovementKey(104, 1, 0, True) ' 104='h'
            Case 1027 ' Right arrow
                ProcessMovementKey(108, 1, 0, True) ' 108='l'
            Case 1028 ' Home
                ProcessMovementKey(94, 1, 0, True) ' 94='^'
            Case 1029 ' End
                ProcessMovementKey(36, 1, 0, True) ' 36='$'
            Case 1030 ' Page Up
                dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
                nMult = getMultiplier()
                For i = 1 To nMult
                    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:PageUpSel", "", 0, Array())
                Next i
            Case 1031 ' Page Down
                dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
                nMult = getMultiplier()
                For i = 1 To nMult
                    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:PageDownSel", "", 0, Array())
                Next i
            Case Else
                bMatched = False
        End Select
    End If

    ProcessStandardMovementKey = bMatched
End Function

Function ProcessSearchKey(oTextCursor, searchType, keyChar, bExpand)
    '-----------
    ' Searching
    ' keyChar here is a string (the literal character to find), not an ascii int.
    ' It is passed in from ProcessMovementKey after converting with Chr().
    '-----------
    Dim bMatched, oSearchDesc, oFoundRange, bIsBackwards, oStartRange
    bMatched = True
    bIsBackwards = (searchType = "F" Or searchType = "T")

    If Not bIsBackwards Then
        ' VISUAL mode will goRight AFTER the selection
        If MODE <> "VISUAL" Then
            ' Start searching from next character
            oTextCursor.goRight(1, bExpand)
        End If

        oStartRange = oTextCursor.getEnd()
        ' Go back one
        oTextCursor.goLeft(1, bExpand)
    Else
        oStartRange = oTextCursor.getStart()
    End If

    oSearchDesc = thisComponent.createSearchDescriptor()
    oSearchDesc.setSearchString(keyChar)
    oSearchDesc.SearchCaseSensitive = True
    oSearchDesc.SearchBackwards = bIsBackwards

    oFoundRange = thisComponent.findNext(oStartRange, oSearchDesc)

    If Not IsNull(oFoundRange) Then
        Dim oText, foundPos, curPos, bSearching
        oText = oTextCursor.getText()
        foundPos = oFoundRange.getStart()

        ' Unfortunately, we must go go to this "found" position one character at
        ' a time because I have yet to find a way to consistently move the
        ' Start range of the text cursor and leave the End range intact.
        If bIsBackwards Then
            curPos = oTextCursor.getEnd()
        Else
            curPos = oTextCursor.getStart()
        End If
        Do Until oText.compareRegionStarts(foundPos, curPos) = 0
            If bIsBackwards Then
                bSearching = oTextCursor.goLeft(1, bExpand)
                curPos = oTextCursor.getStart()
            Else
                bSearching = oTextCursor.goRight(1, bExpand)
                curPos = oTextCursor.getEnd()
            End If

            ' Prevent infinite if unable to find, but shouldn't ever happen (?)
            If Not bSearching Then
                bMatched = False
                Exit Do
            End If
        Loop

        If searchType = "t" Then
            oTextCursor.goLeft(1, bExpand)
        ElseIf searchType = "T" Then
            oTextCursor.goRight(1, bExpand)
        End If

    Else
        bMatched = False
    End If

    ' If matched, then we want to select PAST the character
    ' Else, this will counteract some weirdness. hack either way
    If Not bIsBackwards And MODE = "VISUAL" Then
        oTextCursor.goRight(1, bExpand)
    End If

    ProcessSearchKey = bMatched

End Function

Function GetSymbol(keyChar As Integer, modifier As String) As Boolean
    Dim startChar, endChar As String
    Dim bMatched As Boolean
    Dim oTextCursor As Object
    Set oTextCursor = getTextCursor()

    ' Map key code to delimiter pair
    Select Case keyChar
        Case 40, 41 ' (, )
            startChar = "("
            endChar = ")"
        Case 123, 125 ' {, }
            startChar = "{"
            endChar = "}"
        Case 91, 93 ' [, ]
            startChar = "["
            endChar = "]"
        Case 60, 62 ' <, >
            startChar = "<"
            endChar = ">"
        Case 46 ' . 
            startChar = "."
            endChar = "."
        Case 44 ' , 
            startChar = ","
            endChar = ","
        Case 39 ' ' (single quote)
            ' Try smart quotes first
            startChar = Chr(8216) ' ‘
            endChar = Chr(8217) ' ’
            bMatched = FindMatchingPair(startChar, endChar, modifier)
            If Not bMatched Then
                ' Fallback to straight quotes
                startChar = "'"
                endChar = "'"
                bMatched = FindMatchingPair(startChar, endChar, modifier)
            End If
            GetSymbol = bMatched
            Exit Function
        Case 34 ' " (double quote)
            ' Try smart quotes first
            startChar = Chr(8220) ' “
            endChar = Chr(8221) ' ”
            bMatched = FindMatchingPair(startChar, endChar, modifier)
            If Not bMatched Then
                ' Fallback to straight quotes
                startChar = Chr(34)
                endChar = Chr(34)
                bMatched = FindMatchingPair(startChar, endChar, modifier)
            End If
            GetSymbol = bMatched
            Exit Function
        Case Else
            GetSymbol = False
            Exit Function
    End Select

    GetSymbol = FindMatchingPair(startChar, endChar, modifier)
End Function

Function FindMatchingPair(startChar As String, endChar As String, modifier As String) As Boolean
    Dim oDoc, oCursor, oTempCursor As Object
    Dim i, j As Integer
    Dim foundForward As Boolean
    Dim forwardPos As Object

    oDoc = ThisComponent

    ' Search forward for endChar from current view cursor position
    Set oCursor = oDoc.Text.createTextCursorByRange(getCursor().getStart())
    foundForward = False
    For i = 1 To 1000
        If Not oCursor.goRight(1, False) Then Exit For
        Set oTempCursor = oDoc.Text.createTextCursorByRange(oCursor.getStart())
        oTempCursor.goRight(1, True)
        If oTempCursor.getString() = endChar Then
            foundForward = True
            Set forwardPos = oTempCursor.getStart()
            Exit For
        End If
    Next i

    If Not foundForward Then
        FindMatchingPair = False
        Exit Function
    End If

    ' Search backward from forwardPos for startChar
    Set oCursor = oDoc.Text.createTextCursorByRange(forwardPos)
    For j = 1 To 2000
        If Not oCursor.goLeft(1, False) Then Exit For
        Set oTempCursor = oDoc.Text.createTextCursorByRange(oCursor.getStart())
        oTempCursor.goLeft(1, True)
        If oTempCursor.getString() = startChar Then
            Dim startPos, endPos, oEndCursor As Object

            If modifier = "a" Then
                Set startPos = oTempCursor.getStart()
                Set oEndCursor = oDoc.Text.createTextCursorByRange(forwardPos)
                oEndCursor.goRight(1, False)
                Set endPos = oEndCursor.getStart()
            Else ' "i"
                oTempCursor.goRight(1, True)
                Set startPos = oTempCursor.getStart()
                Set endPos = forwardPos
            End If

            ' Move the VIEW cursor to startPos, then select to endPos
            ' getTextCursor() derives from getCursor(), so this is what yankSelection needs
            getCursor().gotoRange(startPos, False)
            getCursor().gotoRange(endPos, True)

            FindMatchingPair = True
            Exit Function
        End If
    Next j

    FindMatchingPair = False
End Function

' Function for both undo and redo
Sub Undo(bUndo)
     On Error Goto ErrorHandler

    If bUndo Then
        thisComponent.getUndoManager().undo()
    Else
        thisComponent.getUndoManager().redo()
    End If
    Exit Sub

    ' Ignore errors from no more undos/redos in stack
    ErrorHandler:
    Resume Next
End Sub


' Yanks selection to system clipboard.
' If bDelete is true, will delete selection.
Sub yankSelection(bDelete)
    Dim dispatcher As Object
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Copy", "", 0, Array())

    If bDelete Then
        getTextCursor().setString("")
    End If
End Sub


Sub pasteSelection(bUnformatted As Boolean, bAfter As Boolean, nMultiplier As Integer)
    Dim oDispatcher As Object
    Dim bWasVisual As Boolean
    bWasVisual = (MODE = "VISUAL" Or MODE = "VISUAL_LINE")

    ' --- Normal mode: deselect and optionally move cursor ---
    If MODE = "NORMAL" Then
        Dim oTempCursor
        oTempCursor = getTextCursor()
        oTempCursor.gotoRange(oTempCursor.getStart(), False)
        thisComponent.getCurrentController.Select(oTempCursor)

        If bAfter Then
            ' Move right one to paste *after* the current character
            getCursor().goRight(1, False)
        End If
    End If

    ' --- Perform paste nMultiplier times ---
    Set oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    Dim i As Integer
    For i = 1 To nMultiplier
        If bUnformatted Then
            oDispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:PasteUnformatted", "", 0, Array())
            getCursor().goRight(1, False)
        Else
            oDispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Paste", "", 0, Array())
            getCursor().goRight(1, False)
        End If
    Next i

    ' --- Position cursor on the last pasted character ---
    If MODE = "NORMAL" Then
        ' In Normal mode, cursor should rest on the last character of pasted text
        getCursor().goLeft(1, False)
    ElseIf bWasVisual Then
        ' If we started in Visual mode, return to Normal and land on last character
        gotoMode("NORMAL")
        getCursor().goLeft(1, False)
    End If
End Sub

' -----------
' Singletons
' -----------
Function getCursor
    getCursor = VIEW_CURSOR
End Function

Function getTextCursor
    Dim oTextCursor
     On Error Goto ErrorHandler
    oTextCursor = getCursor().getText.createTextCursorByRange(getCursor())

    getTextCursor = oTextCursor
    Exit Function

    ErrorHandler:
    ' Text Cursor does not work in some instances, such as in Annotations
    getTextCursor = Nothing
End Function

' -----------------
' Helper Functions
' -----------------
Sub restoreStatus 'restore original statusbar
    Dim oLayout as Object
    oLayout = thisComponent.getCurrentController.getFrame.LayoutManager
    oLayout.destroyElement("private:resource/statusbar/statusbar")
    oLayout.createElement("private:resource/statusbar/statusbar")
End Sub

Sub setRawStatus(rawText)
    thisComponent.Currentcontroller.StatusIndicator.Start(rawText, 0)
End Sub

Function PadRight(text As String, length As Integer) As String
    Dim sText As String
    sText = text
    Do While Len(sText) < length
        sText = sText & " "
    Loop
    If Len(sText) > length Then
        sText = Left(sText, length)
    End If
    PadRight = sText
End Function

Sub setStatus(statusText)
    Dim nTotalPages As Long
    Dim nCurrentPage As Integer
    Dim sFinalStatus, sMode, sStatusText, sSpec, sMod, sPage, sParagraphs, sWords, sFileName As String

    nCurrentPage = getCursor().getPage()
    nTotalPages = getPageCount()

    sMode = PadRight(MODE, 12)
    sStatusText = PadRight(statusText, 5)
    sSpec = PadRight("special: " & getSpecial(), 11)
    sMod = PadRight("modifier: " & getMovementModifier(), 12)
    sPage = PadRight("page: " & nCurrentPage & "/" & nTotalPages, 14)
    sParagraphs = PadRight("paragraphs: " & getParagraphCount(), 17)
    sWords = PadRight("Words: " & getWordCount(), 13)
    sFileName = "File: " & getCurrentFileName()

    sFinalStatus = sMode & " | " & sStatusText & " | " & sSpec & "  | " & sMod & " | " & sPage & " | " & sParagraphs & " | " & sWords & " | " & sFileName

    setRawStatus(sFinalStatus)
End Sub

Sub setMode(modeName)
    MODE = modeName
    setRawStatus(modeName)
End Sub

' --- Added: formatVisualBase for VISUAL_LINE mode ---
Function formatVisualBase()
    Dim oTextCursor as Object
    oTextCursor = getTextCursor()
    VISUAL_BASE = getCursor().getPosition()

    ' Select the current line by moving cursor to start of the line below and
    ' then back to the start of the current line.
    getCursor().gotoEndOfLine(False)
    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
        getCursor().goRight(1, False)
    End If
    getCursor().goLeft(1, True)
    getCursor().gotoStartOfLine(True)
End Function
' ----------------------------------------------------

Function MoveWordBackward(textCursor As Object, shouldExpand As Boolean) As Boolean
    ' Moves the text cursor backward to the start of the previous word.
    ' Returns True if movement occurred, False if at document start.
     On Error Goto ErrorHandler

    Dim isDeleteChangeContext, hasMoved As Boolean

    isDeleteChangeContext = False
    hasMoved = False

    ' Special handling for start of paragraph
    If textCursor.isStartOfParagraph() Then
        If textCursor.goLeft(1, shouldExpand) Then
            hasMoved = True
            If getSpecial() = "c" Or getSpecial() = "d" Then
                isDeleteChangeContext = True
                textCursor.collapseToStart()
            End If
        Else
            ' At document start
            MoveWordBackward = False
            Exit Function
        End If
    Else
        ' If we are at start of a word, move left once to get out of it
        If textCursor.isStartOfWord() Then
            If Not textCursor.goLeft(1, shouldExpand) Then
                MoveWordBackward = False
                Exit Function
            End If
            hasMoved = True
        End If
    End If

    ' Move left until start of word or empty line
    Do While Not (textCursor.isStartOfWord() Or (textCursor.isStartOfParagraph() And textCursor.isEndOfParagraph()))
        If Not textCursor.goLeft(1, shouldExpand) Then
            Exit Do
        End If
        hasMoved = True
    Loop

    If Not hasMoved Then
        MoveWordBackward = False
        Exit Function
    End If

    ' Handle special case for delete/change commands
    If isDeleteChangeContext Then
        Dim helperCursor As Object
        Set helperCursor = getCursor().getText.createTextCursorByRange(textCursor)

        Do While Not (helperCursor.isEndOfWord() Or helperCursor.isStartOfParagraph())
            If Not helperCursor.goLeft(1, shouldExpand) Then Exit Do
        Loop

        If helperCursor.isStartOfParagraph() Then
            Set textCursor = helperCursor
            textCursor.gotoRange(textCursor.getStart(), shouldExpand)
            If getSpecial() = "d" Then
                textCursor.goRight(1, shouldExpand)
            End If
        End If
    End If

    MoveWordBackward = True
    Exit Function

    ErrorHandler:
    MoveWordBackward = False
End Function

Sub gotoMode(sMode)
    Select Case sMode
        Case "NORMAL" :
            setMode("NORMAL")
            setMovementModifier("")
        Case "INSERT" :
            setMode("INSERT")
        Case "VISUAL" :
            setMode("VISUAL")
            Dim oTextCursor
            oTextCursor = getTextCursor()
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            thisComponent.getCurrentController.Select(oTextCursor)
        Case "VISUAL_LINE" :
            setMode("VISUAL_LINE")
            formatVisualBase()
        Case "CMD" :
            setMode("CMD")
            setMovementModifier("")
            setSpecial("")
    End Select
End Sub

Sub cursorReset(oTextCursor)
    oTextCursor.gotoRange(oTextCursor.getStart(), False)
    oTextCursor.goRight(1, False)
    oTextCursor.goLeft(1, True)
    thisComponent.getCurrentController.Select(oTextCursor)
End Sub

Function samePos(oPos1, oPos2)
    samePos = oPos1.X() = oPos2.X() And oPos1.Y() = oPos2.Y()
End Function

Function genString(sChar, iLen)
    Dim sResult, i
    sResult = ""
    For i = 1 To iLen
        sResult = sResult & sChar
    Next i
    genString = sResult
End Function

' Get the current column position (character offset from start of line).
' Used by dd/cc to preserve horizontal position after deleting a line.
Function GetCursorColumn() As Integer
    Dim oVC, oText, oSaved, oTest

    oVC = ThisComponent.CurrentController.getViewCursor()
    oText = ThisComponent.Text

    ' Save exact current position
    oSaved = oText.createTextCursorByRange(oVC.getStart())

    ' Work on a duplicate model cursor
    oTest = oText.createTextCursorByRange(oSaved)

    ' Move duplicate to visual line start using a temporary ViewCursor
    Dim oTempVC
    oTempVC = ThisComponent.CurrentController.getViewCursor()
    oTempVC.gotoRange(oSaved.getStart(), False)
    oTempVC.gotoStartOfLine(False)

    ' Now select from visual start to original position using model cursor
    oTest.gotoRange(oTempVC, False)
    oTest.gotoRange(oSaved, True)

    GetCursorColumn = Len(oTest.getString())

    ' Restore original cursor position explicitly
    oVC.gotoRange(oSaved, False)
End Function


' Move the cursor to a specific column on the current line, or to the
' end of the line if the line is shorter than the requested column.
Sub SetCursorColumn(col As Integer)
    Dim oVC, oTest As Object
    Dim i, maxCol As Integer

    oVC = getCursor()
    If col <= 0 Then
        oVC.gotoStartOfLine(False)
        Exit Sub
    End If

    oVC.gotoStartOfLine(False)
    oVC.gotoEndOfLine(True)

    maxCol = Len(oVC.getString())
    oVC.gotoStartOfLine(False)
    If col > maxCol Then col = maxCol

    ' Move right up to desired column
    oVC.goRight(col, False)
End Sub

Function GetLineTextCursor(Optional bIncludeBreak As Boolean) As Object
    ' Returns a text cursor that selects the current line.
    ' If bIncludeBreak is True (default), includes the paragraph break.
    Dim oVC, oText, oStartCursor, oEndCursor, oSavedPos
    Set oVC = getCursor()
    Set oText = oVC.getText()

    If IsMissing(bIncludeBreak) Then bIncludeBreak = True

    ' Save original cursor position
    Set oSavedPos = oVC.getStart()

    ' Move to start of line
    oVC.gotoStartOfLine(False)
    Set oStartCursor = oText.createTextCursorByRange(oVC.getStart())

    ' Move to end of line
    oVC.gotoEndOfLine(False)
    If bIncludeBreak Then
        ' Try to include the newline character (if any)
        oVC.goRight(1, False)
    End If
    Set oEndCursor = oText.createTextCursorByRange(oVC.getStart())

    ' Restore original cursor position
    oVC.gotoRange(oSavedPos, False)

    ' Create a cursor spanning from start to end
    Dim oLineCursor
    Set oLineCursor = oText.createTextCursorByRange(oStartCursor.getStart())
    oLineCursor.gotoRange(oEndCursor.getStart(), True)

    Set GetLineTextCursor = oLineCursor
End Function

' -----------
' Selection Change Listener
' -----------
Sub SelectionChange_selectionChanged(oEvent)
     On Error Goto ErrorHandler
    If Not VIBREOFFICE_ENABLED Then Exit Sub
    Dim currentPage As Integer
    currentPage = getPageNum()
    If currentPage <> LAST_PAGE Then
        LAST_PAGE = currentPage
        setStatus(getMultiplier())
    End If
    Exit Sub
    ErrorHandler:
    ' Ignore
End Sub

Sub SelectionChange_disposing(oEvent)
    ' Required by XEventListener interface; do nothing.
End Sub

' -----------------------------------
' Special Mode (for chained commands)
' -----------------------------------
Sub setSpecial(specialName)
    SPECIAL_MODE = specialName

    If specialName = "" Then
        SPECIAL_COUNT = 0
    Else
        SPECIAL_COUNT = 2
    End If
End Sub

Function getSpecial()
    getSpecial = SPECIAL_MODE
End Function

Function getPageNum()
    getPageNum = getCursor().getPage()
End Function

Function getPageCount()
    getPageCount = thisComponent.CurrentController.PageCount
End Function

Function getWordCount()
    getWordCount = thisComponent.WordCount
End Function

Function getParagraphCount()
    getParagraphCount = thisComponent.ParagraphCount
End Function

Function getCurrentFileName()
    getCurrentFileName = thisComponent.getTitle()
End Function

Sub delaySpecialReset()
    SPECIAL_COUNT = SPECIAL_COUNT + 1
End Sub

Sub resetSpecial(Optional bForce)
    If IsMissing(bForce) Then bForce = False

    SPECIAL_COUNT = SPECIAL_COUNT - 1
    If SPECIAL_COUNT <= 0 Or bForce Then
        setSpecial("")
    End If
End Sub

' -----------------
' Movement Modifier
' -----------------
'f,i,a
Sub setMovementModifier(modifierName)
    MOVEMENT_MODIFIER = modifierName
End Sub

Function getMovementModifier()
    getMovementModifier = MOVEMENT_MODIFIER
End Function


' --------------------
' Multiplier functions
' --------------------
Sub _setMultiplier(n as integer)
        MULTIPLIER = n
End Sub

Sub resetMultiplier()
     _setMultiplier(0)
    End Sub

Sub addToMultiplier(n as integer)
    Dim sMultiplierStr as String
    ' Max multiplier: 10000 (stop accepting additions after 1000)
    If MULTIPLIER <= 1000 Then
        sMultiplierStr = CStr(MULTIPLIER) & CStr(n)
         _setMultiplier(CInt(sMultiplierStr))
        End If
End Sub

' Should only be used if you need the raw value
Function getRawMultiplier()
    getRawMultiplier = MULTIPLIER
End Function

' Same as getRawMultiplier, but defaults to 1 if it is unset (0)
Function getMultiplier()
    If MULTIPLIER = 0 Then
        getMultiplier = 1
    Else
        getMultiplier = MULTIPLIER
    End If
End Function
