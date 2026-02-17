
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

' Global State
global MODE as string
global VIEW_CURSOR as object
global MULTIPLIER as integer

' -----------
' Singletons
' -----------
Function getCursor
    getCursor = VIEW_CURSOR
End Function

Function getTextCursor
    dim oTextCursor
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
    dim oLayout
    oLayout = thisComponent.getCurrentController.getFrame.LayoutManager
    oLayout.destroyElement("private:resource/statusbar/statusbar")
    oLayout.createElement("private:resource/statusbar/statusbar")
End Sub

Sub setRawStatus(rawText)
    thisComponent.Currentcontroller.StatusIndicator.Start(rawText, 0)
End Sub

Sub setStatus(statusText)
    setRawStatus(MODE & " | " & statusText & " | special: " & getSpecial() & " | " & "modifier: " & getMovementModifier())
End Sub

Sub setMode(modeName)
    MODE = modeName
    setRawStatus(modeName)
End Sub

Function gotoMode(sMode)
    Select Case sMode
        Case "NORMAL":
            setMode("NORMAL")
            setMovementModifier("")
        Case "INSERT":
            setMode("INSERT")
        Case "VISUAL":
            setMode("VISUAL")

            dim oTextCursor
            oTextCursor = getTextCursor()
            ' Deselect TextCursor
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            ' Show TextCursor selection
            thisComponent.getCurrentController.Select(oTextCursor)
    End Select
End Function

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
    dim sResult, i
    sResult = ""
    For i = 1 To iLen
        sResult = sResult & sChar
    Next i
    genString = sResult
End Function

' Yanks selection to system clipboard.
' If bDelete is true, will delete selection.
Sub yankSelection(bDelete)
    dim dispatcher As Object
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Copy", "", 0, Array())

    If bDelete Then
        getTextCursor().setString("")
    End If
End Sub


Sub pasteSelection()
    dim oTextCursor, dispatcher As Object

    ' Deselect if in NORMAL mode to avoid overwriting the character underneath
    ' the cursor
    If MODE = "NORMAL" Then
        oTextCursor = getTextCursor()
        oTextCursor.gotoRange(oTextCursor.getStart(), False)
        thisComponent.getCurrentController.Select(oTextCursor)
    End If

    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Paste", "", 0, Array())
End Sub


' -----------------------------------
' Special Mode (for chained commands)
' -----------------------------------
global SPECIAL_MODE As string
global SPECIAL_COUNT As integer

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
global MOVEMENT_MODIFIER As string

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
    dim sMultiplierStr as String
    dim iMultiplierInt as integer

    ' Max multiplier: 10000 (stop accepting additions after 1000)
    If MULTIPLIER <= 1000 then
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
End Sub

Sub XKeyHandler_Disposing(oEvent)
End Sub


' --------------------
' Main Key Processing
' --------------------
function KeyHandler_KeyPressed(oEvent) as boolean
    dim oTextCursor

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

    dim bConsumeInput, bIsMultiplier, bIsModified, bIsSpecial
    bConsumeInput = True ' Block all inputs by default
    bIsMultiplier = False ' reset multiplier by default
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    bIsSpecial = getSpecial() <> ""


    ' --------------------------
    ' Process global shortcuts, exit if matched (like ESC)
    If ProcessGlobalKey(oEvent) Then
        ' Pass

    ' If INSERT mode, allow all inputs
    ElseIf MODE = "INSERT" Then
        bConsumeInput = False

    ' If Change Mode
    ' ElseIf MODE = "NORMAL" And Not bIsSpecial And getMovementModifier() = "" And ProcessModeKey(oEvent) Then
    ElseIf ProcessModeKey(oEvent) Then
        ' Pass

    ' Replace Key
    ElseIf getSpecial() = "r" And Not bIsModified Then
        dim iLen
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

    ' If standard movement key (in VISUAL mode) like arrow keys, home, end
    ElseIf MODE = "VISUAL" And ProcessStandardMovementKey(oEvent) Then
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
    If not bIsMultiplier and getSpecial() = "" and getMovementModifier() = "" Then
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
    dim bMatched, bIsControl
    bMatched = True
    bIsControl = (oEvent.Modifiers = 2) or (oEvent.Modifiers = 8)

    ' PRESSED ESCAPE (or ctrl+[)
    ' KeyCode 1281 = Escape, KeyCode 1315 = '[', ascii 91
    if oEvent.KeyCode = 1281 Or (oEvent.KeyCode = 1315 And bIsControl) Then
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


Function ProcessStandardMovementKey(oEvent)
    dim c, bMatched
    c = oEvent.KeyCode

    bMatched = True

    If MODE <> "VISUAL" Then
        bMatched = False
        'Pass
    ElseIf c = 1024 Then ' Down arrow
        ProcessMovementKey(106, True) ' 106='j'
    ElseIf c = 1025 Then ' Up arrow
        ProcessMovementKey(107, True) ' 107='k'
    ElseIf c = 1026 Then ' Left arrow
        ProcessMovementKey(104, True) ' 104='h'
    ElseIf c = 1027 Then ' Right arrow
        ProcessMovementKey(108, True) ' 108='l'
    ElseIf c = 1028 Then ' Home
        ProcessMovementKey(94, True)  ' 94='^'
    ElseIf c = 1029 Then ' End
        ProcessMovementKey(36, True)  ' 36='$'
    Else
        bMatched = False
    End If

    ProcessStandardMovementKey = bMatched
End Function


Function ProcessNumberKey(oEvent)
    dim key as Integer
    key = oEvent.KeyChar

    ' 48='0' through 57='9'
    If key >= 48 and key <= 57 Then
        addToMultiplier(key - 48)
        ProcessNumberKey = True
    Else
        ProcessNumberKey = False
    End If
End Function


Function ProcessModeKey(oEvent)
    dim bIsModified, key as Integer
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    ' Don't change modes in these circumstances
    If MODE <> "NORMAL" Or bIsModified Or getSpecial() <> "" Or getMovementModifier() <> "" Then
        ProcessModeKey = False
        Exit Function
    End If

    key = oEvent.KeyChar

    ' Mode matching
    dim bMatched
    bMatched = True
    Select Case key
        ' Insert modes
        ' 105='i', 97='a', 73='I', 65='A', 111='o', 79='O'
        Case 105, 97, 73, 65, 111, 79:
            If key = 97  Then getCursor().goRight(1, False) ' 'a': move right before insert
            If key = 73  Then ProcessMovementKey(94)        ' 'I': go to start of line
            If key = 65  Then ProcessMovementKey(36)        ' 'A': go to end of line

            If key = 111 Then ' 'o': open line below
                ProcessMovementKey(36)  ' '$'
                ProcessMovementKey(108) ' 'l'
                getCursor().setString(chr(13))
                If Not getCursor().isAtStartOfLine() Then
                    getCursor().setString(chr(13) & chr(13))
                    ProcessMovementKey(108) ' 'l'
                End If
            End If

            If key = 79 Then ' 'O': open line above
                ProcessMovementKey(94)  ' '^'
                getCursor().setString(chr(13))
                If Not getCursor().isAtStartOfLine() Then
                    ProcessMovementKey(104) ' 'h'
                    getCursor().setString(chr(13))
                    ProcessMovementKey(108) ' 'l'
                End If
            End If

            gotoMode("INSERT")
        Case 118: ' 118='v'
            gotoMode("VISUAL")
        Case Else:
            bMatched = False
    End Select
    ProcessModeKey = bMatched
End Function


Function ProcessNormalKey(keyChar, modifiers)
    dim i, bMatched, bIsVisual, iIterations, bIsControl
    bIsControl = (modifiers = 2) or (modifiers = 8)

    bIsVisual = (MODE = "VISUAL") ' is this hardcoding bad? what about visual block?

    ' ----------------------
    ' 1. Check Movement Key
    ' ----------------------
    iIterations = getMultiplier()
    bMatched = False
    For i = 1 To iIterations
        dim bMatchedMovement

        ' Movement Key
        bMatchedMovement = ProcessMovementKey(keyChar, bIsVisual, modifiers)
        bMatched = bMatched or bMatchedMovement

        ' If Special: d/c + movement
        If bMatched And (getSpecial() = "d" Or getSpecial() = "c" Or getSpecial() = "y") Then
            yankSelection((getSpecial() <> "y"))
        End If
    Next i

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
    ' 2. Undo/Redo
    ' --------------------
    ' 117='u', 114='r' (Ctrl+r = redo)
    If keyChar = 117 Or (bIsControl And keyChar = 114) Then ' 'u' or Ctrl+'r'
        For i = 1 To iIterations
            Undo(keyChar = 117) ' 117='u'
        Next i

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 3. Paste
    '   Note: in vim, paste will result in cursor being over the last character
    '   of the pasted content. Here, the cursor will be the next character
    '   after that. Fix?
    ' --------------------
    ' 112='p', 80='P'
    If keyChar = 112 or keyChar = 80 Then ' 'p' or 'P'
        ' Move cursor right if "p" to paste after cursor
        If keyChar = 112 Then ' 'p'
            ProcessMovementKey(108, False) ' 108='l'
        End If

        For i = 1 To iIterations
            pasteSelection()
        Next i

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
    If keyChar <> 120 and getSpecial() = "" Then ' 120='x'
        iIterations = 1
    End If
    For i = 1 To iIterations
        dim bMatchedSpecial

        ' Special/Delete Key
        bMatchedSpecial = ProcessSpecialKey(keyChar)

        bMatched = bMatched or bMatchedSpecial
    Next i


    ProcessNormalKey = bMatched
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


Function ProcessSpecialKey(keyChar)
    dim oTextCursor, bMatched, bIsSpecial, bIsDelete
    bMatched = True
    bIsSpecial = getSpecial() <> ""

    ' 100='d', 99='c', 115='s', 121='y'
    If keyChar = 100 Or keyChar = 99 Or keyChar = 115 Or keyChar = 121 Then ' 'd','c','s','y'
        bIsDelete = (keyChar <> 121) ' 121='y'

        ' Special Cases: 'dd' and 'cc'
        If bIsSpecial Then
            dim bIsSpecialCase
            ' 100='d', 99='c'
            bIsSpecialCase = (keyChar = 100 And getSpecial() = "d") Or (keyChar = 99 And getSpecial() = "c") ' 'dd' or 'cc'

            If bIsSpecialCase Then
                ProcessMovementKey(94, False)  ' 94='^'
                ProcessMovementKey(106, True)  ' 106='j'

                oTextCursor = getTextCursor()
                thisComponent.getCurrentController.Select(oTextCursor)
                yankSelection(bIsDelete)
            Else
                bMatched = False
            End If

            ' Go to INSERT mode after 'cc', otherwise NORMAL
            If bIsSpecialCase And keyChar = 99 Then ' 99='c'
                gotoMode("INSERT")
            Else
                gotoMode("NORMAL")
            End If


        ' visual mode: delete selection
        ElseIf MODE = "VISUAL" Then
            oTextCursor = getTextCursor()
            thisComponent.getCurrentController.Select(oTextCursor)

            yankSelection(bIsDelete)

            ' 99='c', 115='s', 100='d', 121='y'
            If keyChar = 99 Or keyChar = 115 Then gotoMode("INSERT") ' 'c' or 's'
            If keyChar = 100 Or keyChar = 121 Then gotoMode("NORMAL") ' 'd' or 'y'


        ' Enter Special mode: 'd', 'c', or 'y' ('s' => 'cl')
        ElseIf MODE = "NORMAL" Then

            ' 115='s' => 'cl'
            If keyChar = 115 Then ' 's'
                setSpecial("c")
                gotoMode("VISUAL")
                ProcessNormalKey(108, 0) ' 108='l'
            Else
                setSpecial(Chr(keyChar)) ' store as string for getSpecial() comparisons
                gotoMode("VISUAL")
            End If
        End If

    ' If is 'r' for replace: 114='r'
    ElseIf keyChar = 114 Then ' 'r'
        setSpecial("r")

    ' Otherwise, ignore if bIsSpecial
    ElseIf bIsSpecial Then
        bMatched = False

    ' 120='x'
    ElseIf keyChar = 120 Then ' 'x'
        oTextCursor = getTextCursor()
        thisComponent.getCurrentController.Select(oTextCursor)
        yankSelection(True)

        ' Reset Cursor
        cursorReset(oTextCursor)

        ' Goto NORMAL mode (in the case of VISUAL mode)
        gotoMode("NORMAL")

    ' 68='D', 67='C'
    ElseIf keyChar = 68 Or keyChar = 67 Then ' 'D' or 'C'
        If MODE = "VISUAL" Then
            ProcessMovementKey(94, False)  ' 94='^'
            ProcessMovementKey(36, True)   ' 36='$'
            ProcessMovementKey(108, True)  ' 108='l'
        Else
            ' Deselect
            oTextCursor = getTextCursor()
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            thisComponent.getCurrentController.Select(oTextCursor)
            ProcessMovementKey(36, True) ' 36='$'
        End If

        yankSelection(True)

        If keyChar = 68 Then     ' 'D'
            gotoMode("NORMAL")
        ElseIf keyChar = 67 Then ' 'C'
            gotoMode("INSERT")
        End IF

    ' 83='S', only valid in NORMAL mode
    ElseIf keyChar = 83 And MODE = "NORMAL" Then ' 'S'
        ProcessMovementKey(94, False) ' 94='^'
        ProcessMovementKey(36, True)  ' 36='$'
        yankSelection(True)
        gotoMode("INSERT")

    Else
        bMatched = False
    End If

    ProcessSpecialKey = bMatched
End Function


Function ProcessMovementModifierKey(keyChar)
    dim bMatched

    bMatched = True
    ' 102='f', 116='t', 70='F', 84='T', 105='i', 97='a'
    Select Case keyChar
        Case 102: setMovementModifier("f") ' 'f'
        Case 116: setMovementModifier("t") ' 't'
        Case 70:  setMovementModifier("F") ' 'F'
        Case 84:  setMovementModifier("T") ' 'T'
        Case 105: setMovementModifier("i") ' 'i'
        Case 97:  setMovementModifier("a") ' 'a'
        Case Else:
            bMatched = False
    End Select

    ProcessMovementModifierKey = bMatched
End Function


Function ProcessSearchKey(oTextCursor, searchType, keyChar, bExpand)
    '-----------
    ' Searching
    ' keyChar here is a string (the literal character to find), not an ascii int.
    ' It is passed in from ProcessMovementKey after converting with Chr().
    '-----------
    dim bMatched, oSearchDesc, oFoundRange, bIsBackwards, oStartRange
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

    oFoundRange = thisComponent.findNext( oStartRange, oSearchDesc )

    If not IsNull(oFoundRange) Then
        dim oText, foundPos, curPos, bSearching
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
        do until oText.compareRegionStarts(foundPos, curPos) = 0
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





' -----------------------
' Main Movement Function
' -----------------------
'   Default: bExpand = False, keyModifiers = 0
'   keyChar is an ASCII integer (e.g. 104 for 'h')
Function ProcessMovementKey(keyChar, Optional bExpand, Optional keyModifiers)
    dim oTextCursor, bSetCursor, bMatched
    oTextCursor = getTextCursor()
    bMatched = True
    If IsMissing(bExpand) Then bExpand = False
    If IsMissing(keyModifiers) Then keyModifiers = 0


    ' Check for modified keys (Ctrl, Alt, not Shift)
    If keyModifiers > 1 Then
        dim bIsControl
        bIsControl = (keyModifiers = 2) or (keyModifiers = 8)

        ' Ctrl+d (100) and Ctrl+u (117)
        If bIsControl and keyChar = 100 Then ' 100='d'
            getCursor().ScreenDown(bExpand)
        ElseIf bIsControl and keyChar = 117 Then ' 117='u'
            getCursor().ScreenUp(bExpand)
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
            Case "f", "t", "F", "T":
                bMatched = ProcessSearchKey(oTextCursor, getMovementModifier(), Chr(keyChar), bExpand)
            Case "a", "i":
                bMatched = GetSymbol(Chr(keyChar), getMovementModifier())
                bSetCursor = False
            Case Else:
                bSetCursor = False
                bMatched = False
        End Select

        If Not bMatched Then
            bSetCursor = False
        End If
    ' ---------------------------------

    ElseIf keyChar = 108 Then ' 108='l'
        oTextCursor.goRight(1, bExpand)

    ElseIf keyChar = 104 Then ' 104='h'
        oTextCursor.goLeft(1, bExpand)

    ' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
    ' So this is a weird hack
    ElseIf keyChar = 107 Then ' 107='k'
        getCursor().goUp(1, bExpand)
        bSetCursor = False

    ElseIf keyChar = 106 Then ' 106='j'
        getCursor().goDown(1, bExpand)
        bSetCursor = False
    ' ----------

    ElseIf keyChar = 94 Then ' 94='^'
        getCursor().gotoStartOfLine(bExpand)
        bSetCursor = False

    ElseIf keyChar = 36 Then ' 36='$'
        dim oldPos, newPos
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

    ElseIf keyChar = 119 Or keyChar = 87 Then ' 119='w', 87='W'
        oTextCursor.gotoNextWord(bExpand)
    ElseIf keyChar = 98 Or keyChar = 66 Then  ' 98='b', 66='B'
        oTextCursor.gotoPreviousWord(bExpand)
    ElseIf keyChar = 101 Then                  ' 101='e'
        oTextCursor.gotoEndOfWord(bExpand)

    ElseIf keyChar = 41 Then ' 41=')'
        oTextCursor.gotoNextSentence(bExpand)
    ElseIf keyChar = 40 Then ' 40='('
        oTextCursor.gotoPreviousSentence(bExpand)
    ElseIf keyChar = 125 Then ' 125='}'
        oTextCursor.gotoNextParagraph(bExpand)
    ElseIf keyChar = 123 Then ' 123='{'
        oTextCursor.gotoPreviousParagraph(bExpand)

    Else
        bSetCursor = False
        bMatched = False
    End If

    ' If oTextCursor was moved, set global cursor to its position
    If bSetCursor Then
        getCursor().gotoRange(oTextCursor.getStart(), False)
    End If

    ' If oTextCursor was moved and is in VISUAL mode, update selection
    if bSetCursor and bExpand then
        thisComponent.getCurrentController.Select(oTextCursor)
    end if

    ProcessMovementKey = bMatched
End Function

Function GetSymbol(symbol As String, modifier As String) As Boolean
    Dim endSymbol As String
    Select Case symbol
        Case "(", ")"
            symbol = "(" : endSymbol = ")"
        Case "{", "}"
            symbol = "{" : endSymbol = "}"
        Case "[", "]"
            symbol = "[" : endSymbol = "]"
        Case "<", ">"
            symbol = "<" : endSymbol = ">"
        Case "'":
            symbol = "‘" : endSymbol = "’"
            GetSymbol = FindMatchingPair(symbol, endSymbol, modifier)
            If GetSymbol = "" Then
            	GetSymbol = FindMatchingPair("'", "'", modifier)
            	Exit Function
            End If
        Case Chr(34):
            symbol = "“" : endSymbol = "”"
            GetSymbol = FindMatchingPair(symbol, endSymbol, modifier)
            If GetSymbol = "" Then
            	GetSymbol = FindMatchingPair(Chr(34), Chr(34), modifier)
            	Exit Function
            End If
        Case Else
            GetSymbol = False
            Exit Function
    End Select
    GetSymbol = FindMatchingPair(symbol, endSymbol, modifier)
End Function

Function FindMatchingPair(startChar As String, endChar As String, modifier As String) As Boolean
    Dim oDoc As Object
    Dim oCursor As Object
    Dim oTempCursor As Object
    Dim i As Integer, j As Integer
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
            Dim startPos As Object
            Dim endPos As Object
            Dim oEndCursor As Object

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
Sub initVibreoffice
    dim oTextCursor
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
End Sub


Sub Main
    If Not VIBREOFFICE_STARTED Then
        initVibreoffice()
    End If

    ' Toggle enable/disable
    VIBREOFFICE_ENABLED = Not VIBREOFFICE_ENABLED

    ' Restore statusbar
    If Not VIBREOFFICE_ENABLED Then restoreStatus()
End Sub