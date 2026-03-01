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
global VIBREOFFICE_STARTED as boolean ' Defaults To False
global VIBREOFFICE_ENABLED as boolean ' Defaults To False

global oXKeyHandler as object

' Global State
global MODE as string
global VIEW_CURSOR as object
global MULTIPLIER as integer
global VISUAL_BASE as object 
global LAST_PAGE as integer
global oSelectionListener as object
' -----------
' Singletons
' -----------
Function getCursor
    getCursor = VIEW_CURSOR
End Function

Function getTextCursor
    Dim oTextCursor  as Object
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
    Dim oLayout
    oLayout = thisComponent.getCurrentController.getFrame.LayoutManager
    oLayout.destroyElement("private:resource/statusbar/statusbar")
    oLayout.createElement("private:resource/statusbar/statusbar")
End Sub

Sub setRawStatus(rawText)
    thisComponent.Currentcontroller.StatusIndicator.Start(rawText, 0)
End Sub

Sub setStatus(statusText)
	    setRawStatus( _
		MODE & _
		" | " & statusText & _
		" | special: " & getSpecial() & _
		" | " & "modifier: " & getMovementModifier() & _
		" | Page: " & getPageNum() & "/" & getPageCount() & _
		" | Words: " & getWordCount() & _ 
		" | Paragraphs: " & getParagraphCount() & _
		" | " & getCurrentFileName() _
		)

End Sub

Sub setMode(modeName)
    MODE = modeName
    setRawStatus(modeName)
End Sub

Function gotoMode(sMode)
    Select Case sMode
        Case "NORMAL"
            setMode("NORMAL")
            setMovementModifier("")
        Case "INSERT"
            setMode("INSERT")
        Case "VISUAL"
            setMode("VISUAL")
			resetSpecial()

            Dim oTextCursor
            oTextCursor = getTextCursor()
            ' Deselect TextCursor
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            ' Show TextCursor selection
            thisComponent.getCurrentController.Select(oTextCursor)
		Case "VISUAL LINE"
			setMode("VISUAL LINE")
		
			Dim oVC
            oVC = getTextCursor()
			
			oVC.gotoStartOfLine(False)
            oVC.gotoEndOfLine(True)
			
			 thisComponent.getCurrentController.Select(oTextCursor)
		Case "FORMAT"
            setMode("FORMAT")
		Case "COMMAND"
			setMode("COMMAND")
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

Function getPageNum()
    getPageNum = getCursor().getPage()
End Function

Function getPageCount()
    getPageCount = thisComponent.CurrentController.PageCount
End Function
Function getWordCount()
    getWordCount = ThisComponent.DocumentProperties.DocumentStatistics(5).value
End Function
Function getParagraphCount()
	getParagraphCount = ThisComponent.ParagraphCount
End Function

Function getCurrentFileName()
    getCurrentFileName = thisComponent.getTitle()
End Function

Function genString(sChar, iLen)
    Dim sResult, i
    sResult = ""
    For i = 1 To iLen
        sResult = sResult & sChar
    Next i
    genString = sResult
End Function

' Yanks selection To system clipboard.
' If bDelete is true, will delete selection.
Sub yankSelection(bDelete)
    Dim dispatcher as object
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Copy", "", 0, Array())

    If bDelete Then
        getTextCursor().setString("")
    End If
End Sub


Sub pasteSelection()
    Dim oTextCursor, dispatcher as object

    ' Deselect if in NORMAL mode To avoid overwriting the character underneath
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
global SPECIAL_MODE as string
global SPECIAL_COUNT as integer

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
'f,i,a,u (can be combined: "iu" or "au")
global MOVEMENT_MODIFIER as string

' For the "u" (until) modifier, we need To capture two symbols.
' This tracks which keypress we're on (0=none, 1=waiting for second symbol)
global UNMATCHED_STATE as integer
global UNTIL_FIRST_SYMBOL as integer  ' ASCII code of first symbol

Sub setMovementModifier(modifierName)
    ' Special case: if we're transitioning from "i" or "a" To "u",
    ' concatenate To preserve the i/a information (becomes "iu" or "au")
    If modifierName = "u" And (MOVEMENT_MODIFIER = "i" Or MOVEMENT_MODIFIER = "a") Then
        MOVEMENT_MODIFIER = MOVEMENT_MODIFIER & "u"
    Else
        MOVEMENT_MODIFIER = modifierName
    End If
    
    ' Reset "until" state when changing modifiers away from "iu"/"au"
    If MOVEMENT_MODIFIER <> "iu" And MOVEMENT_MODIFIER <> "au" Then
        UNMATCHED_STATE = 0
        UNTIL_FIRST_SYMBOL = 0
    End If
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
    Dim sMultiplierStr as string
    Dim iMultiplierInt as integer

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

' Same as getRawMultiplier, but defaults To 1 if it is unset (0)
Function getMultiplier()
    If MULTIPLIER = 0 Then
        getMultiplier = 1
    Else
        getMultiplier = MULTIPLIER
    End If
End Function
'Visual Line mode selector function'
Function formatVisualBase()
    dim oTextCursor
    oTextCursor = getTextCursor()
    VISUAL_BASE = getCursor().getPosition()

    ' Select the current line by moving cursor To start of the line below and
    ' then back To the start of the current line.
    getCursor().gotoEndOfLine(False)
    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
        getCursor().goRight(1, False)
    End If
    getCursor().goLeft(1, True)
    getCursor().gotoStartOfLine(True)
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
Sub XKeyHandler_Disposing(oEvent)
End Sub
Sub SelectionChange_disposing(oEvent)
    ' Required by XEventListener interface; do nothing.
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

Sub initVibreoffice
    Dim oTextCursor
    ' Initializing
    VIBREOFFICE_STARTED = True
    VIEW_CURSOR = thisComponent.getCurrentController.getViewCursor

    resetMultiplier()
    gotoMode("NORMAL")

    ' Show terminal cursor
    oTextCursor = getTextCursor()

    If oTextCursor Then
        cursorReset(oTextCursor)
    End If

    sStartXKeyHandler()
	LAST_PAGE = getPageNum()
	oSelectionListener = CreateUnoListener("SelectionChange_", "com.sun.star.view.XSelectionChangeListener")
	ThisComponent.CurrentController.addSelectionChangeListener(oSelectionListener)
End Sub

Sub UndoRedo(bUndo)
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

' Get the current column position (character offset from start of line).
' Used by dd/cc To preserve horizontal position after deleting a line.
Function GetCursorColumn() as integer
	Dim oVC, oText, oSaved, oTest, oTempVC as object

    oVC = ThisComponent.CurrentController.getViewCursor()
    oText = ThisComponent.Text

    ' Save exact current position
    oSaved = oText.createTextCursorByRange(oVC.getStart())

    ' Work on a duplicate model cursor
    oTest = oText.createTextCursorByRange(oSaved)

    ' Move duplicate To visual line start using a temporary ViewCursor
    oTempVC = ThisComponent.CurrentController.getViewCursor()
    oTempVC.gotoRange(oSaved.getStart(), False)
    oTempVC.gotoStartOfLine(False)

    ' Now select from visual start To original position using model cursor
    oTest.gotoRange(oTempVC, False)
    oTest.gotoRange(oSaved, True)

    GetCursorColumn = Len(oTest.getString())

    ' Restore original cursor position explicitly
    oVC.gotoRange(oSaved, False)
End Function


' Move the cursor To a specific column on the current line, or To the
' end of the line if the line is shorter than the requested column.
Sub SetCursorColumn(col as integer)
	Dim oVC, oTest as object
    Dim i, maxCol as integer

    oVC = getCursor()
    If col <= 0 Then 
    	oVC.gotoStartOfLine(False)
    	Exit Sub
    End If
    ' Go To start of line
    oVC.gotoStartOfLine(False)
    oVC.gotoEndOfLine(True)
    maxCol = Len(oVC.getString())
    
    ' Reset back To start
    oVC.gotoStartOfLine(False)
    If col > maxCol then col = maxCol

    oVC.goRight(col, False)
End Sub

Function GetSymbol(symbol as Integer, modifier as String) as boolean
    Dim endSymbol as String
    Select Case symbol
        Case 40,41 ' (, )
            symbol = "(" : endSymbol = ")"
        Case 123, 125 ' {, }
            symbol = "{" : endSymbol = "}"
        Case 91, 93 ' [, ]
            symbol = "[" : endSymbol = "]"
        Case 60, 62 ' <, >
            symbol = "<" : endSymbol = ">"
        Case 46 ' . 
            symbol = "." : endSymbol = "."
        Case 44 ' ,
            symbol = "," : endSymbol = ","
        Case "'":
            symbol = "‘" : endSymbol = "’"
            GetSymbol = FindMatchingPair(symbol, endSymbol, modifier)
            If Not GetSymbol Then
            	GetSymbol = FindMatchingPair("'", "'", modifier)
            End If
			Exit Function
        Case Chr(34):
            symbol = "“" : endSymbol = "”"
            GetSymbol = FindMatchingPair(symbol, endSymbol, modifier)
            If Not GetSymbol Then
            	GetSymbol = FindMatchingPair(Chr(34), Chr(34), modifier)
            End If
			Exit Function
        Case Else
            GetSymbol = False
            Exit Function
    End Select
    GetSymbol = FindMatchingPair(symbol, endSymbol, modifier)
End Function

Function FindMatchingPair(startChar as string, endChar as string, modifier as string) as boolean
    Dim oDoc, oCursor, oTempCursor, forwardPos as object
    Dim i, j as integer
    Dim foundForward as boolean

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
            Dim startPos as object
            Dim endPos as object
            Dim oEndCursor as object

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

            ' Move the VIEW cursor To startPos, then select To endPos
            ' getTextCursor() derives from getCursor(), so this is what yankSelection needs
            getCursor().gotoRange(startPos, False)
            getCursor().gotoRange(endPos, True)

            FindMatchingPair = True
            Exit Function
        End If
    Next j

    FindMatchingPair = False
End Function
Function HandleUnMatchedPairs(keyChar as Integer) as boolean
	If getMovementModifier() = "iu" or getMovementModifier() = "au" Then
	   If UNMATCHED_STATE = 0 Then
			' First keypress after 'u': save the start symbol, wait for end symbol
			UNTIL_FIRST_SYMBOL = keyChar
			UNMATCHED_STATE = 1
			HandleUnMatchedPairs = True
		ElseIf UNMATCHED_STATE = 1 Then
			' Second keypress: we have both symbols, call FindMatchingPair
			' Extract the "i" or "a" from the modifier string
			Dim innerOrAround as string
			innerOrAround = Left(MOVEMENT_MODIFIER, 1)  ' "i" or "a"
			UNMATCHED_STATE = 2
			' Reset state after consuming both symbols
			
			HandleUnMatchedPairs = FindMatchingPair(Chr(UNTIL_FIRST_SYMBOL), Chr(keyChar), innerOrAround)
			UNTIL_FIRST_SYMBOL = 0
		End If
	Else
		HandleUnMatchedPairs = False
	End If
End Function
' Returns:
'   True  if the key was a recognised movement key and was handled
'   False if the key was not a movement key (caller should handle it)
'
' Side effects:
'   Always syncs getCursor() (the view cursor) To wherever the movement
'   landed.  When selecting=True, also commits the new selection To the
'   controller so LibreOffice renders the highlight.
' ========================================================================
Function HandleMovements(keyChar As Integer, keyCode As Integer, selecting As Boolean) As Boolean

    Dim oTC   As Object   ' model text cursor — used for character/word moves
    Dim oldPos As Object
    Dim newPos As Object
    Dim i      As Integer

    oTC = getTextCursor()

    ' Assume we handle the key; set To False in the Case Else.
    HandleMovements = True

    Select Case keyChar

        Case 104    ' h — left one character per multiplier count
            For i = 1 To getMultiplier()
                oTC.goLeft(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 108    ' l — right one character per multiplier count
            For i = 1 To getMultiplier()
                oTC.goRight(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 106    ' j — down one line per multiplier count
            ' Line navigation lives on the view cursor, not the model cursor.
            For i = 1 To getMultiplier()
                getCursor().goDown(1, selecting)
            Next i

        Case 107    ' k — up one line per multiplier count
            For i = 1 To getMultiplier()
                getCursor().goUp(1, selecting)
            Next i

        Case 4      ' Ctrl+d — scroll down half a screen
            getCursor().ScreenDown(selecting)

        Case 21     ' Ctrl+u — scroll up half a screen
            getCursor().ScreenUp(selecting)
			
		Case 1030 ' Page Up
			dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
			For i = 1 To getMultiplier()
				dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:PageUpSel", "", 0, Array())
			Next i
		Case 1031 ' Page Down
			dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
			For i = 1 To getMultiplier()
				dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:PageDownSel", "", 0, Array())
			Next i

        Case 119, 87    ' w / W — forward To start of next word
            For i = 1 To getMultiplier()
                oTC.gotoNextWord(selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 101    ' e — forward To end of current or next word
            For i = 1 To getMultiplier()
                oTC.goRight(1, selecting)
                oTC.gotoEndOfWord(selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 98, 66     ' b / B — backward To start of previous word
            For i = 1 To getMultiplier()
				If oTC.isStartOfParagraph() Then
					oTC.goLeft(1, bExpand)
				End If
                oTC.gotoPreviousWord(selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 94     ' ^ — jump To first non-blank character of line
            ' Use the view cursor directly; gotoStartOfLine is not on model cursors.
            getCursor().gotoStartOfLine(selecting)

        Case 36     ' $ — jump To last character of line
            oldPos = getCursor().getPosition()
            getCursor().gotoEndOfLine(selecting)
            newPos = getCursor().getPosition()
            ' gotoEndOfLine can wrap To the start of the NEXT line on some
            ' paragraph ends; step back one position if that happened.
            If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
                getCursor().goLeft(1, selecting)
            End If

        Case 48     ' 0 — jump To absolute start of line (column 0)
			If getRawMultiplier = 0 then
				getCursor().gotoStartOfLine(selecting)
			End If
        Case 71     ' G — jump To end of document
            oTC.gotoEnd(selecting)
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 123    ' { — jump To start of previous paragraph
            For i = 1 To getMultiplier()
                oTC.gotoPreviousParagraph(selecting)
                oTC.goRight(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 125    ' } — jump To start of next paragraph
            For i = 1 To getMultiplier()
                oTC.gotoNextParagraph(selecting)
                oTC.goLeft(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 40     ' ( — jump To start of previous sentence
            For i = 1 To getMultiplier()
                oTC.gotoPreviousSentence(selecting)
                oTC.goLeft(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 41     ' ) — jump To start of next sentence
            For i = 1 To getMultiplier()
                oTC.gotoNextSentence(selecting)
                oTC.goLeft(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case Else
		    Select Case keyCode
				Case 1024 ' ↓
					For i = 1 To getMultiplier()
						getCursor().goDown(1, selecting) 
					Next i
				Case 1025 ' ↑
					For i = 1 To getMultiplier()
						getCursor().goUp(1, selecting)
					Next i
				Case 1026 ' ←
					For i = 1 To getMultiplier()
						getCursor().goLeft(1, selecting)           
					Next i
				Case 1027 ' →
					For i = 1 To getMultiplier()
						getCursor().goRight(1, selecting)          
					Next i
				Case 1028 ' Home => ^
					getCursor().gotoStartOfLine(selecting)     
				Case 1029  ' End  => $
					oldPos = getCursor().getPosition()
					getCursor().gotoEndOfLine(selecting)
					newPos = getCursor().getPosition()
					If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
						getCursor().goLeft(1, selecting)
					End If
				Case Else 
					HandleMovements = False     ' not a movement key; caller handles it
            Exit Function
			End Select
    End Select

    ' When selecting, commit the updated selection To the controller so
    ' LibreOffice renders the highlight correctly.
    If selecting Then
        ThisComponent.getCurrentController().Select(getTextCursor())
    End If

End Function
Function KeyHandler_KeyPressed(oEvent) as boolean
    Dim oTextCursor
	Dim consumeInput, IsMultiplier as boolean
	Dim keyChar, i as integer
	keyChar = oEvent.KeyChar
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
	ElseIf HandleUnMatchedPairs(keyChar) then
		KeyHandler_KeyPressed = true
			If UNMATCHED_STATE = 2 Then
			yankSelection(True)
			If getSpecial() = "c" Then
				gotoMode("INSERT")
			End If
			setMovementModifier("")
			resetSpecial(True)
			UNMATCHED_STATE = 0
			End If
		Exit Function
	EndIf
	consumeInput = True
	IsMultiplier = False
	
	KeyHandler_KeyPressed = true
	Select Case MODE
		Case "INSERT"
			
			resetMultiplier()
			If oEvent.KeyCode = 1281 Then
			
				resetSpecial(True)
				gotoMode("NORMAL")
				cursorReset(oTextCursor)
				Exit Function
			Else	
				resetSpecial()
				oTextCursor.gotoRange(oTextCursor.getStart(), False)
				thisComponent.getCurrentController.Select(oTextCursor)
				KeyHandler_KeyPressed = False 'don't consume the input
				Exit Function
			End If
		Case "NORMAL"
			If getMovementModifier() = "f" Or getMovementModifier() = "t" Or _
			   getMovementModifier() = "F" Or getMovementModifier() = "T" Then
				Dim bExpFTN As Boolean
				bExpFTN = (getSpecial() = "d" Or getSpecial() = "c" Or getSpecial() = "y")
				For i = 1 To getMultiplier()
					ProcessSearchKey(oTextCursor, getMovementModifier(), Chr(keyChar), bExpFTN)
				Next i
				getCursor().gotoRange(oTextCursor.getStart(), False)
				If bExpFTN Then
					ThisComponent.getCurrentController().Select(oTextCursor)
					yankSelection(getSpecial() <> "y")
				End If
				If getSpecial() = "c" Then 
					gotoMode("INSERT") 
				Else 
					gotoMode("NORMAL")
				End If
				
				setMovementModifier("") : resetSpecial(True)
				GoTo NormalDone
			End If
			' ── 'g' special pending: second key must be 'g' for gg ─────────
			If getSpecial() <> "" Then
				Select Case getSpecial()
					Case "g"
						If keyChar = 103 Then   ' gg → go To start of document
							getCursor().gotoStart(False)
						End If
						resetSpecial(True)
						GoTo NormalDone

					Case "d", "c"
						If getMovementModifier() <> "" Then
							If getMovementModifier() = "i" Or getMovementModifier() = "a" Then
								If keyChar = 117 Then   ' 'u' chains i→iu or a→au
									setMovementModifier("u")
									UNMATCHED_STATE = 0
									delaySpecialReset()
									GoTo NormalDone
								End If
								If GetSymbol(Chr(keyChar), getMovementModifier()) Then
									yankSelection(getSpecial() <> "y")
									If getSpecial() = "c" Then gotoMode("INSERT") Else gotoMode("NORMAL")
									resetSpecial(True)
									GoTo NormalDone
								End If
							Else
								' iu / au two-symbol capture is handled by the global guard;
								' if we somehow land here with another modifier just cancel.
								GoTo Fallout
							End If
							GoTo Fallout
						Else
							Select Case keyChar
								Case 105    ' i — (c/d)i: set inside modifier, wait for symbol
									setMovementModifier("i")
									delaySpecialReset()
									GoTo NormalDone

								Case 97     ' a — (c/d)a: set around modifier, wait for symbol
									setMovementModifier("a")
									delaySpecialReset()
									GoTo NormalDone

								Case 100    ' d — dd: delete whole line
									If getSpecial() = "c" Then GoTo Fallout  ' cd does nothing
									Dim savedCol
									savedCol = GetCursorColumn()
									getCursor().gotoStartOfLine(False)
									oTextCursor = ThisComponent.getText().createTextCursorByRange(getCursor().getStart())
									getCursor().gotoEndOfLine(False)
									oTextCursor.gotoRange(getCursor().getEnd(), True)
									ThisComponent.getCurrentController().Select(oTextCursor)
									yankSelection(True)
									SetCursorColumn(savedCol)
									gotoMode("NORMAL")
									resetSpecial(True)
									GoTo NormalDone

								Case 99     ' c — cc: change whole line
									If getSpecial() = "d" Then GoTo Fallout  ' dc does nothing
									getCursor().gotoStartOfLine(False)
									Set oTextCursor = ThisComponent.getText().createTextCursorByRange(getCursor().getStart())
									getCursor().gotoEndOfLine(False)
									oTextCursor.gotoRange(getCursor().getEnd(), True)
									ThisComponent.getCurrentController().Select(oTextCursor)
									yankSelection(True)
									gotoMode("INSERT")
									resetSpecial(True)
									GoTo NormalDone

								Case Else   ' any other key: treat as motion
									If HandleMovements(keyChar, oEvent.KeyCode, True) Then
										yankSelection(getSpecial() <> "y")
										If getSpecial() = "c" Then gotoMode("INSERT") Else gotoMode("NORMAL")
										resetSpecial(True)
									Else
										GoTo Fallout
									End If
							End Select
						End If
						Fallout:
						resetSpecial(True)
						GoTo NormalDone

					Case "y"
						If HandleMovements(keyChar, oEvent.KeyCode, True) Then
							yankSelection(False)
							gotoMode("NORMAL")
						ElseIf keyChar = 121 Then   ' yy — yank whole line
							getCursor().gotoStartOfLine(False)
							Dim oTCstart
							oTCstart = ThisComponent.getText().createTextCursorByRange(getCursor().getStart())
							getCursor().gotoEndOfLine(False)
							getCursor().goRight(1, False)
							oTCstart.gotoRange(getCursor().getEnd(), True)
							ThisComponent.getCurrentController().Select(oTCstart)
							yankSelection(False)
							gotoMode("NORMAL")
						End If
						resetSpecial(True)
						GoTo NormalDone

					Case "r" 
						getCursor().setString(Chr(oEvent.KeyChar))
						resetSpecial(True)
						cursorReset(oTextCursor)
						KeyHandler_KeyPressed = True
						Exit Function

				End Select
			End If
			If oEvent.KeyCode = 1281 Then
			
				resetSpecial(True)
				setMovementModifier("")
				KeyHandler_KeyPressed = True
				cursorReset(oTextCursor)
				Exit Function
			End If
			
			If HandleMovements(keyChar, oEvent.KeyCode, False) Then
				GoTo NormalDone
			End If
			
			Select Case keyChar
			
				Case 48 ' 0
					If getRawMultiplier <> 0 then
						addToMultiplier(0)
						IsMultiplier = True
					End If
					
				Case 49,50,51,52,53,54,55,56,57 ' 1-9
					addToMultiplier(keyChar - 48)
					IsMultiplier = True
					
				Case 99 'c => change
					setSpecial("c")
					delaySpecialReset()
					
				Case 100    ' d – delete: enter VISUAL, set special="d", await motion
					setSpecial("d")
					delaySpecialReset()
					
				Case 103    ' g —  gg or page jump
					If getRawMultiplier() > 0 Then 
						Dim targetPage As Integer
						Dim itotalPages As Integer

						targetPage = getRawMultiplier()
						itotalPages = getPageCount()
						If targetPage > itotalPages Then targetPage = itotalPages

						getCursor().jumpToPage(targetPage, False)
						oTextCursor.gotoRange(getCursor().getStart(), False)
					Else
						setSpecial("g") ' gg
						delaySpecialReset()
					End If
				Case 121 ' y => yank
					setSpecial("y")
					delaySpecialReset()
				Case 114    ' r — enter replace-char mode; next key replaces
					setSpecial("r")
					delaySpecialReset()
				Case 118    ' v - visual mode
					gotoMode("VISUAL")
				Case 86: ' 86='V'
					gotoMode("VISUAL LINE")
					
				Case 105    ' i — INSERT at cursor position
					gotoMode("INSERT")
					Exit Function


				Case 97     ' a — APPEND: INSERT after the cursor character
					getCursor().goRight(1, False)
					gotoMode("INSERT")
					Exit Function

				Case 73     ' I — INSERT at start of the current line
					getCursor().gotoPreviousParagraph(False)
					getCursor().goRight(1, False)
					gotoMode("INSERT")
					Exit Function

				Case 65     ' A — APPEND at end of the current line
					oTextCursor.gotoNextParagraph(False)
					oTextCursor.goLeft(1, False)
					gotoMode("INSERT")
					Exit Function

				Case 111    ' o — open a new line BELOW the cursor, enter INSERT
					getCursor().gotoEndOfLine(False)
					getCursor().goRight(1, False)
					getCursor().setString(Chr(13))
					If Not getCursor().isAtStartOfLine() Then
						getCursor().setString(Chr(13) & Chr(13))
						getCursor().goRight(1, False)
					End If
					gotoMode("INSERT")
					Exit Function

				Case 79     ' O — open a new line ABOVE the cursor, enter INSERT
					getCursor().gotoStartOfLine(False)
					getCursor().setString(Chr(13))
					If Not getCursor().isAtStartOfLine() Then
						getCursor().goLeft(1, False)
						getCursor().setString(Chr(13))
						getCursor().goRight(1, False)
					End If
					gotoMode("INSERT")
					Exit Function
					
				Case 47     ' / — focus the find bar
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					dsp.executeDispatch(ThisComponent.CurrentController.Frame, _
						"vnd.sun.star.findbar:FocusToFindbar", "", 0, Array())

				Case 92     ' \ — open find-and-replace dialog
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					dsp.executeDispatch(ThisComponent.CurrentController.Frame, _
						".uno:SearchDialog", "", 0, Array())
				
				Case 68     ' D — delete from cursor To end of line
					oTextCursor.gotoRange(oTextCursor.getStart(), False)
					ThisComponent.getCurrentController().Select(oTextCursor)
					oldPos = getCursor().getPosition()
					getCursor().gotoEndOfLine(True)
					newPos = getCursor().getPosition()
					If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
						getCursor().goLeft(1, True)
					End If
					oTextCursor = ThisComponent.getText().createTextCursorByRange(getCursor())
					ThisComponent.getCurrentController().Select(oTextCursor)
					yankSelection(True)
					gotoMode("NORMAL")

				Case 67     ' C — change from cursor To end of line
					oTextCursor.gotoRange(oTextCursor.getStart(), False)
					ThisComponent.getCurrentController().Select(oTextCursor)
					oldPos = getCursor().getPosition()
					getCursor().gotoEndOfLine(True)
					newPos = getCursor().getPosition()
					If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
						getCursor().goLeft(1, True)
					End If
					oTextCursor = ThisComponent.getText().createTextCursorByRange(getCursor())
					ThisComponent.getCurrentController().Select(oTextCursor)
					yankSelection(True)
					gotoMode("INSERT")
					Exit Function

				Case 83     ' S — substitute entire line (change whole line)
					getCursor().gotoStartOfLine(False)
					oTextCursor = ThisComponent.getText().createTextCursorByRange(getCursor().getStart())
					getCursor().gotoEndOfLine(False)
					oTextCursor.gotoRange(getCursor().getEnd(), True)
					ThisComponent.getCurrentController().Select(oTextCursor)
					yankSelection(True)
					gotoMode("INSERT")
					Exit Function

				Case 120    ' x — delete character under cursor
					For i = 1 To getMultiplier()
						oTextCursor = getTextCursor()
						ThisComponent.getCurrentController().Select(oTextCursor)
						yankSelection(True)
					Next i
					oTextCursor = getTextCursor()
					If Not (oTextCursor Is Nothing) Then cursorReset(oTextCursor)
				
				Case 112    ' p — paste AFTER cursor (move right first)
					getCursor().goRight(1, False)
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					For i = 1 To getMultiplier()
						dsp.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Paste", "", 0, Array())
					Next i

				Case 80     ' P — paste BEFORE cursor (no movement)
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					For i = 1 To getMultiplier()
						dsp.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Paste", "", 0, Array())
					Next i
				
				Case 117, 18 ' u, C-r => Undo, Redo
					UndoRedo(keyChar = 117)
					
				Case 102    ' f — find char FORWARD, cursor lands ON it
					setMovementModifier("f")
					delaySpecialReset()

				Case 116    ' t — till char FORWARD, cursor lands BEFORE it
					setMovementModifier("t")
					delaySpecialReset()

				Case 70     ' F — find char BACKWARD, cursor lands ON it
					setMovementModifier("F")
					delaySpecialReset()

				Case 84     ' T — till char BACKWARD, cursor lands AFTER it
					setMovementModifier("T")
					delaySpecialReset()
									
				Case Else
					consumeInput = False
			End Select
			NormalDone:
			resetSpecial()
			If Not IsMultiplier And getSpecial() = "" And getMovementModifier() = "" And UNMATCHED_STATE = 0 Then
				resetMultiplier()
			End If
			setStatus(getMultiplier())
			oTextCursor = getTextCursor()
			If Not (oTextCursor Is Nothing) Then cursorReset(oTextCursor)
			KeyHandler_KeyPressed = consumeInput
			Exit Function
		Case "FORMAT"
			

			Select Case keyChar:
			
			
			
			End Select
			If HandleMovements(keyChar, oEvent.KeyCode, False) Then
				GoTo FormatDone
			End If
			FormatDone:
			
			
			
		Case "VISUAL"
		
			If oEvent.KeyCode = 1281 Then

				resetSpecial(True)
				gotoMode("NORMAL")
				KeyHandler_KeyPressed = True
				cursorReset(oTextCursor)
				Exit Function
			End If
			If getSpecial () <> "" Then
				' for now only g triggers special and only gg is implemented
				If keyChar = 103 Then   ' gg → go To start of document
					getCursor().gotoStart(False)
				End If
						
				resetSpecial(True)  ' unknown two-key sequence; cancel
				GoTo VisualDone
			End If
			
			' ── Active f/t/F/T modifier: this keypress IS the target char ────
			If getMovementModifier() = "f" Or getMovementModifier() = "t" Or _
			   getMovementModifier() = "F" Or getMovementModifier() = "T" Then
				Dim sModFTV As String
				sModFTV = getMovementModifier()
				For i = 1 To getMultiplier()
					FindChar(oTextCursor, sModFTV, Chr(keyChar), True)
				Next i
				getCursor().gotoRange(oTextCursor.getStart(), False)
				ThisComponent.getCurrentController().Select(oTextCursor)
				setMovementModifier("")
				GoTo VisualDone
			End If
			If getMovementModifier() <> "" Then ' movement modifier must be i or a in current implementation
				GetSymbol(keyChar, getMovementModifier())
				setMovementModifier("")
				Goto VisualDone
			
			End If
			If HandleMovements(keyChar, oEvent.KeyCode, True) Then
				GoTo VisualDone
			End If
			Select Case keyChar:
			
				Case 48 ' 0
					If getRawMultiplier <> 0 then
						addToMultiplier(0)
						IsMultiplier = True
					End If
					
				Case 49,50,51,52,53,54,55,56,57 ' 1-9
					addToMultiplier(key - 48)
					IsMultiplier = True
					
				Case 99 'c => change
					yankSelection(True)
					gotoMode("INSERT")
					
				Case 100    ' d => delete: enter VISUAL, set special="d", await motion
					yankSelection(True)
					gotoMode("NORMAL")
					
				Case 103    ' g => first half of gg
					setSpecial("g") 
					delaySpecialReset()
				Case 121 ' y => yank
					yankSelection(False)

				Case 118    ' v - visual mode
					gotoMode("NORMAL")

				Case 73     ' I => INSERT at start of the current line
					getCursor().gotoStartOfLine(False)
					gotoMode("INSERT")

				Case 65     ' A => APPEND at end of the current line
					getCursor().gotoEndOfLine(False)
					gotoMode("INSERT")
				
				Case 105 ' i => inside
					setMovementModifier("i")
				Case 97 ' a => around
					setMovementModifier("a")
				' Case 111    '  o => It hops around the selection???
				' Case 79     ' O => It hops around the selection???
				' Case 68     ' D => is supposed To delete selection and current line, feels unnecessary
				' Case 67     ' C => is supposed To delete selection and current line, feels unnecessary
				' Case 83     ' S —=> is supposed To delete selection and current line (I think), feels unnecessary
				' Case 114    r => It is supposed To replace all selected characters... idk If that's necessary
				
				Case 47     ' / — focus the find bar
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					dsp.executeDispatch(ThisComponent.CurrentController.Frame, _
						"vnd.sun.star.findbar:FocusToFindbar", "", 0, Array())

				Case 92     ' \ — open find-and-replace dialog
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					dsp.executeDispatch(ThisComponent.CurrentController.Frame, _
						".uno:SearchDialog", "", 0, Array())
				
				Case 120    ' x — just deletes selection and goes To normal mode
					For i = 1 To getMultiplier()
						oTextCursor = getTextCursor()
						ThisComponent.getCurrentController().Select(oTextCursor)
						yankSelection(True)
					Next i
					oTextCursor = getTextCursor()
					If Not (oTextCursor Is Nothing) Then cursorReset(oTextCursor)
					gotoMode("Normal")
				
				Case 112    ' p — paste AFTER cursor (move right first)
					getCursor().goRight(1, False)
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					For i = 1 To getMultiplier()
						dsp.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Paste", "", 0, Array())
					Next i

				Case 80     ' P — paste BEFORE cursor (no movement)
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					For i = 1 To getMultiplier()
						dsp.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Paste", "", 0, Array())
					Next i
				
				Case 117, 18 ' u, C-r => Undo, Redo
					UndoRedo(keyChar = 117)
					
				Case 102    ' f — find char FORWARD, cursor lands ON it
					setMovementModifier("f")
					delaySpecialReset()

				Case 116    ' t — till char FORWARD, cursor lands BEFORE it
					setMovementModifier("t")
					delaySpecialReset()

				Case 70     ' F — find char BACKWARD, cursor lands ON it
					setMovementModifier("F")
					delaySpecialReset()

				Case 84     ' T — till char BACKWARD, cursor lands AFTER it
					setMovementModifier("T")
					delaySpecialReset()
				Case Else
					consumeInput = False
			End Select
			
			VisualDone:
			resetSpecial()
			If getSpecial() = "" And getMovementModifier() = "" And UNMATCHED_STATE = 0 Then
				resetMultiplier()
			End If
			setStatus(getMultiplier())
			' In VISUAL we leave the selection visible; cursorReset is for NORMAL only.
			If MODE = "NORMAL" Then
				If Not (getTextCursor() Is Nothing) Then cursorReset(oTextCursor)
			End If
			KeyHandler_KeyPressed = consumeInput
			Exit Function
		Case "Command"
			If oEvent.KeyCode = 1281 Then

				resetSpecial(True)
				gotoMode("NORMAL")
				KeyHandler_KeyPressed = True
				cursorReset(oTextCursor)
				Exit Function
			End If
			
		Case "VISUAL"
			If oEvent.KeyCode = 1281 Then

				resetSpecial(True)
				gotoMode("NORMAL")
				KeyHandler_KeyPressed = True
				cursorReset(oTextCursor)
				Exit Function
			End If
			
	End Select

	KeyHandler_KeyPressed = consumeInput
End Function
Function KeyHandler_KeyReleased(oEvent) as boolean
    ' Always consume the KeyReleased event in NORMAL, FORMAT, and VISUAL
    ' mode. Returning False lets the release propagate To LibreOffice, 
    ' which re-processes it as input — causing every command To fire
    ' twice. INSERT mode passes releases through so the application can
    ' handle auto-complete, IME, etc. normally.
    KeyHandler_KeyReleased = (MODE <> "INSERT")
End Function