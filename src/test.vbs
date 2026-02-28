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
    dim oTextCursor  as Object
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
	setRawStatus( _
		MODE & _
		" | Page: " & getPageNum() & "/" & getTotalPages() & _
		" | Word Count: " & getWordcount() & _
		" | " & statusText & _
		" | special: " & getSpecial() & _
		" | modifier: " & getMovementModifier() _
	)

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
			resetSpecial()

            dim oTextCursor
            oTextCursor = getTextCursor()
            ' Deselect TextCursor
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            ' Show TextCursor selection
            thisComponent.getCurrentController.Select(oTextCursor)
		Case "FORMAT":           ' <-- ADD THIS
            setMode("FORMAT")
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

Function getTotalPages()
    getTotalPages = thisComponent.CurrentController.PageCount
End Function
Function getWordcount()
    getWordcount = ThisComponent.DocumentProperties.DocumentStatistics(5).value

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
    dim dispatcher as object
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Copy", "", 0, Array())

    If bDelete Then
        getTextCursor().setString("")
    End If
End Sub


Sub pasteSelection()
    dim oTextCursor, dispatcher as object

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

' For the "u" (until) modifier, we need to capture two symbols.
' This tracks which keypress we're on (0=none, 1=waiting for second symbol)
global UNMATCHED_STATE as integer
global UNTIL_FIRST_SYMBOL as integer  ' ASCII code of first symbol

Sub setMovementModifier(modifierName)
    ' Special case: if we're transitioning from "i" or "a" to "u",
    ' concatenate to preserve the i/a information (becomes "iu" or "au")
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
    dim sMultiplierStr as string
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
    dim oTextCursor
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
' Used by dd/cc to preserve horizontal position after deleting a line.
Function GetCursorColumn() as integer
	dim oVC, oText, oSaved, oTest as object

    oVC = ThisComponent.CurrentController.getViewCursor()
    oText = ThisComponent.Text

    ' Save exact current position
    oSaved = oText.createTextCursorByRange(oVC.getStart())

    ' Work on a duplicate model cursor
    oTest = oText.createTextCursorByRange(oSaved)

    ' Move duplicate to visual line start using a temporary ViewCursor
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
Sub SetCursorColumn(col as integer)
	dim oVC, oTest as object
    dim i, maxCol as integer

    oVC = getCursor()
    If col <= 0 Then 
    	oVC.gotoStartOfLine(False)
    	Exit Sub
    End If
    ' Go to start of line
    oVC.gotoStartOfLine(False)
    oVC.gotoEndOfLine(True)
    maxCol = Len(oVC.getString())
    
    ' Reset back to start
    oVC.gotoStartOfLine(False)
    If col > maxCol then col = maxCol

    oVC.goRight(col, False)
End Sub

Function GetSymbol(symbol as string, modifier as string) as boolean
    dim endSymbol as string
    Select Case symbol
        Case "(", ")"
            symbol = "(" : endSymbol = ")"
        Case "{", "}"
            symbol = "{" : endSymbol = "}"
        Case "[", "]"
            symbol = "[" : endSymbol = "]"
        Case "<", ">"
            symbol = "<" : endSymbol = ">"
        Case "."
            symbol = "." : endSymbol = "."
        Case ","
            symbol = "," : endSymbol = ","
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

Function FindMatchingPair(startChar as string, endChar as string, modifier as string) as boolean
    dim oDoc, oCursor, oTempCursor, forwardPos as object
    dim i, j as integer
    dim foundForward as boolean

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
            dim startPos as object
            dim endPos as object
            dim oEndCursor as object

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
Function HandleUnMatchedPairs() as boolean
	If getMovementModifier() = "iu" or getMovementModifier() = "au" Then
	   If UNMATCHED_STATE = 0 Then
			' First keypress after 'u': save the start symbol, wait for end symbol
			UNTIL_FIRST_SYMBOL = keyChar
			UNMATCHED_STATE = 1
			HandleUnMatchedPairs = True
		ElseIf UNMATCHED_STATE = 1 Then
			' Second keypress: we have both symbols, call FindMatchingPair
			' Extract the "i" or "a" from the modifier string
			dim innerOrAround as string
			innerOrAround = Left(MOVEMENT_MODIFIER, 1)  ' "i" or "a"
			
			' Reset state after consuming both symbols
			UNMATCHED_STATE = 0
			UNTIL_FIRST_SYMBOL = 0
			HandleUnMatchedPairs = FindMatchingPair(Chr(UNTIL_FIRST_SYMBOL), Chr(keyChar), innerOrAround)
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
'   Always syncs getCursor() (the view cursor) to wherever the movement
'   landed.  When selecting=True, also commits the new selection to the
'   controller so LibreOffice renders the highlight.
' ========================================================================
Function HandleMovements(keyChar As Integer, selecting As Boolean) As Boolean

    Dim oTC   As Object   ' model text cursor — used for character/word moves
    Dim oldPos As Object
    Dim newPos As Object
    Dim i      As Integer

    oTC = getTextCursor()

    ' Assume we handle the key; set to False in the Case Else.
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
            getCursor().ScreenDown(False)

        Case 21     ' Ctrl+u — scroll up half a screen
            getCursor().ScreenUp(False)

        Case 119, 87    ' w / W — forward to start of next word
            For i = 1 To getMultiplier()
                oTC.gotoNextWord(selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 101    ' e — forward to end of current or next word
            For i = 1 To getMultiplier()
                oTC.goRight(1, selecting)
                oTC.gotoEndOfWord(selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 98, 66     ' b / B — backward to start of previous word
            For i = 1 To getMultiplier()
                oTC.gotoPreviousWord(selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 94     ' ^ — jump to first non-blank character of line
            ' Use the view cursor directly; gotoStartOfLine is not on model cursors.
            getCursor().gotoStartOfLine(selecting)

        Case 36     ' $ — jump to last character of line
            oldPos = getCursor().getPosition()
            getCursor().gotoEndOfLine(selecting)
            newPos = getCursor().getPosition()
            ' gotoEndOfLine can wrap to the start of the NEXT line on some
            ' paragraph ends; step back one position if that happened.
            If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
                getCursor().goLeft(1, selecting)
            End If

        Case 48     ' 0 — jump to absolute start of line (column 0)
			If getRawMultiplier = 0 then
				getCursor().gotoStartOfLine(selecting)
			End If
        Case 71     ' G — jump to end of document
            oTC.gotoEnd(selecting)
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 123    ' { — jump to start of previous paragraph
            For i = 1 To getMultiplier()
                oTC.gotoPreviousParagraph(selecting)
                oTC.goRight(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 125    ' } — jump to start of next paragraph
            For i = 1 To getMultiplier()
                oTC.gotoNextParagraph(selecting)
                oTC.goLeft(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 40     ' ( — jump to start of previous sentence
            For i = 1 To getMultiplier()
                oTC.gotoPreviousSentence(selecting)
                oTC.goLeft(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case 41     ' ) — jump to start of next sentence
            For i = 1 To getMultiplier()
                oTC.gotoNextSentence(selecting)
                oTC.goLeft(1, selecting)
            Next i
            getCursor().gotoRange(oTC.getStart(), selecting)

        Case Else
		    Select Case oEvent.KeyCode
				Case 1024 : getCursor().goDown(1, selecting)           ' ↓
				Case 1025 : getCursor().goUp(1, selecting)             ' ↑
				Case 1026 : getCursor().goLeft(1, selecting)           ' ←
				Case 1027 : getCursor().goRight(1, selecting)          ' →
				Case 1028 : getCursor().gotoStartOfLine(selecting)     ' Home => ^
				Case 1029                                          ' End  => $
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

    ' When selecting, commit the updated selection to the controller so
    ' LibreOffice renders the highlight correctly.
    If selecting Then
        ThisComponent.getCurrentController().Select(getTextCursor())
    End If

End Function
Function KeyHandler_KeyPressed(oEvent) as boolean
    dim oTextCursor
	dim consumeInput, IsMultiplier as boolean
	dim keyChar as integer

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
    ElseIf getSpecial() = "r" Then
		dim iLen
		iLen = Len(getCursor().getString())
		getCursor().setString(Chr(oEvent.KeyChar))
		resetSpecial()
		cursorReset(oTextCursor)
		Exit Function
	ElseIf HandleUnMatchedPairs() then
		KeyHandler_KeyPressed = true
		Exit Function
	EndIf
	consumeInput = True
	IsMultiplier = False
	keyChar = oEvent.KeyChar
	Select Case MODE
		Case "INSERT"
			
			resetMultiplier()
			If oEvent.KeyCode = 1281 Then
			
				resetSpecial(True)
				gotoMode("NORMAL")
				KeyHandler_KeyPressed = True
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
				iMult = getMultiplier()
				For i = 1 To iMult
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
						If keyChar = 103 Then   ' gg → go to start of document
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
									savedCol = GetCursorColumn()
									getCursor().gotoStartOfLine(False)
									Set oTCstart = ThisComponent.getText().createTextCursorByRange(getCursor().getStart())
									getCursor().gotoEndOfLine(False)
									getCursor().goRight(1, False)           ' include line break
									oTCstart.gotoRange(getCursor().getEnd(), True)
									ThisComponent.getCurrentController().Select(oTCstart)
									yankSelection(True)
									SetCursorColumn(savedCol)
									gotoMode("NORMAL")
									resetSpecial(True)
									GoTo NormalDone

								Case 99     ' c — cc: change whole line
									If getSpecial() = "d" Then GoTo Fallout  ' dc does nothing
									getCursor().gotoStartOfLine(False)
									Set oTCstart = ThisComponent.getText().createTextCursorByRange(getCursor().getStart())
									getCursor().gotoEndOfLine(False)
									getCursor().goRight(1, False)
									oTCstart.gotoRange(getCursor().getEnd(), True)
									ThisComponent.getCurrentController().Select(oTCstart)
									yankSelection(True)
									gotoMode("INSERT")
									resetSpecial(True)
									GoTo NormalDone

								Case Else   ' any other key: treat as motion
									If HandleMovements(keyChar, True) Then
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
						If HandleMovements(keyChar, True) Then
							yankSelection(False)
							gotoMode("NORMAL")
						ElseIf keyChar = 121 Then   ' yy — yank whole line
							getCursor().gotoStartOfLine(False)
							Set oTCstart = ThisComponent.getText().createTextCursorByRange(getCursor().getStart())
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
						iLen = Len(getCursor().getString())
						getCursor().setString(Chr(oEvent.KeyChar))
						resetSpecial()
						cursorReset(oTextCursor)
						resetSpecial(True)
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
			
			If HandleMovements(keyChar, False) Then
				GoTo NormalDone
			End If
			
			Select Case keyChar
			
				Case 48 ' 0
					If getRawMultiplier <> 0 then
						addToMultiplier(0)
						IsMultiplier = True
					End If
					
				Case 49,50,51,52,53,54,55,56,57 ' 1-9
					addToMultiplier(key - 48)
					IsMultiplier = True
					
				Case 99 'c => change
					setSpecial("c")
					delaySpecialReset()
					
				Case 100    ' d – delete: enter VISUAL, set special="d", await motion
					setSpecial("d")
					delaySpecialReset()
					
				Case 103    ' g — first half of gg
					setSpecial("g") 
					delaySpecialReset()
				Case 121 ' y => yank
					setSpecial("y")
					delaySpecialReset()
				Case 114    ' r — enter replace-char mode; next key replaces
					setSpecial("r")
					delaySpecialReset()
				Case 118    ' v - visual mode
					gotoMode("VISUAL")
				Case 105    ' i — INSERT at cursor position
					gotoMode("INSERT")

				Case 97     ' a — APPEND: INSERT after the cursor character
					getCursor().goRight(1, False)
					gotoMode("INSERT")

				Case 73     ' I — INSERT at start of the current line
					getCursor().gotoStartOfLine(False)
					gotoMode("INSERT")

				Case 65     ' A — APPEND at end of the current line
					getCursor().gotoEndOfLine(False)
					gotoMode("INSERT")

				Case 111    ' o — open a new line BELOW the cursor, enter INSERT
					getCursor().gotoEndOfLine(False)
					getCursor().goRight(1, False)
					getCursor().setString(Chr(13))
					If Not getCursor().isAtStartOfLine() Then
						getCursor().setString(Chr(13) & Chr(13))
						getCursor().goRight(1, False)
					End If
					gotoMode("INSERT")

				Case 79     ' O — open a new line ABOVE the cursor, enter INSERT
					getCursor().gotoStartOfLine(False)
					getCursor().setString(Chr(13))
					If Not getCursor().isAtStartOfLine() Then
						getCursor().goLeft(1, False)
						getCursor().setString(Chr(13))
						getCursor().goRight(1, False)
					End If
					gotoMode("INSERT")
					
				Case 47     ' / — focus the find bar
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					dsp.executeDispatch(ThisComponent.CurrentController.Frame, _
						"vnd.sun.star.findbar:FocusToFindbar", "", 0, Array())

				Case 92     ' \ — open find-and-replace dialog
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					dsp.executeDispatch(ThisComponent.CurrentController.Frame, _
						".uno:SearchDialog", "", 0, Array())
				
				Case 68     ' D — delete from cursor to end of line
					oTC.gotoRange(oTC.getStart(), False)
					ThisComponent.getCurrentController().Select(oTC)
					oldPos = getCursor().getPosition()
					getCursor().gotoEndOfLine(True)
					newPos = getCursor().getPosition()
					If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
						getCursor().goLeft(1, True)
					End If
					Set oTCop = ThisComponent.getText().createTextCursorByRange(getCursor())
					ThisComponent.getCurrentController().Select(oTCop)
					yankSelection(True)
					gotoMode("NORMAL")

				Case 67     ' C — change from cursor to end of line
					oTC.gotoRange(oTC.getStart(), False)
					ThisComponent.getCurrentController().Select(oTC)
					oldPos = getCursor().getPosition()
					getCursor().gotoEndOfLine(True)
					newPos = getCursor().getPosition()
					If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
						getCursor().goLeft(1, True)
					End If
					Set oTCop = ThisComponent.getText().createTextCursorByRange(getCursor())
					ThisComponent.getCurrentController().Select(oTCop)
					yankSelection(True)
					gotoMode("INSERT")

				Case 83     ' S — substitute entire line (change whole line)
					getCursor().gotoStartOfLine(False)
					Set oTCop = ThisComponent.getText().createTextCursorByRange(getCursor().getStart())
					getCursor().gotoEndOfLine(False)
					oTCop.gotoRange(getCursor().getEnd(), True)
					ThisComponent.getCurrentController().Select(oTCop)
					yankSelection(True)
					gotoMode("INSERT")

				Case 120    ' x — delete character under cursor
					For i = 1 To iMult
						oTC = getTextCursor()
						ThisComponent.getCurrentController().Select(oTC)
						yankSelection(True)
					Next i
					oTC = getTextCursor()
					If Not (oTC Is Nothing) Then cursorReset(oTC)
				
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
			If HandleMovements(keyChar, False) Then
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
				If keyChar = 103 Then   ' gg → go to start of document
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
					FindChar(oTC, sModFTV, Chr(keyChar), True)
				Next i
				getCursor().gotoRange(oTC.getStart(), False)
				ThisComponent.getCurrentController().Select(oTC)
				setMovementModifier("")
				GoTo VisualDone
			End If
			If getMovementModifier() <> "" Then ' movement modifier must be i or a in current implementation
				GetSymbol(keyChar, getMovementModifier())
				setMovementModifier("")
				Goto VisualDone
			
			End If
			If HandleMovements(keyChar, True) Then
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
				' Case 68     ' D => is supposed to delete selection and current line, feels unnecessary
				' Case 67     ' C => is supposed to delete selection and current line, feels unnecessary
				' Case 83     ' S —=> is supposed to delete selection and current line (I think), feels unnecessary
				' Case 114    r => It is supposed to replace all selected characters... idk If that's necessary
				
				Case 47     ' / — focus the find bar
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					dsp.executeDispatch(ThisComponent.CurrentController.Frame, _
						"vnd.sun.star.findbar:FocusToFindbar", "", 0, Array())

				Case 92     ' \ — open find-and-replace dialog
					dsp = createUnoService("com.sun.star.frame.DispatchHelper")
					dsp.executeDispatch(ThisComponent.CurrentController.Frame, _
						".uno:SearchDialog", "", 0, Array())
				
				Case 120    ' x — just deletes selection and goes to normal mode
					For i = 1 To iMult
						oTC = getTextCursor()
						ThisComponent.getCurrentController().Select(oTC)
						yankSelection(True)
					Next i
					oTC = getTextCursor()
					If Not (oTC Is Nothing) Then cursorReset(oTC)
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
				oTC = getTextCursor()
				If Not (oTC Is Nothing) Then cursorReset(oTC)
			End If
			KeyHandler_KeyPressed = consumeInput
			Exit Function
	End Select

	KeyHandler_KeyPressed = consumeInput
End Function
Function KeyHandler_KeyReleased(oEvent) as boolean
    ' Always consume the KeyReleased event in NORMAL, FORMAT, and VISUAL
    ' mode. Returning False lets the release propagate to LibreOffice, 
    ' which re-processes it as input — causing every command to fire
    ' twice. INSERT mode passes releases through so the application can
    ' handle auto-complete, IME, etc. normally.
    KeyHandler_KeyReleased = (MODE <> "INSERT")
End Function