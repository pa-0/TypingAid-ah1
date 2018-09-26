;  TypingAid
;  http://www.autohotkey.com/board/topic/49517-ahk-11typingaid-v2200-word-autocompletion-utility/
;
;  Press 1 to 0 keys to autocomplete the word upon suggestion 
;  Or use the Up/Down keys to select an item
;  (0 will match suggestion 10) 
;                              Credits:
;                                -Maniac
;                                -Jordi S
;                                -hugov
;                                -kakarukeys
;                                -Asaptrad
;                                -j4hangir
;                                -Theclaw
;___________________________________________ 

; Press 1 to 0 keys to autocomplete the word upon suggestion 
;___________________________________________

;    CONFIGURATIONS 

#NoTrayIcon
;disable hotkeys until setup is complete
Suspend, On 
#NoEnv
ListLines Off

g_OSVersion := GetOSVersion()

;Set the Coordinate Modes before any threads can be executed
CoordMode, Caret, Screen
CoordMode, Mouse, Screen

EvaluateScriptPathAndTitle()

SuspendOn()
BuildTrayMenu()      

OnExit, SaveScript

;Change the setup performance speed
SetBatchLines, 20ms
;read in the preferences file
ReadPreferences()

SetTitleMatchMode, 2

; === MY SETTINGS ===
alexF_config_PreventScrollbar := true
alexF_config_OnAfterCompletion := "On also after completion"
alexF_config_OrderByLength := true
alexF_config_GroupSimilarWords := true ; This supersedes alexF_config_OrderByLength
alexF_config_MinDeltaLength := 2
alexF_config_Ellipsis := Chr(0x2026)
alexF_config_PurgeInterval := 3600 * 1000 ; One hour in milliseconds; 1 * 60 * 1000 ;1 min    

alexF_config_PreventScrollbar := true
alexF_config_BackupWordsAndCounts := true


;set windows constants
g_EVENT_SYSTEM_FOREGROUND := 0x0003
g_EVENT_SYSTEM_SCROLLINGSTART := 0x0012
g_EVENT_SYSTEM_SCROLLINGEND := 0x0013
g_GCLP_HCURSOR := -12
g_IDC_HAND := 32649
g_IDC_HELP := 32651
g_IMAGE_CURSOR := 2
g_LR_SHARED := 0x8000
g_NormalizationKD := 0x6
g_NULL := 0
g_Process_DPI_Unaware := 0
g_Process_System_DPI_Aware  := 1
g_Process_Per_Monitor_DPI_Aware := 2
g_PROCESS_QUERY_INFORMATION := 0x0400
g_PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
g_SB_VERT := 0x1
g_SIF_POS := 0x4
g_SM_CMONITORS := 80
g_SM_CXVSCROLL := 2
g_SM_CXFOCUSBORDER := 83
g_WINEVENT_SKIPOWNPROCESS := 0x0002
g_WM_LBUTTONUP := 0x202
g_WM_LBUTTONDBLCLK := 0x203
g_WM_MOUSEMOVE := 0x200
g_WM_SETCURSOR := 0x20

;setup code
g_DpiScalingFactor := A_ScreenDPI/96
g_Helper_Id = 
g_HelperManual = 
g_DelimiterChar := Chr(2)
g_cursor_hand := DllCall( "LoadImage", "Ptr", g_NULL, "Uint", g_IDC_HAND , "Uint", g_IMAGE_CURSOR, "int", g_NULL, "int", g_NULL, "Uint", g_LR_SHARED ) 
if (A_PtrSize == 8) {
   g_SetClassLongFunction := "SetClassLongPtr"
} else {
   g_SetClassLongFunction := "SetClassLong"
}
g_PID := DllCall("GetCurrentProcessId")

; dllPath := "C:\PROJECTS\GitHub\TypingAid\TAHelper\Win32\Debug\TAHelperU32.dll"
  dllPath :=           "D:\GitHub\TypingAid\TAHelper\x64\Debug\TAHelperU64.dll"
g_AlexF_Dll := DllCall("LoadLibrary", "WStr", dllPath) ; , "Ptr")
; MsgBox % "Loaded " . dllPath . ". Handle: " . g_AlexF_Dll . ". ErrorLevel: " . ErrorLevel

; foo := 3
; foo1 := DllCall("TAHelperU64.dll\Add5", "Int", foo, "Int")
; MsgBox % "Calculated " . foo . " Add5 = " . foo1 . ". ErrorLevel: " . ErrorLevel



AutoTrim, Off 

InitializeListBox()

BlockInput, Send

InitializeHotKeys()

DisableKeyboardHotKeys()

;Change the Running performance speed (Priority changed to High in GetIncludedActiveWindow)
SetBatchLines, -1

;Read in the WordList
ReadWordList()

g_WinChangedCallback := RegisterCallback("WinChanged")
g_ListBoxScrollCallback := RegisterCallback("ListBoxScroll")

if !(g_WinChangedCallback)
{
   MsgBox, Failed to register callback function
   ExitApp
}

if !(g_ListBoxScrollCallback)
{
   MsgBox, Failed to register ListBox Scroll callback function
   ExitApp
}
   
;Find the ID of the window we are using
GetIncludedActiveWindow()

MainLoop()

; END

MainLoop()
{
   global g_TerminatingEndKeys
   global g_LearnedWordInsertionTime
   
   g_LearnedWordInsertionTime := 0 ; no insertions yet
   Loop 
   { 

      ;If the active window has changed, wait for a new one
      IF !( ReturnWinActive() ) 
      {
         Critical, Off
         GetIncludedActiveWindow()
      } else {    
         Critical, Off
      }
   
      ;Get one key at a time 
      Input, InputChar, L1 V I, {BS}%g_TerminatingEndKeys%
   
      Critical
      EndKey := ErrorLevel
   
      ProcessKey(InputChar,EndKey)
   }
}

; AlexF
;  Given pressed character or terminating key, takes appropriate action.
;  (I think that only one of the two is not empty)
ProcessKey(InputChar,EndKey)
{
   global g_Active_Id
   global g_Helper_Id
   global g_IgnoreSend
   global g_LastInput_Id
   global g_OldCaretX
   global g_OldCaretY
   global g_TerminatingCharactersParsed
   global g_Word  ; AlexF word typed by user
   global prefs_DetectMouseClickMove
   global prefs_EndWordCharacters
   global prefs_ForceNewWordCharacters
   global prefs_Length
   
   IfEqual, g_IgnoreSend, 1
   {
      g_IgnoreSend = 
      Return
   }

   IfEqual, EndKey,
   {
      EndKey = Max
   }
   
   IfEqual, EndKey, NewInput  ; AlexF: redundant? Typo? NewInput is not defined...
      Return

   IfEqual, EndKey, Endkey:Tab
      If ( GetKeyState("Alt") =1 || GetKeyState("LWin") =1 || GetKeyState("RWin") =1 )
         Return
   
   ;If we have no window activated for typing, we don't want to do anything with the typed character
   IfEqual, g_Active_Id,
   {
      if (!GetIncludedActiveWindow())
      {
         Return
      }
   }


   IF !( ReturnWinActive() )
   {
      if (!GetIncludedActiveWindow())
      {
         Return
      }
   }
   
   IfEqual, g_Active_Id, %g_Helper_Id%
   {
      Return
   }
   
   ;If we haven't typed anywhere, set this as the last window typed in
   IfEqual, g_LastInput_Id,
      g_LastInput_Id = %g_Active_Id%
   
   IfNotEqual, prefs_DetectMouseClickMove, On
   {
      ifequal, g_OldCaretY,
         g_OldCaretY := HCaretY()
         
      if ( g_OldCaretY != HCaretY() )
      {
         ;Don't do anything if we aren't in the original window and aren't starting a new word
         IfNotEqual, g_LastInput_Id, %g_Active_Id%
            Return
            
         ; add the word if switching lines
         AddWordToList(g_Word,0)
         ClearAllVars(true)
         g_Word := InputChar
         Return         
      } 
   }

   g_OldCaretY := HCaretY()
   g_OldCaretX := HCaretX()
   
   ;Backspace clears last letter 
   ifequal, EndKey, Endkey:BackSpace
   {
      ;Don't do anything if we aren't in the original window and aren't starting a new word
      IfNotEqual, g_LastInput_Id, %g_Active_Id%
         Return
      
      StringLen, len, g_Word
      IfEqual, len, 1   
      {
         ClearAllVars(true)
      } else IfNotEqual, len, 0
      {
         StringTrimRight, g_Word, g_Word, 1
      }
   } else if ( ( EndKey == "Max" ) && !(InStr(g_TerminatingCharactersParsed, InputChar)) )
   {
      ; If active window has different window ID from the last input,
      ;learn and blank word, then assign number pressed to the word
      IfNotEqual, g_LastInput_Id, %g_Active_Id%
      {
         AddWordToList(g_Word,0)
         ClearAllVars(true)
         g_Word := InputChar
         g_LastInput_Id := g_Active_Id
         Return
      }
   
      if InputChar in %prefs_ForceNewWordCharacters%
      {
         AddWordToList(g_Word,0)
         ClearAllVars(true)
         g_Word := InputChar
      ; } else if InputChar in %prefs_EndWordCharacters% ; AlexF decided not to support this option
      ; {
         ; g_Word .= InputChar
         ; AddWordToList(g_Word, 1)
         ; ClearAllVars(true)
      } else { 
         g_Word .= InputChar
      }
      
   } else IfNotEqual, g_LastInput_Id, %g_Active_Id%
   {
      ;Don't do anything if we aren't in the original window and aren't starting a new word
      Return
   } else {
      AddWordToList(g_Word,0)
      ClearAllVars(true)
      Return
   }
   
   ;Wait till minimum letters 
   IF ( StrLen(g_Word) < prefs_Length )
   {
      CloseListBox()
      Return
   }
   SetTimer, RecomputeMatchesTimer, -1
}

RecomputeMatchesTimer:
   Thread, NoTimers
   RecomputeMatches()
   Return

; Takes the given word (g_Word), recompiles the list of matches and redisplays the wordlist.
RecomputeMatches()
{
   global g_MatchTotal         ;AlexF count of matched words
   global g_SingleMatchDb      ;AlexF array of matched words, as they are stored in the database. Used for deleting only.
   global g_SingleMatchAdj     ;AlexF array of matched words, with adjusted capitalization. This is what user sees. 
   global g_Word               ;AlexF word typed by user
   global g_WordListDB
   global prefs_ArrowKeyMethod
   global prefs_LearnMode
   global prefs_ListBoxRows
   global prefs_NoBackSpace
   global prefs_ShowLearnedFirst
   global prefs_SuppressMatchingWord
   global alexF_config_OrderByLength
   global alexF_config_PreventScrollbar
   global alexF_config_GroupSimilarWords
   global alexF_config_MinDeltaLength
   global alexF_config_PurgeInterval
   global g_LearnedWordInsertionTime ; AlexF, in milliseconds, 10 ms resolution
   
   ;_0. Maybe purge database of one-time entered words (possibly typos)
   if(g_LearnedWordInsertionTime) {
         elapsedTime := A_TickCount - g_LearnedWordInsertionTime
         
        if(elapsedTime > alexF_config_PurgeInterval) {
            ; MsgBox % "elapsedTime = " elapsedTime " msec; purge interval = " alexF_config_PurgeInterval
            g_WordListDB.Query("DELETE FROM words WHERE count = 1;")
            g_LearnedWordInsertionTime := 0 ; reset "timer"
         }
   }
   
   SavePriorMatchPosition()

   ;_1. Prepare for the search in DB
   ;Match part-word with command 
   g_MatchTotal = 0 
   
   IfEqual, prefs_ArrowKeyMethod, Off
   {
      IfLess, prefs_ListBoxRows, 10
         LimitTotalMatches := prefs_ListBoxRows
      else LimitTotalMatches = 10
   } else {
      if(alexF_config_GroupSimilarWords) {
         LimitTotalMatches = 200
      } else if (alexF_config_PreventScrollbar) {
         LimitTotalMatches := prefs_ListBoxRows
      } else {
         LimitTotalMatches = 200
      }
   }
   
   StringUpper, WordAllCaps, g_Word
   
   ; AlexF added: store word's capitalization to match it later
   targetCapitalization := GetCapitalization(g_Word)
   
   WordMatch := StrUnmark(WordAllCaps) ;AlexF - redundant?
   
   StringUpper, WordMatch, WordMatch
   
   ; if a user typed an accented character, we should exact match on that accented character
   if (WordMatch != WordAllCaps) {
      WordAccentQuery =
      LoopCount := StrLen(g_Word)
      Loop, %LoopCount%
      {
         Position := A_Index
         SubChar := SubStr(g_Word, Position, 1)
         SubCharNormalized := StrUnmark(SubChar)
         if !(SubCharNormalized == SubChar) {
            StringUpper, SubCharUpper, SubChar
            StringLower, SubCharLower, SubChar
            StringReplace, SubCharUpperEscaped, SubCharUpper, ', '', All
            StringReplace, SubCharLowerEscaped, SubCharLower, ', '', All
            PrefixChars =
            Loop, % Position - 1
            {
               PrefixChars .= "?"
            }
            ; because SQLite cannot do case-insensitivity on accented characters using LIKE, we need
            ; to handle it manually, so we need 2 searches for each accented character the user typed.
            ;GLOB is used for consistency with the wordindexed search.
            WordAccentQuery .= " AND (word GLOB '" . PrefixChars . SubCharUpperEscaped . "*' OR word GLOB '"
            WordAccentQuery .= PrefixChars . SubCharLowerEscaped . "*')"
         }         
      }
   } else {
      WordAccentQuery =
   }
   
   StringReplace, WordExactEscaped, g_Word, ', '', All
   StringReplace, WordMatchEscaped, WordMatch, ', '', All
   
   IfEqual, prefs_SuppressMatchingWord, On
   {
      IfEqual, prefs_NoBackSpace, Off
      {
         SuppressMatchingWordQuery := " AND word <> '" . WordExactEscaped . "'"
      } else {
               SuppressMatchingWordQuery := " AND wordindexed <> '" . WordMatchEscaped . "'"
            }
   }
   
   ;_2. First query - just find minimal frequency of matched words ('Normalize') 
   ; to retrieve more frequent and longer words first- AlexF removed 
   WhereQuery := " WHERE wordindexed GLOB '" . WordMatchEscaped . "*' " . SuppressMatchingWordQuery . WordAccentQuery
   
   ;_3. Second query - actually retrieve matches, in certain order
   if (alexF_config_GroupSimilarWords) {
      typedLength := StrLen(g_word)
      ;WhereQuery .= " AND LENGTH(word) > " . (typedLength - 1 + alexF_config_MinDeltaLength)  ;TODO reconcile with min word length? Think!!
      WhereQuery .= " AND LENGTH(word) > " . typedLength  
      ;AlexF, from previous OrderByQuery: (count - min) * (1 - 0.75/nExtraChars) -- advantage to more frequent and longer words.
      OrderByQuery := " ORDER BY word ASC" ; order alphabetically, case-sensitive.
  } else if (alexF_config_OrderByLength) { 
      OrderByQuery := " ORDER BY LENGTH(word)"
  } else {
      MsgBox % "Error. Missing setting to display matches."
      Return
  }


   ;AlexF 1-column table of matched words
   query := "SELECT word FROM Words" 
   query .= WhereQuery . OrderByQuery . " LIMIT " . LimitTotalMatches . ";"
   Matches := g_WordListDB.Query(query)
   
   words := []
   for each, row in Matches.Rows {
      words.Push(row[1])
   }
   
   if(alexF_config_GroupSimilarWords) {
      groupedWords := GroupMatches(words, StrLen(g_Word) + alexF_config_MinDeltaLength)
   } else {
      groupedWords := words
   }

   g_SingleMatchDb := Object() ; for deleting from the database
   g_SingleMatchAdj := Object()
   for each, truncWord in groupedWords {

      g_SingleMatchDb[++g_MatchTotal] := truncWord
      ; If truncWord has "normal" capitalization ("|firstCap|"), it will be adjusted to match word.
      adjWord := AdjustCapitalization(truncWord, targetCapitalization, g_Word)
      g_SingleMatchAdj[g_MatchTotal] := adjWord

      if(g_MatchTotal == prefs_ListBoxRows) {
         break
      }
   }
   
   ;If no match then clear Tip 
   IfEqual, g_MatchTotal, 0
   {
      ClearAllVars(false)
      Return 
   } 

   SetupMatchPosition() ; what position to highlight in the listbox
   RebuildMatchList() ; generate g_Match - concatenation of all the lines in the listbox
   ShowListBox()
}

;------------------------------------------------------------------------

; Given array of words, sorted alphabetically, and if the table
; is long enough, truncate each long word, add ellipsis and removed same words.
; truncLength - integer, minimal length of truncated words.
GroupMatches(words, truncLength) {
   global prefs_ListBoxRows
   global alexF_config_Ellipsis

   if(words.MaxIndex() <= prefs_ListBoxRows)
      Return words
   
   ;Try to truncate and group some of the words
   truncWords := []
   prevWord := "" ; previous word, possibly truncated and with ellipsis
   
   for each, w in words {
      L := StrLen(w)
      prevL := StrLen(prevWord)
      
      ; Don't add ellipsis to short words
      if(L <= truncLength || prevL <= truncLength) {
         truncWords.Push(w)
         prevWord := w
         continue
         }
      
      ; This and previous words are long enough. Find length of match with previous word.
      minL := L < prevL ? L : prevL
      matchedLength := 0
      Loop % minL {
         if(SubStr(w, A_index, 1) != SubStr(prevWord, A_index, 1)) {
            matchedLength := A_index - 1
            truncW := SubStr(w, 1, matchedLength) . alexF_config_Ellipsis
            Goto, ReplaceOrPush
         }
      }
      
      ; The shorter word fully matches the beginning of the longer one.
      if(L < prevL) {
         truncW := w . alexF_config_Ellipsis
      } else {
         truncW := prevWord . alexF_config_Ellipsis
      }
      matchedLength := StrLen(truncW)

      ReplaceOrPush:
      if(matchedLength >= truncLength) {
         truncWords[truncWords.MaxIndex()] := truncW ; replace previous match by possibly shorter match
         prevWord := truncW
      } else {
         truncWords.Push(w) ; match is too short, don't use ellipsis
         prevWord := w
      }
   }
   
   Return truncWords
}

;------------------------------------------------------------------------

~LButton:: 
CheckForCaretMove("LButton","UpdatePosition")
return
   

;------------------------------------------------------------------------

~RButton:: 
CheckForCaretMove("RButton","UpdatePosition")
Return

;------------------------------------------------------------------------

CheckForCaretMove(MouseButtonClick, UpdatePosition = false)
{
   global g_LastInput_Id
   global g_MouseWin_Id
   global g_OldCaretX
   global g_OldCaretY
   global g_Word
   global prefs_DetectMouseClickMove
   
   ;If we aren't using the DetectMouseClickMoveScheme, skip out
   IfNotEqual, prefs_DetectMouseClickMove, On
      Return
   
   if (UpdatePosition)
   {
      ; Update last click position in case Caret is not detectable
      ;  and update the Last Window Clicked in
      MouseGetPos, MouseX, MouseY, g_MouseWin_Id
      WinGetPos, ,TempY, , , ahk_id %g_MouseWin_Id%
   }
   
   IfEqual, MouseButtonClick, LButton
   {
      KeyWait, LButton, U    
   } else KeyWait, RButton, U
   
   IfNotEqual, g_LastInput_Id, %g_MouseWin_Id%
   {
      Return
   }
   
   SysGet, SM_CYCAPTION, 4
   SysGet, SM_CYSIZEFRAME, 33
   
   TempY += SM_CYSIZEFRAME
   IF ( ( MouseY >= TempY ) && (MouseY < (TempY + SM_CYCAPTION) ) )
   {
      Return
   }
   
   ; If we have a g_Word and an g_OldCaretX, check to see if the Caret moved
   IfNotEqual, g_OldCaretX, 
   {
      IfNotEqual, g_Word, 
      {
         if (( g_OldCaretY != HCaretY() ) || (g_OldCaretX != HCaretX() ))
         {
            ; add the word if switching lines
            AddWordToList(g_Word,0)
            ClearAllVars(true)
         }
      }
   }

   Return
}
   
   
;------------------------------------------------------------------------

InitializeHotKeys()
{
   global g_DelimiterChar
   global g_EnabledKeyboardHotKeys
   global prefs_ArrowKeyMethod
   global prefs_DisabledAutoCompleteKeys
   global prefs_LearnMode  
   
   g_EnabledKeyboardHotKeys =

   ;Setup toggle-able hotkeys

   ;Can't disable mouse buttons as we need to check to see if we have clicked the ListBox window


   ; If we disable the number keys they never get to the input for some reason,
   ; so we need to keep them enabled as hotkeys

   if(!InStr(prefs_LearnMode, "On"))
   {
      Hotkey, $^+Delete, Off
   } else {
      Hotkey, $^+Delete, Off
      ; We only want Ctrl-Shift-Delete enabled when the listbox is showing.
      g_EnabledKeyboardHotKeys .= "$^+Delete" . g_DelimiterChar
   }
   
   HotKey, $^+c, On
   
   IfEqual, prefs_ArrowKeyMethod, Off
   {
      Hotkey, $^Enter, Off
      Hotkey, $^Space, Off
      Hotkey, $Tab, Off
      Hotkey, $Right, Off
      Hotkey, $Up, Off
      Hotkey, $Down, Off
      Hotkey, $PgUp, Off
      Hotkey, $PgDn, Off
      HotKey, $Enter, Off
      Hotkey, $NumpadEnter, Off
   } else {
      g_EnabledKeyboardHotKeys .= "$Up" . g_DelimiterChar
      g_EnabledKeyboardHotKeys .= "$Down" . g_DelimiterChar
      g_EnabledKeyboardHotKeys .= "$PgUp" . g_DelimiterChar
      g_EnabledKeyboardHotKeys .= "$PgDn" . g_DelimiterChar
      If prefs_DisabledAutoCompleteKeys contains E
         Hotkey, $^Enter, Off
      else g_EnabledKeyboardHotKeys .= "$^Enter" . g_DelimiterChar
      If prefs_DisabledAutoCompleteKeys contains S
         HotKey, $^Space, Off
      else g_EnabledKeyboardHotKeys .= "$^Space" . g_DelimiterChar
      If prefs_DisabledAutoCompleteKeys contains T
         HotKey, $Tab, Off
      else g_EnabledKeyboardHotKeys .= "$Tab" . g_DelimiterChar
      If prefs_DisabledAutoCompleteKeys contains R
         HotKey, $Right, Off
      else g_EnabledKeyboardHotKeys .= "$Right" . g_DelimiterChar
      If prefs_DisabledAutoCompleteKeys contains U
         HotKey, $Enter, Off
      else g_EnabledKeyboardHotKeys .= "$Enter" . g_DelimiterChar
      If prefs_DisabledAutoCompleteKeys contains M
         HotKey, $NumpadEnter, Off
      else g_EnabledKeyboardHotKeys .= "$NumpadEnter" . g_DelimiterChar
   }

   ; remove last ascii 2
   StringTrimRight, g_EnabledKeyboardHotKeys, g_EnabledKeyboardHotKeys, 1
   
}

EnableKeyboardHotKeys()
{
   global g_DelimiterChar
   global g_EnabledKeyboardHotKeys
   Loop, Parse, g_EnabledKeyboardHotKeys, %g_DelimiterChar%
   {
      HotKey, %A_LoopField%, On
   }
   Return
}

DisableKeyboardHotKeys()
{
   global g_DelimiterChar
   global g_EnabledKeyboardHotKeys
   Loop, Parse, g_EnabledKeyboardHotKeys, %g_DelimiterChar%
   {
      HotKey, %A_LoopField%, Off
   }
   Return
}
   
;------------------------------------------------------------------------

#MaxThreadsPerHotkey 1 
    
$1:: 
$2:: 
$3:: 
$4:: 
$5:: 
$6:: 
$7:: 
$8:: 
$9:: 
$0::
CheckWord(A_ThisHotkey)
Return

$^Enter::
$^Space::
$Tab::
$Up::
$Down::
$PgUp::
$PgDn::
$Right::
$Enter::
$NumpadEnter::
EvaluateUpDown(A_ThisHotKey)
Return

$^+h::
MaybeOpenOrCloseHelperWindowManual()
Return

$^+c:: 
AddSelectedWordToList()
Return

$^+Delete::
DeleteSelectedWordFromList()
Return

; AlexF
$Esc::
HandleEscapeKey()
Return
;------------------------------------------------------------------------

; AlexF. If Esc is pressed and listbox is opened, close the listbox, but do not 
; pass Esc further.
HandleEscapeKey()
{
   global g_ListBox_Id
   
   IfNotEqual, g_ListBox_Id, 
   {
      ClearAllVars(1)
   } else {
      SendKey(A_ThisHotKey)
   }
   Return
}

; If hotkey was pressed, check whether there's a match going on and send it, otherwise send the number(s) typed 
CheckWord(Key)
{
   global g_ListBox_Id
   global g_Match          ;AlexF input, concatenation of all the lines in the listbox, separated by g_DelimiterChar
   global g_MatchStart     ;AlexF position of the first word (match) to be shown in the listbox
   global g_NumKeyMethod
   global g_SingleMatchAdj
   global g_Word
   global prefs_ListBoxRows
   global prefs_NumPresses
   
   StringRight, Key, Key, 1 ;Grab just the number pushed, trim off the "$"
   
   IfEqual, Key, 0
   {
      WordIndex := g_MatchStart + 9
   } else {
            WordIndex := g_MatchStart - 1 + Key
         }  
   
   IfEqual, g_NumKeyMethod, Off
   {
      SendCompatible(Key,0)
      ProcessKey(Key,"")
      Return
   }
   
   IfEqual, prefs_NumPresses, 2
      SuspendOn()

   ; If active window has different window ID from before the input, blank word 
   ; (well, assign the number pressed to the word) 
   if !(ReturnWinActive())
   { 
      SendCompatible(Key,0)
      ProcessKey(Key,"")
      IfEqual, prefs_NumPresses, 2
         SuspendOff()
      Return 
   } 
   
   if ReturnLineWrong() ;Make sure we are still on the same line
   { 
      SendCompatible(Key,0)
      ProcessKey(Key,"") 
      IfEqual, prefs_NumPresses, 2
         SuspendOff()
      Return 
   } 

   IfNotEqual, g_Match, 
   {
      ifequal, g_ListBox_Id,        ; only continue if match is not empty and list is showing
      { 
         SendCompatible(Key,0)
         ProcessKey(Key,"")
         IfEqual, prefs_NumPresses, 2
            SuspendOff()
         Return 
      }
   }

   ifEqual, g_Word,        ; only continue if g_word is not empty 
   { 
      SendCompatible(Key,0)
      ProcessKey(Key,"")
      IfEqual, prefs_NumPresses, 2
         SuspendOff()
      Return 
   }
      
   if ( ( (WordIndex + 1 - MatchStart) > prefs_ListBoxRows) || ( g_Match = "" ) || (g_SingleMatchAdj[WordIndex] = "") )   ; only continue if g_SingleMatchAdj is not empty 
   { 
      SendCompatible(Key,0)
      ProcessKey(Key,"")
      IfEqual, prefs_NumPresses, 2
         SuspendOff()
      Return 
   }

   IfEqual, prefs_NumPresses, 2
   {
      Input, KeyAgain, L1 I T0.5, 1234567890
      
      ; If there is a timeout, abort replacement, send key and return
      IfEqual, ErrorLevel, Timeout
      {
         SendCompatible(Key,0)
         ProcessKey(Key,"")
         SuspendOff()
         Return
      }

      ; Make sure it's an EndKey, otherwise abort replacement, send key and return
      IfNotInString, ErrorLevel, EndKey:
      {
         SendCompatible(Key . KeyAgain,0)
         ProcessKey(Key,"")
         ProcessKey(KeyAgain,"")
         SuspendOff()
         Return
      }
   
      ; If the 2nd key is NOT the same 1st trigger key, abort replacement and send keys   
      IfNotInString, ErrorLevel, %Key%
      {
         StringTrimLeft, KeyAgain, ErrorLevel, 7
         SendCompatible(Key . KeyAgain,0)
         ProcessKey(Key,"")
         ProcessKey(KeyAgain,"")
         SuspendOff()
         Return
      }

      ; If active window has different window ID from before the input, blank word 
      ; (well, assign the number pressed to the word) 
      if !(ReturnWinActive())
      { 
         SendCompatible(Key . KeyAgain,0)
         ProcessKey(Key,"")
         ProcessKey(KeyAgain,"")
         SuspendOff()
         Return 
      } 
   
      if ReturnLineWrong() ;Make sure we are still on the same line
      { 
         SendCompatible(Key . KeyAgain,0)
         ProcessKey(Key,"")
         ProcessKey(KeyAgain,"")
         SuspendOff()
         Return 
      } 
   }

   SendWord(WordIndex) ;AlexF. Type the word after numeric key was pressed. I am not using this (yet).
   IfEqual, prefs_NumPresses, 2
      SuspendOff()
   Return 
}

;------------------------------------------------------------------------

;If a hotkey related to the up/down arrows was pressed
EvaluateUpDown(Key)
{
   global g_ListBox_Id
   global g_Match            ;AlexF input, concatenation of all the lines in the listbox, separated by g_DelimiterChar
   global g_MatchPos         ;AlexF position (index) of the currently selected word (match)
   global g_MatchStart       ;AlexF position of the first word (match) to be shown in the listbox
   global g_MatchTotal
   global g_OriginalMatchStart
   global g_SingleMatchAdj
   global g_Word
   global prefs_ArrowKeyMethod
   global prefs_DisabledAutoCompleteKeys
   global prefs_ListBoxRows
   
   IfEqual, prefs_ArrowKeyMethod, Off
   {
      if (Key != "$LButton")
      {
         SendKey(Key)
         Return
      }
   }
   
   IfEqual, g_Match,
   {
      SendKey(Key)
      Return
   }

   IfEqual, g_ListBox_Id,
   {
      SendKey(Key)
      Return
   }

   if !(ReturnWinActive())
   {
      SendKey(Key)
      ClearAllVars(false)
      Return
   }

   if ReturnLineWrong()
   {
      SendKey(Key)
      ClearAllVars(true)
      Return
   }   
   
   IfEqual, g_Word, ; only continue if word is not empty
   {
      SendKey(Key)
      ClearAllVars(false)
      Return
   }
   
   if ( ( Key = "$^Enter" ) || ( Key = "$Tab" ) || ( Key = "$^Space" ) || ( Key = "$Right") || ( Key = "$Enter") || ( Key = "$LButton") || ( Key = "$NumpadEnter") )
   {
      IfEqual, Key, $^Enter
      {
         KeyTest = E
      } else IfEqual, Key, $Tab
      {
         KeyTest = T
      } else IfEqual, Key, $^Space
      {   
         KeyTest = S 
      } else IfEqual, Key, $Right
      {
         KeyTest = R
      } else IfEqual, Key, $Enter
      {
         KeyTest = U
      } else IfEqual, Key, $LButton
      {
         KeyTest = L
      } else IfEqual, Key, $NumpadEnter
      {
         KeyTest = M
      }
      
      if (KeyTest == "L") {
         ;when hitting LButton, we've already handled this condition         
      } else if prefs_DisabledAutoCompleteKeys contains %KeyTest%
      {
         SendKey(Key)
         Return     
      }
      
      if (g_SingleMatchAdj[g_MatchPos] = "") ;only continue if g_SingleMatchAdj is not empty
      {
         SendKey(Key)
         g_MatchPos := g_MatchTotal
         RebuildMatchList()
         ShowListBox()
         Return
      }
      
      SendWord(g_MatchPos)
      Return
   }

   PreviousMatchStart := g_OriginalMatchStart
   
   IfEqual, Key, $Up
   {   
      g_MatchPos--
   
      IfLess, g_MatchPos, 1
      {
         g_MatchStart := g_MatchTotal - (prefs_ListBoxRows - 1)
         IfLess, g_MatchStart, 1
            g_MatchStart = 1
         g_MatchPos := g_MatchTotal
      } else IfLess, g_MatchPos, %g_MatchStart%
      {
         g_MatchStart --
      }      
   } else IfEqual, Key, $Down
   {
      g_MatchPos++
      IfGreater, g_MatchPos, %g_MatchTotal%
      {
         g_MatchStart =1
         g_MatchPos =1
      } Else If ( g_MatchPos > ( g_MatchStart + (prefs_ListBoxRows - 1) ) )
      {
         g_MatchStart ++
      }            
   } else IfEqual, Key, $PgUp
   {
      IfEqual, g_MatchPos, 1
      {
         g_MatchPos := g_MatchTotal - (prefs_ListBoxRows - 1)
         g_MatchStart := g_MatchTotal - (prefs_ListBoxRows - 1)
      } Else {
         g_MatchPos-=prefs_ListBoxRows   
         g_MatchStart-=prefs_ListBoxRows
      }
      
      IfLess, g_MatchPos, 1
         g_MatchPos = 1
      IfLess, g_MatchStart, 1
         g_MatchStart = 1
      
   } else IfEqual, Key, $PgDn
   {
      IfEqual, g_MatchPos, %g_MatchTotal%
      {
         g_MatchPos := prefs_ListBoxRows
         g_MatchStart := 1
      } else {
         g_MatchPos+=prefs_ListBoxRows
         g_MatchStart+=prefs_ListBoxRows
      }
   
      IfGreater, g_MatchPos, %g_MatchTotal%
         g_MatchPos := g_MatchTotal
   
      If ( g_MatchStart > ( g_MatchTotal - (prefs_ListBoxRows - 1) ) )
      {
         g_MatchStart := g_MatchTotal - (prefs_ListBoxRows - 1)   
         IfLess, g_MatchStart, 1
            g_MatchStart = 1
      }
   }
   
   IfEqual, g_MatchStart, %PreviousMatchStart%
   {
      Rows := GetRows()
      IfNotEqual, g_MatchPos,
      {
         ListBoxChooseItem(Rows)
      }
   } else {
      RebuildMatchList()
      ShowListBox()
   }
   Return
}

;------------------------------------------------------------------------

ReturnLineWrong()
{
   global g_OldCaretY
   global prefs_DetectMouseClickMove
   ; Return false if we are using DetectMouseClickMove
   IfEqual, prefs_DetectMouseClickMove, On
      Return
      
   Return, ( g_OldCaretY != HCaretY() )
}

;------------------------------------------------------------------------

AddSelectedWordToList()
{      
   ClipboardSave := ClipboardAll
   Clipboard =
   Sleep, 100
   SendCompatible("^c",0)
   ClipWait, 0
   IfNotEqual, Clipboard, 
   {
      AddWordToList(Clipboard,1)
   }
   Clipboard = %ClipboardSave%
}

DeleteSelectedWordFromList()
{
   global g_MatchPos
   global g_SingleMatchDb    ;AlexF array of matched words, as they are stored in the database
   global prefs_ArrowKeyMethod ; AlexF
   global alexF_config_Ellipsis

   word := g_SingleMatchDb[g_MatchPos]
   if (word == "") ;only continue if g_SingleMatchDb is not empty
      Return
      
   lastChar := SubStr(word, 0)
   if(lastChar == alexF_config_Ellipsis) {
      Return ; cannot delete
   }

   alexF := prefs_ArrowKeyMethod
   prefs_ArrowKeyMethod := "LastPosition"
   
   DeleteWordFromList(word)
   RecomputeMatches()
   
   prefs_ArrowKeyMethod := alexF
   
}

;------------------------------------------------------------------------

EvaluateScriptPathAndTitle()
{
   ;relaunches to 64 bit or sets script title
   global g_ScriptTitle

   SplitPath, A_ScriptName,,,ScriptExtension,ScriptNoExtension,

   If A_Is64bitOS
   {
      IF (A_PtrSize = 4)
      {
         IF A_IsCompiled
         {
         
            ScriptPath64 := A_ScriptDir . "\" . ScriptNoExtension . "64." . ScriptExtension
         
            IfExist, %ScriptPath64%
            {
               Run, %ScriptPath64%, %A_WorkingDir%
               ExitApp
            }
         }
      }
   }

   if (SubStr(ScriptNoExtension, StrLen(ScriptNoExtension)-1, 2) == "64" )
   {
      StringTrimRight, g_ScriptTitle, ScriptNoExtension, 2
   } else {
      g_ScriptTitle := ScriptNoExtension
   }

   if (InStr(g_ScriptTitle, "TypingAid"))
   {
      g_ScriptTitle = TypingAid
   }
   
   return
}

;------------------------------------------------------------------------

InactivateAll()
{
   ;Force unload of Keyboard Hook and WinEventHook
   Input
   SuspendOn()
   CloseListBox()
   MaybeSaveHelperWindowPos()
   DisableWinHook()
}

SuspendOn()
{
   global g_ScriptTitle
   Suspend, On
   Menu, Tray, Tip, %g_ScriptTitle% - Inactive
   If A_IsCompiled
   {
      Menu, tray, Icon, %A_ScriptFullPath%,3,1
   } else
   {
      Menu, tray, Icon, %A_ScriptDir%\%g_ScriptTitle%-Inactive.ico, ,1
   }
}

SuspendOff()
{
   global g_ScriptTitle
   Suspend, Off
   Menu, Tray, Tip, %g_ScriptTitle% - Active
   If A_IsCompiled
   {
      Menu, tray, Icon, %A_ScriptFullPath%,1,1
   } else
   {
      Menu, tray, Icon, %A_ScriptDir%\%g_ScriptTitle%-Active.ico, ,1
   }
}   

;------------------------------------------------------------------------

BuildTrayMenu()
{

   Menu, Tray, DeleteAll
   Menu, Tray, NoStandard
   Menu, Tray, add, Settings, Configuration
   Menu, Tray, add, Pause, PauseResumeScript
   IF (A_IsCompiled)
   {
      Menu, Tray, add, Exit, ExitScript
   } else {
      Menu, Tray, Standard
   }
   Menu, Tray, Default, Settings
   ;Initialize Tray Icon
   Menu, Tray, Icon
}

;------------------------------------------------------------------------

; This is to blank all vars related to matches, ListBox and (optionally) word 
; AlexF - this resets the whole search  for matching words.
;        ClearWord - if true, forgets word typed by the user (and some other stuff)
ClearAllVars(ClearWord)
{
   global
   CloseListBox()
   Ifequal,ClearWord,1
   {
      g_Word =
      g_OldCaretY=
      g_OldCaretX=
      g_LastInput_id=
      g_ListBoxFlipped=
      g_ListBoxMaxWordHeight=
   }
   
   g_SingleMatchDb =
   g_SingleMatchAdj =
   g_Match= 
   g_MatchPos=
   g_MatchStart= 
   g_OriginalMatchStart=
   Return
}

;------------------------------------------------------------------------

;AlexF: writes 'Text' into given file 'FileName'. "Dispatch" means that 
;       the encoding (like UTF-8) is selected based on settings.
FileAppendDispatch(Text,FileName,ForceEncoding=0)
{
   IfEqual, A_IsUnicode, 1
   {
      IfNotEqual, ForceEncoding, 0
      {
         FileAppend, %Text%, %FileName%, %ForceEncoding%
      } else
      {
         FileAppend, %Text%, %FileName%, UTF-8
      }
   } else {
            FileAppend, %Text%, %FileName%
         }
   Return
}

;AlexF: reads file and writes it again, with correct encoding.
MaybeFixFileEncoding(File,Encoding)
{
   IfGreaterOrEqual, A_AhkVersion, 1.0.90.0
   {
      
      IfExist, %File%
      {    
         IfNotEqual, A_IsUnicode, 1
         {
            Encoding =
         }
         
         
         EncodingCheck := FileOpen(File,"r")
         
         If EncodingCheck
         {
            If Encoding
            {
               IF !(EncodingCheck.Encoding = Encoding)
                  WriteFile = 1
            } else
            {
               IF (SubStr(EncodingCheck.Encoding, 1, 3) = "UTF")
                  WriteFile = 1
            }
         
            IF WriteFile
            {
               Contents := EncodingCheck.Read()
               EncodingCheck.Close()
               EncodingCheck =
               FileCopy, %File%, %File%.preconv.bak
               FileDelete, %File%
               FileAppend, %Contents%, %File%, %Encoding%
               
               Contents =
            } else
            {
               EncodingCheck.Close()
               EncodingCheck =
            }
         }
      }
   }
}

;------------------------------------------------------------------------

GetOSVersion()
{
   return ((r := DllCall("GetVersion") & 0xFFFF) & 0xFF) "." (r >> 8)
}

;------------------------------------------------------------------------

MaybeCoInitializeEx()
{
   global g_NULL
   global g_ScrollEventHook
   global g_WinChangedEventHook
   
   if (!g_WinChangedEventHook && !g_ScrollEventHook)
   {
      DllCall("CoInitializeEx", "Ptr", g_NULL, "Uint", g_NULL)
   }
   
}


MaybeCoUninitialize()
{
   global g_WinChangedEventHook
   global g_ScrollEventHook
   if (!g_WinChangedEventHook && !g_ScrollEventHook)
   {
      DllCall("CoUninitialize")
   }
}

;------------------------------------------------------------------------

Configuration:
GoSub, LaunchSettings
Return

PauseResumeScript:
if (g_PauseState == "Paused")
{
   g_PauseState =
   Pause, Off
   EnableWinHook()
   Menu, tray, Uncheck, Pause
} else {
   g_PauseState = Paused
   DisableWinHook()
   SuspendOn()
   Menu, tray, Check, Pause
   Pause, On, 1
}
Return

ExitScript:
ExitApp
Return
   
SaveScript:
; Close the ListBox if it's open
CloseListBox()

SuspendOn()

;Change the cleanup performance speed
SetBatchLines, 20ms
Process, Priority,,Normal

;Grab the Helper Window Position if open
MaybeSaveHelperWindowPos()

;Write the Helper Window Position to the Preferences File
MaybeWriteHelperWindowPos()

; Update the Learned Words
if(alexF_config_BackupWordsAndCounts) {
   MaybeUpdateWordAndCountTextFile()
} else {
   MaybeUpdateWordlist()
}

ExitApp

#Include %A_ScriptDir%\Includes\Conversions.ahk
#Include %A_ScriptDir%\Includes\Helper.ahk
#Include %A_ScriptDir%\Includes\ListBox.ahk
#Include %A_ScriptDir%\Includes\Preferences File.ahk
#Include %A_ScriptDir%\Includes\Sending.ahk
#Include %A_ScriptDir%\Includes\Settings.ahk
#Include %A_ScriptDir%\Includes\Window.ahk
#Include %A_ScriptDir%\Includes\Wordlist.ahk
#Include <DBA>
#Include <_Struct>
