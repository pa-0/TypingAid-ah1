
RecomputeMatches()
{
   ; This function will take the given word, and will recompile the list of matches and redisplay the wordlist.
   global g_MatchTotal         ;AlexF count of matched words
   global g_SingleMatchDb      ;AlexF array of matched words, as they are stored in the database
   global g_SingleMatchAdj     ;AlexF array of matched words, with adjusted capitalization. This is what user sees. 
   global g_Word               ; AlexF word typed by user
   global g_WordListDB
   global prefs_ArrowKeyMethod
   global prefs_LearnMode
   global prefs_ListBoxRows
   global prefs_NoBackSpace
   global prefs_ShowLearnedFirst
   global prefs_SuppressMatchingWord
   global alexF_config_OrderByLength
   global alexF_config_PreventScrollbar
   
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
      if (alexF_config_PreventScrollbar) {
         LimitTotalMatches := prefs_ListBoxRows
      } else {
         LimitTotalMatches = 200
      }
   }
   
   StringUpper, WordAllCaps, g_Word
   
   ; AlexF added: check capitalization
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
   WhereQuery := " WHERE wordindexed GLOB '" . WordMatchEscaped . "*' " . SuppressMatchingWordQuery . WordAccentQuery
   ;AlexF one-column, one-row table with minimum of counts of occurrences of first 'LimitTotalMatches' learned words.
   NormalizeTable := g_WordListDB.Query("SELECT MIN(count) AS normalize FROM Words"
   NormalizeTable .= WhereQuery . "AND count IS NOT NULL LIMIT " . LimitTotalMatches . ";")
   
   for each, row in NormalizeTable.Rows
   {
      Normalize := row[1]
   }
      
   IfEqual, Normalize,
   {
      Normalize := 0
   }
   
   ;_3. Second query - actually retrieve matches, in certain order (was: more frequent and longer words first, AlexF changed)
   if (alexF_config_OrderByLength) 
   {
      OrderByQuery := " ORDER BY LENGTH(word)"
   } 
   else 
   {
      WordLen := StrLen(g_Word)
      OrderByQuery := " ORDER BY CASE WHEN count IS NULL then "
      IfEqual, prefs_ShowLearnedFirst, On
      {
         OrderByQuery .= "ROWID + 1 else 0"
      } 
      else 
      {
         OrderByQuery .= "ROWID else 'z'"
      }
      
      ;AlexF (count - min) * (1 - 0.75/nExtraChars) -- advantage to more frequent and longer words. The lines below or above break function list
      clause1 := " end, CASE WHEN count IS NOT NULL then "
      clause2 := "( (count - " .Normalize . ") * ( 1 - ( '0.75' / (LENGTH(word) - ". WordLen . ")))) end DESC, Word"
      OrderByQuery .= clause1 . clause2
   }
   ;AlexF table of matched words
   query := "SELECT word FROM Words" 
   query .= WhereQuery . OrderByQuery . " LIMIT " . LimitTotalMatches . ";"
   Matches := g_WordListDB.Query(query)
   
   g_SingleMatchDb := Object()
   g_SingleMatchAdj := Object()
   for each, row in Matches.Rows
   {      
  
      oldLength := StrLen(row[1])
      VarSetCapacity(word, oldLength + 12, 0) ;AlexF added. 12 bytes, just in case. They say, for Unicode 2 bytes per char is enough. Apparently, they are not using UTF-8 here?
      
      g_SingleMatchDb[++g_MatchTotal] := row[1]
      ; If row[1] has "normal" capitalization ("|firstCap|"), it will be adjusted to match word.
      word := AdjustCapitalization(row[1], targetCapitalization, g_Word)
      g_SingleMatchAdj[g_MatchTotal] := word

; AlexF  - Works:     DllCall("TAHelperU64.dll\AddEllipses1", "Str", word)
;      MsgBox % "Converted '" . row[1] . "' of length " . oldLength . " to '" . word . "' of length " . StrLen(word) . ". ErrorLevel: " . ErrorLevel
      
      continue
   }
   
   ;If no match then clear Tip 
   IfEqual, g_MatchTotal, 0
   {
      ClearAllVars(false)
      Return 
   } 
   
   /* 
   ; AlexF - Learning DllCall
   ;------------------------------------------
   numbers := 0 ; [] WORKS - 1
   index := 0
   intSize := 4 ; A_PtrSize = 8 on 64-bit system?
   VarSetCapacity(numbers,g_MatchTotal * intSize) ; create a block of memory
   Loop % g_MatchTotal
   {
      index += 1
      NumPut(10010 * index, numbers, (index - 1) * intSize)
   }
   DllCall("TAHelperU64.dll\ReadNumbers", "Ptr", &numbers, "Int", index)

   ;------------------------------------------
   strings := "" ; WORKS - 2
   foo1 := "Foo1"
   foo2 := "fOO2"
   MsgBox % "Address of foo1 is " . &foo1 . "; Content of foo1 is " . foo1 . "; The code of the first char is " *(&foo1)
   MsgBox % "Address of foo2 is " . &foo2 . "; Content of foo2 is " . foo2 . "; The code of the first char is " *(&foo2)
   VarSetCapacity(strings,2 * A_PtrSize + 128) ; create a block of memory
   NumPut(&foo1, strings, 0)
   NumPut(&foo2, strings, A_PtrSize)
   DllCall("TAHelperU64.dll\AddEllipses", "Ptr", &strings, "Int", 2)

   ;------------------------------------------
   strings := "" ; WORKS - 3
   VarSetCapacity(strings,g_MatchTotal * A_PtrSize + 128) ; create a block of memory
   Loop % g_MatchTotal
   {
      index += 1
      word%index% := g_SingleMatchAdj[index]
      NumPut(&(word%index%), &strings, (index - 1) * A_PtrSize)
   }
   DllCall("TAHelperU64.dll\AddEllipses", "Ptr", &strings, "Int", g_MatchTotal)
   */

  
   SetupMatchPosition() ; what position to highlight in the listbox
   RebuildMatchList() ; generate g_Match - concatenation of all the lines in the listbox
   ShowListBox()
}
