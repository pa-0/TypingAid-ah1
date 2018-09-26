; These functions and labels are related to maintenance of the wordlist

ReadWordList()
{
   global g_LegacyLearnedWords
   global g_ScriptTitle
   global g_WordListDone
   global g_WordListDB
   global alexF_config_BackupWordsAndCounts
   
   ;mark the wordlist as not done
   g_WordListDone = 0
   
   ;_1. Prepare two files and database for reading; reset journal
   WordlistFileName = wordlist.txt
   
   Wordlist = %A_ScriptDir%\%WordlistFileName% ; AlexF - file with predefined words (I do not have it so far)
   WordlistLearned = %A_ScriptDir%\WordlistLearned.txt ; AlexF - file with learned words (a backup for the database)
   if(alexF_config_BackupWordsAndCounts) {
      WordlistLearned = %A_ScriptDir%\WordlistLearned.csv
   }
   
   MaybeFixFileEncoding(Wordlist,"UTF-8")
   MaybeFixFileEncoding(WordlistLearned,"UTF-8")

   g_WordListDB := DBA.DataBaseFactory.OpenDataBase("SQLite", A_ScriptDir . "\WordlistLearned.db" )
   
   if !g_WordListDB
   {
      msgbox Problem opening database '%A_ScriptDir%\WordlistLearned.db' - fatal error...
      exitapp
   }
	
   g_WordListDB.Query("PRAGMA journal_mode = TRUNCATE;")
   
   ;_2. Maybe create a fresh new database (schema only)
   DatabaseRebuilt := MaybeConvertDatabase()
         
   FileGetSize, WordlistSize, %Wordlist%
   FileGetTime, WordlistModified, %Wordlist%, M
   FormatTime, WordlistModified, %WordlistModified%, yyyy-MM-dd HH:mm:ss
   
   ;_3. I do not have 'Wordlists' table in database, hence LearnedWordsTable is empty.
   ;    Not sure what this step does. Its result is mere LoadWordlist := Insert'
   if (!DatabaseRebuilt) {
      LearnedWordsTable := g_WordListDB.Query("SELECT wordlistmodified, wordlistsize FROM Wordlists WHERE wordlist = '" . WordlistFileName . "';")
      
      LoadWordlist := "Insert"
      
      For each, row in LearnedWordsTable.Rows
      {
         WordlistLastModified := row[1]
         WordlistLastSize := row[2]
         
         if (WordlistSize != WordlistLastSize || WordlistModified != WordlistLastModified) {
            LoadWordlist := "Update"
            CleanupWordList()
         } else {
            LoadWordlist =
            CleanupWordList(true)
         }
      }
   } else {
      LoadWordlist := "Insert"
   }
   
   ;_4. Read from multiple predefined wordlists. I have none, so far.
   if (LoadWordlist) {
      Progress, M, Please wait..., Loading wordlist, %g_ScriptTitle%
      g_WordListDB.BeginTransaction()
      ;reads list of words from file. AlexF: these are non-learned, predefined lists.
      FileRead, ParseWords, %Wordlist%
      Loop, Parse, ParseWords, `n, `r
      {
         ParseWordsCount++
      }
      Loop, Parse, ParseWords, `n, `r
      {
         ParseWordsSubCount++
         ProgressPercent := Round(ParseWordsSubCount/ParseWordsCount * 100)
         if (ProgressPercent <> OldProgressPercent)
         {
            Progress, %ProgressPercent%
            OldProgressPercent := ProgressPercent
         }
         /*  AlexF - I do not have this legacy/old files with ';LEARNEDWORDS;'
         IfEqual, A_LoopField, `;LEARNEDWORDS`;
         {
            if (DatabaseRebuilt)
            {
               LearnedWordsCount=0
               g_LegacyLearnedWords=1 ; Set Flag that we need to convert wordlist file
            } else {
               break
            }
         } else {
            AddWordToList(A_LoopField,0,"ForceLearn",LearnedWordsCount)
         }
         */
         LearnedWordCountValue := 1 ;AlexF added this line -- don't know whether it makes sense. This step _4 is not exercised in my settings.
         AddWordToList(A_LoopField,0,LearnedWordCountValue)
      }
      ParseWords =
      g_WordListDB.EndTransaction()
      Progress, Off
      
      if (LoadWordlist == "Update") {
         g_WordListDB.Query("UPDATE wordlists SET wordlistmodified = '" . WordlistModified . "', wordlistsize = '" . WordlistSize . "' WHERE wordlist = '" . WordlistFileName . "';")
      } else {
         g_WordListDB.Query("INSERT INTO Wordlists (wordlist, wordlistmodified, wordlistsize) VALUES ('" . WordlistFileName . "','" . WordlistModified . "','" . WordlistSize . "');")
      }
   }
   
   ;_5. Read words and counts from the backup, 'WordlistLearned.csv'
   if (DatabaseRebuilt)
   {
      Progress, M, Please wait..., Converting learned words, %g_ScriptTitle%
    
      ;Force LearnedWordCountValue to 0 as we are now processing Learned Words
      LearnedWordCountValue=0
      
      g_WordListDB.BeginTransaction()
      if(alexF_config_BackupWordsAndCounts) {
         ;reads list of words and counts from file 
         FileRead, WordsAndCounts, %WordlistLearned%
         Loop, Parse, WordsAndCounts, `n, `r
         {
            wordAndCount := StrSplit(A_LoopField, ",") 
            AddWordToList(wordAndCount[1] ,0, wordAndCount[2])
         }
         WordsAndCounts =
      } else {
         ;reads list of words from file 
         FileRead, ParseWords, %WordlistLearned%
         Loop, Parse, ParseWords, `n, `r
         {
            
            AddWordToList(A_LoopField,0,LearnedWordCountValue)
         }
         ParseWords =
      }
      g_WordListDB.EndTransaction()
      
      Progress, 50, Please wait..., Converting learned words, %g_ScriptTitle%

      if(!alexF_config_BackupWordsAndCounts) {
         ;reverse the numbers of the word counts in memory
         ReverseWordNums(LearnedWordCountValue)
      }
      
      g_WordListDB.Query("INSERT INTO LastState VALUES ('tableConverted','1',NULL);")
      
      Progress, Off
   }

   ;mark the wordlist as completed
   g_WordlistDone = 1
   Return
}

;------------------------------------------------------------------------

;AlexF. DON'T USE.
;       This function is called if the database content in WordlistLearned.db was lost 
;       and then recreated from the backup - WordlistLearned.csv. It assigns word counts 
;       according to the inverse order as the words are listed in WordlistLearned.txt.
;       (They are ordered by frequency in WordlistLearned.txt, and thus their *relative* frequencies are
;       imitated). I SHOULD MAKE SURE NOT TO USE IT, because I will store actual frequencies.
ReverseWordNums(LearnedWordCountValue)
{
   ; This function will reverse the read numbers since now we know the total number of words
   global prefs_LearnCount
   global g_WordListDB

   LearnedWordCountValue+= (prefs_LearnCount - 1)

   ; AlexF: table of learned words only
   LearnedWordsTable := g_WordListDB.Query("SELECT word FROM Words WHERE count IS NOT NULL;")

   g_WordListDB.BeginTransaction()
   For each, row in LearnedWordsTable.Rows
   {
      SearchValue := row[1]
      StringReplace, SearchValueEscaped, SearchValue, ', '', All
      WhereQuery := "WHERE word = '" . SearchValueEscaped . "'"
      g_WordListDB.Query("UPDATE words SET count = (SELECT " . LearnedWordCountValue . " - count FROM words " . WhereQuery . ") " . WhereQuery . ";")
   }
   g_WordListDB.EndTransaction()

   Return
   
}

;------------------------------------------------------------------------

; AlexF  Adds word to the database (or increases count of the existing word), if appropriate
AddWordToList(AddWordRaw,ForceCountNewOnly, LearnedWordCountValue=0)
{
   ;AddWord = Word to add to the list
   ;ForceCountNewOnly = force this word to be permanently learned even if learnmode is off (because of prefs_EndWordCharacters)
   ;ForceLearn = disables some checks in CheckValid - this parameter is ignored removed by AlexF.
   ;LearnedWordCountValue = will be non-zero only when words are read in from the backup "WordlistLearned.csv"
   global prefs_DoNotLearnStrings
   global prefs_ForceNewWordCharacters
   global prefs_LearnCount
   global prefs_LearnLength
   global prefs_LearnMode
   global g_WordListDone
   global g_WordListDB
   global g_LearnedWordInsertionTime ; AlexF, in milliseconds, 10 ms resolution
   
   if !(CheckValid(AddWordRaw))
      return
   
   ;AlexF. Before entering to the database, need to "normalize" capitalization
   AddWord := AdjustCapitalization(AddWordRaw, "|firstCap|")
   
   TransformWord(AddWord, AddWordTransformed, AddWordIndexTransformed)

   ; If wordlist is not completed yet
   IfEqual, g_WordListDone, 0 
   {
      g_WordListDB.Query("INSERT INTO words (wordindexed, word, count) VALUES ('" 
      . AddWordIndexTransformed . "','" . AddWordTransformed . "','" 
      . LearnedWordCountValue . "');") ;if this is read from the wordlist, AlexF expects LearnedWordCountValue > 0

      Return
   } 
   
   if (!InStr(prefs_LearnMode, "On")) {
      Return
   }
   
   ; This is an on-the-fly learned word
   AddWordInList := g_WordListDB.Query("SELECT * FROM words WHERE word = '" . AddWordTransformed . "';")
   
   IF !( AddWordInList.Count() > 0 ) ; if the word is not in the list
   {
   
      IF (StrLen(AddWord) < prefs_LearnLength) ; don't add the word if it's not longer than the minimum length for learning
         Return
      
      if AddWord contains %prefs_ForceNewWordCharacters%
         Return
            
      if AddWord contains %prefs_DoNotLearnStrings%
         Return
            
      CountValue = 1
      
      g_WordListDB.Query("INSERT INTO words (wordindexed, word, count) VALUES ('" 
      . AddWordIndexTransformed . "','" 
      . AddWordTransformed . "','" . CountValue . "');")

      g_LearnedWordInsertionTime := A_TickCount
   } else
   {
      UpdateWordCount(AddWord) ;Increment the word count if it's already in the list and we aren't forcing it on
   }
   
   Return
}

;AlexF returns 1 if the word is valid for adding to the database.
CheckValid(Word)
{
   
   Ifequal, Word,  ;If we have no word to add, skip out.
      Return
            
   if Word is space ;If Word is only whitespace, skip out.
      Return
   
   if ( Substr(Word,1,1) = ";" ) ;If first char is ";", clear word and skip out.
   {
      Return
   }
   
   IF ( StrLen(Word) <= prefs_Length ) ; don't add the word if it's not longer than the minimum length
   {
      Return
   }
   
   /* AlexF: I decided not to support this check. I also will remove 'ForceLearn' from AddWordToList().
   
   ;Anything below this line should not be checked if we want to Force Learning the word (Ctrl-Shift-C or coming from wordlist.txt)
   If ForceLearn
      Return, 1
   
   ;if Word does not contain at least one alpha character, skip out.
   IfEqual, A_IsUnicode, 1
   {
      if ( RegExMatch(Word, "S)\pL") = 0 ) ; AlexF: white space (S followed by letter \pL not found
      {
         return
      }
   } else if ( RegExMatch(Word, "S)[a-zA-Zà-öø-ÿÀ-ÖØ-ß]") = 0 ) ; AlexF: not found
   {
      Return
   }
   */
   
   Return, 1
}

;AlexF. 
; AddWord - input, original word
; AddWordTransformed - output, same word, with substitution ' => ''
; AddWordIndexTransformed - output, same word, with normalized 'accents', capitalized and with substitution ' => ''
TransformWord(AddWord, ByRef AddWordTransformed, ByRef AddWordIndexTransformed)
{
   AddWordIndex := AddWord
   
   ; normalize accented characters
   AddWordIndex := StrUnmark(AddWordIndex)
   
   StringUpper, AddWordIndex, AddWordIndex
   
   StringReplace, AddWordTransformed, AddWord, ', '', All
   StringReplace, AddWordIndexTransformed, AddWordIndex, ', '', All
}

DeleteWordFromList(DeleteWord)
{
   global prefs_LearnMode
   global g_WordListDB
   
   IfEqual, DeleteWord,  ;If we have no word to delete, skip out.
      Return
            
   if DeleteWord is space ;If DeleteWord is only whitespace, skip out.
      Return
   
   if(!InStr(prefs_LearnMode, "On"))
      Return
   
   StringReplace, DeleteWordEscaped, DeleteWord, ', '', All
   g_WordListDB.Query("DELETE FROM words WHERE word = '" . DeleteWordEscaped . "';")
      
   Return   
}

;------------------------------------------------------------------------

;AlexF. Modifies count for one word in the database.
UpdateWordCount(word)
{
   global prefs_LearnMode
   global g_WordListDB
   ;Word = Word to increment count for
   
   ;Should only be called when LearnMode is on  
   IfEqual, prefs_LearnMode, Off
      Return
   
   StringReplace, wordEscaped, word, ', '', All
   g_WordListDB.Query("UPDATE words SET count = count + 1 WHERE word = '" . wordEscaped . "';")
   
   Return
}

;------------------------------------------------------------------------

CleanupWordList(LearnedWordsOnly := false)
{
   ;Function cleans up all words that are less than the LearnCount threshold or have a NULL for count
   ;(NULL in count represents a 'wordlist.txt' word, as opposed to a learned word)
   global g_ScriptTitle
   global g_WordListDB
   global prefs_LearnCount
   Progress, M, Please wait..., Cleaning wordlist, %g_ScriptTitle%
   if (LearnedWordsOnly) {
      g_WordListDB.Query("DELETE FROM Words WHERE count < " . prefs_LearnCount . " AND count IS NOT NULL;")
   } else {
      g_WordListDB.Query("DELETE FROM Words WHERE count < " . prefs_LearnCount . " OR count IS NULL;")
   }
   Progress, Off
}

;------------------------------------------------------------------------

;AlexF. Updates content of "WordlistLearned.txt" file.
MaybeUpdateWordlist()
{
   global g_LegacyLearnedWords
   global g_WordListDB
   global g_WordListDone
   global prefs_LearnCount
   
   ; Update the Learned Words
   IfEqual, g_WordListDone, 1
   {
      ;AlexF. Get learned words sorted by frequency, most frequent first.
      SortWordList := g_WordListDB.Query("SELECT Word FROM Words WHERE count >= " . prefs_LearnCount . " AND count IS NOT NULL ORDER BY count DESC;")
      
      for each, row in SortWordList.Rows
      {
         TempWordList .= row[1] . "`r`n"
      }
      
      If ( SortWordList.Count() > 0 )
      {
         StringTrimRight, TempWordList, TempWordList, 2
   
         FileDelete, %A_ScriptDir%\Temp_WordlistLearned.txt
         
         ; AlexF. Write sorted words into "Temp_WordlistLearned.txt"
         FileAppendDispatch(TempWordList, A_ScriptDir . "\Temp_WordlistLearned.txt")
         
         ;AlexF. Then copy it into "WordlistLearned.txt"
         FileCopy, %A_ScriptDir%\Temp_WordlistLearned.txt, %A_ScriptDir%\WordlistLearned.txt, 1
         
         FileDelete, %A_ScriptDir%\Temp_WordlistLearned.txt
         
      /*  AlexF - I do not have this legacy/old files with ';LEARNEDWORDS;'
         ; Convert the Old Wordlist file to not have ;LEARNEDWORDS;
         IfEqual, g_LegacyLearnedWords, 1
         {
            TempWordList =
            FileRead, ParseWords, %A_ScriptDir%\Wordlist.txt
            LearnedWordsPos := InStr(ParseWords, "`;LEARNEDWORDS`;",true,1) ;Check for Learned Words
            TempWordList := SubStr(ParseWords, 1, LearnedwordsPos - 1) ;Grab all non-learned words out of list
            ParseWords = 
            FileDelete, %A_ScriptDir%\Temp_Wordlist.txt
            FileAppendDispatch(TempWordList, A_ScriptDir . "\Temp_Wordlist.txt")
            FileCopy, %A_ScriptDir%\Temp_Wordlist.txt, %A_ScriptDir%\Wordlist.txt, 1
            FileDelete, %A_ScriptDir%\Temp_Wordlist.txt
         }   
      */
      }
   }
   
   g_WordListDB.Close(),
   
}

;------------------------------------------------------------------------

; Removes marks from letters.  Requires Windows Vista or later.
; Code by Lexikos, based on MS documentation
StrUnmark(string) {
   global g_OSVersion
   global g_NormalizationKD
   if (g_OSVersion < 6.0)
   {
      return string
   }
   
   len := DllCall("Normaliz.dll\NormalizeString", "int", g_NormalizationKD, "wstr", string, "int", StrLen(string), "ptr", 0, "int", 0)  ; Get *estimated* required buffer size.
   Loop {
      VarSetCapacity(buf, len * 2)
      len := DllCall("Normaliz.dll\NormalizeString", "int", g_NormalizationKD, "wstr", string, "int", StrLen(string), "ptr", &buf, "int", len)
      if len >= 0
         break
      if (A_LastError != 122) ; ERROR_INSUFFICIENT_BUFFER
         return string
      len *= -1  ; This is the new estimate.
   }
   ; Remove combining marks and return result.
   string := RegExReplace(StrGet(&buf, len, "UTF-16"), "\pM")
   
   StringReplace, string, string, æ, ae, All
   StringReplace, string, string, Æ, AE, All
   StringReplace, string, string, œ, oe, All
   StringReplace, string, string, Œ, OE, All
   StringReplace, string, string, ß, ss, All   
   
   return, string  
   
}

;----- Functions below are added by AlexF --------------------------------------

; Updates content of "WordlistLearned.csv" file,
; which includes learned words and their frequencies.
MaybeUpdateWordAndCountTextFile()
{
   global g_LegacyLearnedWords
   global g_WordListDB
   global g_WordListDone
   global prefs_LearnCount
   
   ; Purge database of one-time entered words (possibly typos)
   g_WordListDB.Query("DELETE FROM words WHERE count = 1;")
   
   ; Update the Learned Words
   IfEqual, g_WordListDone, 1
   {
      ;AlexF. Get learned words sorted alphabetically.
      LearnedWordsTable := g_WordListDB.Query("SELECT word, count FROM Words WHERE count >= " . prefs_LearnCount . " AND count IS NOT NULL ORDER BY wordindexed ASC;")
      
      for each, row in LearnedWordsTable.Rows
      {
         TempWordList .= row[1] . "," . row[2] . "`r`n"
      }
      
      If ( LearnedWordsTable.Count() > 0 )
      {
         StringTrimRight, TempWordList, TempWordList, 2
   
         FileDelete, %A_ScriptDir%\Temp_WordlistLearned.csv
         
         ;Write words and counts into "Temp_WordlistLearned.csv"
         FileAppendDispatch(TempWordList, A_ScriptDir . "\Temp_WordlistLearned.csv")
         
         ;AlexF. Then copy it into "WordlistLearned.csv"
         FileCopy, %A_ScriptDir%\Temp_WordlistLearned.csv, %A_ScriptDir%\WordlistLearned.csv, 1
         
         FileDelete, %A_ScriptDir%\Temp_WordlistLearned.csv
      }
   }
   
   g_WordListDB.Close(),
}

; Returns "|allCaps|", "|firstCap|", "|allLow|" or "|custom|"
GetCapitalization(Word) {
   StringUpper, WordAllCaps, Word
   
   ; AlexF added: check capitalization
   capitalization := "|custom|"
   if (Word == WordAllCaps) {
      capitalization := "|allCaps|"
   } else {
      suffix := SubStr(Word, 2) ; without first char
      StringLower, suffixAllLower, suffix
      if(suffix == suffixAllLower) {
         if(SubStr(Word, 1, 1) == SubStr(WordAllCaps, 1, 1)) {
            capitalization := "|firstCap|"
         } else {
            capitalization := "|allLow|"
         }
      }
   }
   Return capitalization
}

; If Word is already has custom capitalization, does not modify it.
; Changes Word to match targetCapitalization (can be "|allCaps|", "|firstCap|", "|allLow|", or "|custom|").
; Otherwise, changes Word, so that its start is same as the shorter pattern (targetCapitalization)
; and its tail is whatever it was.
; Returns modified Word.
AdjustCapitalization(Word, targetCapitalization, headPattern="") {
   capitalization := GetCapitalization(Word) ; If Word is from database, can be only "|custom|" or "|firstCap|"; if from sendingToWnd, can be anything
   
   if(capitalization == "|custom|") {
      Return Word
   }

   if(targetCapitalization == "|firstCap|") {
      firstChar := SubStr(Word, 1, 1)
      StringUpper, firstCharUpper, firstChar
      tail := SubStr(Word, 2) ; without first char
      StringLower, tailAllLower, tail
      Return firstCharUpper . tailAllLower
   }

   if(targetCapitalization == "|allCaps|") {
      StringUpper, WordAllCaps, Word
      Return WordAllCaps
   }

   if(targetCapitalization == "|allLow|") {
      StringLower, WordAllLower, Word
      Return WordAllLower
   }

   if(targetCapitalization == "|custom|") {
      AdjWord := headPattern . SubStr(Word, StrLen(headPattern) + 1)
      Return AdjWord
   }

   
   Return "Error in script. Word='" . Word . "'; targetCapitalization='" . targetCapitalization . "'."
}


