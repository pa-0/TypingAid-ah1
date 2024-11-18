; These functions and labels are related maintenance of the wordlist

ReadWordList()
{
   global g_LegacyLearnedWords
   global g_ScriptTitle
   global g_WordListDone
   global g_WordListDB
   ;mark the wordlist as not done
   g_WordListDone = 0
   
   global PathToUserFiles
   global Wordlist
   global WordlistFileName

   WordlistLearned = %PathToUserFiles%\WordlistLearned.txt
   
   MaybeFixFileEncoding(Wordlist,"UTF-8")
   MaybeFixFileEncoding(WordlistLearned,"UTF-8")

   g_WordListDB := DBA.DataBaseFactory.OpenDataBase("SQLite", PathToUserFiles . "\WordlistLearned.db" )
   
   if !g_WordListDB
   {
      msgbox Problem opening database '%PathToUserFiles%\WordlistLearned.db' - fatal error...
      exitapp
   }
	
   g_WordListDB.Query("PRAGMA journal_mode = TRUNCATE;")
   
   DatabaseRebuilt := MaybeConvertDatabase()
         
   FileGetSize, WordlistSize, %Wordlist%
   FileGetTime, WordlistModified, %Wordlist%, M
   FormatTime, WordlistModified, %WordlistModified%, yyyy-MM-dd HH:mm:ss
   
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
   
   if (LoadWordlist) {
      Progress, M, Please wait..., Loading wordlist, %g_ScriptTitle%
      g_WordListDB.BeginTransaction()
      ;reads list of words from file 
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
   
   if (DatabaseRebuilt)
   {
      Progress, M, Please wait..., Converting learned words, %g_ScriptTitle%
    
      ;Force LearnedWordsCount to 0 if not already set as we are now processing Learned Words
      IfEqual, LearnedWordsCount,
      {
         LearnedWordsCount=0
      }
      
      g_WordListDB.BeginTransaction()
      ;reads list of words from file 
      FileRead, ParseWords, %WordlistLearned%
      Loop, Parse, ParseWords, `n, `r
      {
         
         AddWordToList(A_LoopField,0,"ForceLearn",LearnedWordsCount)
      }
      ParseWords =
      g_WordListDB.EndTransaction()
      
      Progress, 50, Please wait..., Converting learned words, %g_ScriptTitle%

      ;reverse the numbers of the word counts in memory
      ReverseWordNums(LearnedWordsCount)
      
      g_WordListDB.Query("INSERT INTO LastState VALUES ('tableConverted','1',NULL);")
      
      Progress, Off
   }

   ;mark the wordlist as completed
   g_WordlistDone = 1
   Return
}

;------------------------------------------------------------------------

ReverseWordNums(LearnedWordsCount)
{
   ; This function will reverse the read numbers since now we know the total number of words
   global prefs_LearnCount
   global g_WordListDB

   LearnedWordsCount+= (prefs_LearnCount - 1)

   LearnedWordsTable := g_WordListDB.Query("SELECT word FROM Words WHERE count IS NOT NULL;")

   g_WordListDB.BeginTransaction()
   For each, row in LearnedWordsTable.Rows
   {
      SearchValue := row[1]
      StringReplace, SearchValueEscaped, SearchValue, ', '', All
      WhereQuery := "WHERE word = '" . SearchValueEscaped . "'"
      g_WordListDB.Query("UPDATE words SET count = (SELECT " . LearnedWordsCount . " - count FROM words " . WhereQuery . ") " . WhereQuery . ";")
   }
   g_WordListDB.EndTransaction()

   Return
   
}

;------------------------------------------------------------------------

AddWordToList(AddWord,ForceCountNewOnly,ForceLearn=false, ByRef LearnedWordsCount = false)
{
   ;AddWord = Word to add to the list
   ;ForceCountNewOnly = force this word to be permanently learned even if learnmode is off
   ;ForceLearn = disables some checks in CheckValid
   ;LearnedWordsCount = if this is a stored learned word, this will only have a value when LearnedWords are read in from the wordlist
   global prefs_DoNotLearnStrings
   global prefs_ForceNewWordCharacters
   global prefs_LearnCount
   global prefs_LearnLength
   global prefs_LearnMode
   global g_WordListDone
   global g_WordListDB
   
   if !(LearnedWordsCount) {
      ;This section handles the creation of new wordlist database from text files.
      StringSplit, SplitAddWord, AddWord, |
      
      IfEqual, SplitAddWord2, D
      {
         AddWordDescription := SplitAddWord3
         AddWord := SplitAddWord1
         IfEqual, SplitAddWord4, R
         {
            AddWordReplacement := SplitAddWord5
         }
      } else IfEqual, SplitAddword2, R
      {
         AddWordReplacement := SplitAddWord3
         AddWord := SplitAddWord1
         IfEqual, SplitAddWord4, D
         {
            AddWordDescription := SplitAddWord5
         }
      }
   }
         
   if !(CheckValid(AddWord,ForceLearn))
      return
   
   TransformWord(AddWord, AddWordReplacement, AddWordDescription, AddWordTransformed, AddWordIndexTransformed, AddWordReplacementTransformed, AddWordDescriptionTransformed)

   IfEqual, g_WordListDone, 0 ;if this is read from the wordlist
   {
      ;If wordlist is not yet processed...
      IfNotEqual,LearnedWordsCount,  ;if this is a stored learned word, this will only have a value when LearnedWords are read in from the wordlist
      {
         ; must update wordreplacement since SQLLite3 considers nulls unique
         g_WordListDB.Query("INSERT INTO words (wordindexed, word, count, wordreplacement) VALUES ('" . AddWordIndexTransformed . "','" . AddWordTransformed . "','" . LearnedWordsCount++ . "','');")
      } else {
         if (AddWordReplacement)
         {
            WordReplacementQuery := "'" . AddWordReplacementTransformed . "'"
         } else {
            WordReplacementQuery := "''"
         }
         
         if (AddWordDescription)
         {
            WordDescriptionQuery := "'" . AddWordDescriptionTransformed . "'"
         } else {
            WordDescriptionQuery := "NULL"
         }
         g_WordListDB.Query("INSERT INTO words (wordindexed, word, worddescription, wordreplacement) VALUES ('" . AddWordIndexTransformed . "','" . AddWordTransformed . "'," . WordDescriptionQuery . "," . WordReplacementQuery . ");")
      }
      
   } else if (prefs_LearnMode = "On" || ForceCountNewOnly == 1)
   { 
      ; If this is an on-the-fly learned word
      AddWordInList := g_WordListDB.Query("SELECT * FROM words WHERE word = '" . AddWordTransformed . "';")
      
      IF !( AddWordInList.Count() > 0 ) ; if the word is not in the list
      {
      
         IfNotEqual, ForceCountNewOnly, 1
         {
            IF (StrLen(AddWord) < prefs_LearnLength) ; don't add the word if it's not longer than the minimum length for learning if we aren't force learning it
            {
               ;Word not learned: Length less than LearnLength
               ClearWordHierarchy() 
               Return
            }
            
            if AddWord contains %prefs_ForceNewWordCharacters%
            {
               ;Word not learned: word contains character from prefs_ForceNewWordCharacters
               ClearWordHierarchy() 
               Return
            }
                  
            if AddWord contains %prefs_DoNotLearnStrings%
            {
               ;Word not learned: word contains string from prefs_DoNotLearnStrings
               ClearWordHierarchy() 
               Return
            }
                  
            CountValue = 1
                  
         } else {
            CountValue := prefs_LearnCount ;set the count to LearnCount so it gets written to the file
         }
         
         ; must update wordreplacement since SQLLite3 considers nulls unique
         g_WordListDB.Query("INSERT INTO words (wordindexed, word, count, wordreplacement) VALUES ('" . AddWordIndexTransformed . "','" . AddWordTransformed . "','" . CountValue . "','');")
      } else IfEqual, prefs_LearnMode, On
      {
         IfEqual, ForceCountNewOnly, 1                     
         {
            For each, row in AddWordInList.Rows
            {
               CountValue := row[3]
               break
            }
               
            IF ( CountValue < prefs_LearnCount )
            {
               ;Update the 'lastused' field as well
               g_WordListDB.QUERY("UPDATE words SET count = ('" . prefs_LearnCount . "'), lastused = (select datetime(strftime('%s','now'), 'unixepoch')) WHERE word = '" . AddWordTransformed . "';")
            }
         } else {
            UpdateWordCount(AddWord,0) ;Increment the word count if it's already in the list and we aren't forcing it on
         }
      }
	  
      ;Since this section can only be reached by the above two conditions, this means that the word was either new and added; or it was
      ;previous known and updated in the list.
      ;Time to record the word hierarchy (parent-child relationship) information:
	  
      UpdateWordHierarchy(AddWordTransformed)
   }
   
   Return
}

CheckValid(Word,ForceLearn=false)
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
   
   ;Anything below this line should not be checked if we want to Force Learning the word (Ctrl-Shift-C or coming from wordlist.txt)
   If ForceLearn
      Return, 1
   
   ;if Word does not contain at least one alpha character, skip out.
   IfEqual, A_IsUnicode, 1
   {
      if ( RegExMatch(Word, "S)\pL") = 0 )  
      {
         return
      }
   } else if ( RegExMatch(Word, "S)[a-zA-Zà-öø-ÿÀ-ÖØ-ß]") = 0 )
   {
      Return
   }
   
   Return, 1
}

TransformWord(AddWord, AddWordReplacement, AddWordDescription, ByRef AddWordTransformed, ByRef AddWordIndexTransformed, ByRef AddWordReplacementTransformed, ByRef AddWordDescriptionTransformed)
{
   AddWordIndex := AddWord
   
   ; normalize accented characters
   AddWordIndex := StrUnmark(AddWordIndex)
   
   StringUpper, AddWordIndex, AddWordIndex
   
   StringReplace, AddWordTransformed, AddWord, ', '', All
   StringReplace, AddWordIndexTransformed, AddWordIndex, ', '', All
   if (AddWordReplacement) {
      StringReplace, AddWordReplacementTransformed, AddWordReplacement, ', '', All
   }
   if (AddWordDescription) {
      StringReplace, AddWordDescriptionTransformed, AddWordDescription, ', '', All
   }
}

DeleteWordFromList(DeleteWord)
{
   global prefs_LearnMode
   global g_WordListDB
   
   Ifequal, DeleteWord,  ;If we have no word to delete, skip out.
      Return
            
   if DeleteWord is space ;If DeleteWord is only whitespace, skip out.
      Return
   
   IfNotEqual, prefs_LearnMode, On
      Return
   
   StringReplace, DeleteWordEscaped, DeleteWord, ', '', All
   g_WordListDB.Query("DELETE FROM words WHERE word = '" . DeleteWordEscaped . "';")
      
   Return   
}

;------------------------------------------------------------------------

UpdateWordCount(word,SortOnly)
{
   global prefs_LearnMode
   global g_WordListDB
   ;Word = Word to increment count for
   ;SortOnly = Only sort the words, don't increment the count
   
   ;Should only be called when LearnMode is on  
   IfEqual, prefs_LearnMode, Off
      Return
   
   IfEqual, SortOnly, 
      Return

   StringReplace, wordEscaped, word, ', '', All
   ;Update word count AND the lastused time for this word
   g_WordListDB.Query("UPDATE words SET count = count + 1, lastused = (select datetime(strftime('%s','now'), 'unixepoch')) WHERE word = '" . wordEscaped . "';")

   
   Return
}

;------------------------------------------------------------------------

UpdateWordHierarchyCount(word)
{
   global prefs_LearnMode
   global g_WordListDB
   global g_Word_Minus1
   ;Word = Word to increment count for
   
   ;Should only be called when LearnMode is on  
   IfEqual, prefs_LearnMode, Off
      Return

   StringReplace, wordEscaped, word, ', '', All
   StringReplace, wordMinus1Escaped, g_Word_Minus1, ', '', All
   g_WordListDB.Query("UPDATE WordRelations SET count = count + 1, lastused = (select datetime(strftime('%s','now'), 'unixepoch')) WHERE word = (select ID from words where word = '" . wordEscaped . "') and word_minus1 = (select ID from words where word = '" . wordMinus1Escaped . "');")   
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
      g_WordListDB.Query("DELETE FROM Words WHERE count < " . prefs_LearnCount . " AND count IS NOT NULL and lastused < DATE('now','-7 day');")
   } else {
      g_WordListDB.Query("DELETE FROM Words WHERE count < " . prefs_LearnCount . " OR count IS NULL;")
   }

   ;Zap all lastused values greater than a threshold (anything last used earlier than today) to reduce dataset size for lastused functionality
   g_WordListDB.Query("UPDATE words SET lastused = 0 where lastused < DATE('now','-7 day') and lastused <> 0;")
   
   ;Remove all word relationship that are under the threshold if older than 7 days.
   g_WordListDB.Query("DELETE FROM WordRelations WHERE count < " . prefs_LearnCount . " AND lastused < DATE ('now','-7 day');")
   
   ;Remove all words relationships that are not in the word list [not really required due to constraint on table]
   g_WordListDB.Query("DELETE FROM WordRelations WHERE word NOT IN (SELECT ID FROM words) OR word_minus1 NOT IN (SELECT ID FROM words);")

   Progress, Off
}

;------------------------------------------------------------------------

MaybeUpdateWordlist()
{
   global g_LegacyLearnedWords
   global g_WordListDB
   global g_WordListDone
   global prefs_LearnCount
   
   ; Update the Learned Words
   IfEqual, g_WordListDone, 1
   {
      
      SortWordList := g_WordListDB.Query("SELECT Word FROM Words WHERE count >= " . prefs_LearnCount . " AND count IS NOT NULL ORDER BY count DESC;")
      
      for each, row in SortWordList.Rows
      {
         TempWordList .= row[1] . "`r`n"
      }
      
      If ( SortWordList.Count() > 0 )
      {
         StringTrimRight, TempWordList, TempWordList, 2
   
         FileDelete, %PathToUserFiles%\Temp_WordlistLearned.txt
         FileAppendDispatch(TempWordList, PathToUserFiles . "\Temp_WordlistLearned.txt")
         FileCopy, %PathToUserFiles%\Temp_WordlistLearned.txt, %PathToUserFiles%\WordlistLearned.txt, 1
         FileDelete, %PathToUserFiles%\Temp_WordlistLearned.txt
         
         ; Convert the Old Wordlist file to not have ;LEARNEDWORDS;
         IfEqual, g_LegacyLearnedWords, 1
         {
            TempWordList =
            FileRead, ParseWords, %PathToUserFiles%\Wordlist.txt
            LearnedWordsPos := InStr(ParseWords, "`;LEARNEDWORDS`;",true,1) ;Check for Learned Words
            TempWordList := SubStr(ParseWords, 1, LearnedwordsPos - 1) ;Grab all non-learned words out of list
            ParseWords = 
            FileDelete, %PathToUserFiles%\Temp_Wordlist.txt
            FileAppendDispatch(TempWordList, PathToUserFiles . "\Temp_Wordlist.txt")
            FileCopy, %PathToUserFiles%\Temp_Wordlist.txt, %PathToUserFiles%\Wordlist.txt, 1
            FileDelete, %PathToUserFiles%\Temp_Wordlist.txt
         }   
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
;------------------------------------------------------------------------

BulkLearnFromClipboard(textblock)
{
   global g_TerminatingCharactersParsed
   
   ;Display progress bar window...
   Progress, M, Please wait..., Bulk learning..., %g_ScriptTitle%

   ;Count how many individual items there are, we need this number to display
   ;an accurate progress bar.
   Loop, Parse, textblock, %g_TerminatingCharactersParsed%`r`n%A_Tab%%A_Space%
   {
      ;Count the individual items
      Counter++
   }
   
   Loop, Parse, textblock, %g_TerminatingCharactersParsed%`r`n%A_Tab%%A_Space%
   {
      ;Display words to show progress...
      ProgressPercent := Round(A_Index/Counter * 100)
      Progress, %ProgressPercent%, Please wait..., %A_LoopField%, %g_ScriptTitle%
      AddWordToList(A_LoopField, 0,"ForceLearn")
   }
   
   ;Turn off progress bar window...
   Progress, Off
   
   return
}
