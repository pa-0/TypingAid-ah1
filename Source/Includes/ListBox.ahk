;These functions and labels are related to the shown list of words

InitializeListBox()
{
   global
   
   Gui, ListBoxGui: -DPIScale -Caption +AlwaysOnTop +ToolWindow +Delimiter%g_DelimiterChar%
   
   Local ListBoxFont
   if (prefs_ListBoxFontOverride && prefs_ListBoxFontOverride != "<Default>")
   {
      ListBoxFont := prefs_ListBoxFontOverride
   } else IfEqual, prefs_ListBoxFontFixed, On
   {
      ListBoxFont = Courier New
   } else {
      ListBoxFont = Tahoma
   }
      
   Gui, ListBoxGui:Font, s%prefs_ListBoxFontSize%, %ListBoxFont%

   Loop, %prefs_ListBoxRows%
   {
      GuiControl, ListBoxGui:-Redraw, g_ListBox%A_Index%
      ;can't use a g-label here as windows sometimes passes the click message when spamming the scrollbar arrows
      Gui, ListBoxGui: Add, ListBox, vg_ListBox%A_Index% R%A_Index% X0 Y0 T%prefs_ListBoxFontSize% T32 hwndg_ListBoxHwnd%A_Index%
   }

   Return
}

; AlexF Process click or double click into list of words (matches)
ListBoxClickItem(wParam, lParam, msg, ClickedHwnd)
{
   global
   Local NewClickedItem
   Local TempRows
   static LastClickedItem  ; AlexF index of the selected word
   
   TempRows := GetRows()
   
   if (ClickedHwnd != g_ListBoxHwnd%TempRows%)
   {
      return
   }
   
   ; if we clicked in the scrollbar, jump out
   if (A_GuiX > (g_ListBoxPosX + g_ListBoxContentWidth))
   {
      SetSwitchOffListBoxTimer()
      Return
   }
   
   GuiControlGet, g_MatchPos, ListBoxGui:, g_ListBox%TempRows%
   
   if (msg == g_WM_LBUTTONUP)
   {
      if prefs_DisabledAutoCompleteKeys not contains L
      {
         SwitchOffListBoxIfActive()
         EvaluateUpDown("$LButton")   
      } else {
         ; Track this to make sure we're double clicking on the same item
         NewClickedItem := g_MatchPos
         SetSwitchOffListBoxTimer()
      }
         
   } else if (msg == g_WM_LBUTTONDBLCLK)
   {
      SwitchOffListBoxIfActive()
      
      if prefs_DisabledAutoCompleteKeys contains L
      {
         if (LastClickedItem == g_MatchPos)
         {
            EvaluateUpDown("$LButton")   
         }
      }
   } else {
      SwitchOffListBoxIfActive()
   }
      
   ; clear or set LastClickedItem
   LastClickedItem := NewClickedItem
   
   Return
}

SetSwitchOffListBoxTimer()
{
   static DoubleClickTime
   
   if !(DoubleClickTime)
   {
      DoubleClickTime := DllCall("GetDoubleClickTime")
   }
   ;When single click is off, we have to wait for the double click time to pass
   ; before re-activating the edit window to allow double click to work
   SetTimer, SwitchOffListBoxIfActiveSub, -%DoubleClickTime%
}
   

SwitchOffListBoxIfActiveSub:
SwitchOffListBoxIfActive()
Return

ListBoxScroll(Hook, Event, EventHwnd)
{
   global
   
   Local MatchEnd
   Local SI
   Local TempRows
   Local Position
   
   if (g_ListBox_Id)
   {
   
      TempRows := GetRows()
      if (g_ListBoxHwnd%TempRows% != EventHwnd)
      {
         return
      }
      
      if (Event == g_EVENT_SYSTEM_SCROLLINGSTART)
      {
         ; make sure the timer is clear so we don't switch while scrolling
         SetTimer, SwitchOffListBoxIfActiveSub, Off
         return
      }
      
      SI:=GetScrollInfo(g_ListBoxHwnd%TempRows%)
   
      if (!SI.npos)
      {
         return
      }
   
      if (SI.npos == g_MatchStart)
      {
         return
      }
   
      g_MatchStart := SI.npos
   
      SetSwitchOffListBoxTimer()   
   }
}

; based on code by HotKeyIt
;  http://www.autohotkey.com/board/topic/78829-ahk-l-scrollinfo/
;  http://www.autohotkey.com/board/topic/55150-class-structfunc-sizeof-updated-010412-ahkv2/
GetScrollInfo(ctrlhwnd) {
  global g_SB_VERT
  global g_SIF_POS
  SI:=new _Struct("cbSize,fMask,nMin,nMax,nPage,nPos,nTrackPos")
  SI.cbSize:=sizeof(SI)
  SI.fMask := g_SIF_POS
  If !DllCall("GetScrollInfo","PTR",ctrlhwnd,"Int",g_SB_VERT,"PTR",SI[""])
    Return false
  else Return SI
}

ListBoxChooseItem(Row)
{
   global
   GuiControl, ListBoxGui: Choose, g_ListBox%Row%, %g_MatchPos%
}

;------------------------------------------------------------------------

CloseListBox()
{
   global g_ListBox_Id
   IfNotEqual, g_ListBox_Id,
   {
      Gui, ListBoxGui: Hide
      ListBoxEnd()
   }
   Return
}

DestroyListBox()
{
   Gui, ListBoxGui:Destroy
   ListBoxEnd()
   Return
}

ListBoxEnd()
{
   global g_ScrollEventHook
   global g_ScrollEventHookThread
   global g_ListBox_Id
   global g_WM_LBUTTONUP
   global g_WM_LBUTTONDBLCLK
   
   g_ListBox_Id =
   
   OnMessage(g_WM_LBUTTONUP, "")
   OnMessage(g_WM_LBUTTONDBLCLK, "")

   if (g_ScrollEventHook) {
      DllCall("UnhookWinEvent", "Uint", g_ScrollEventHook)
      g_ScrollEventHook =
      g_ScrollEventHookThread =
      MaybeCoUninitialize()
   }
   DisableKeyboardHotKeys()
   return
}

;------------------------------------------------------------------------

SavePriorMatchPosition()
{
   global g_MatchPos         ;AlexF current highlighted position in the listbox
   global g_MatchStart       ;AlexF position of the first word (match) to be shown in the listbox
   global g_OldMatch         ;AlexF either last highlighted word or last highlighted position in the listbox, depending on prefs_ArrowKeyMethod
   global g_OldMatchStart
   global g_SingleMatchAdj   ;AlexF array of matched words, with adjusted capitalization. This is what user sees. 
   global prefs_ArrowKeyMethod
   
   if !(g_MatchPos)
   {
      g_OldMatch =
      g_OldMatchStart = 
   } else IfEqual, prefs_ArrowKeyMethod, LastWord
   {
      g_OldMatch := g_SingleMatchAdj[g_MatchPos]
      g_OldMatchStart = 
   } else IfEqual, prefs_ArrowKeyMethod, LastPosition
   {
      g_OldMatch := g_MatchPos
      g_OldMatchStart := g_MatchStart
   } else {
      g_OldMatch =
      g_OldMatchStart =
   }
      
   Return
}

; AlexF Calculates initial value of g_MatchPos - highlighted position in the listbox.
; Depending on prefs_ArrowKeyMethod, it is either the first, or last, or previously selected word.
SetupMatchPosition()
{
   global g_MatchPos         ;AlexF current highlighted position in the listbox
   global g_MatchStart       ;AlexF position of the first word (match) to be shown in the listbox
   global g_MatchTotal       ;AlexF count of matched words
   global g_OldMatch         ;AlexF either last highlighted word or last highlighted position in the listbox, depending on prefs_ArrowKeyMethod
   global g_OldMatchStart
   global g_SingleMatchAdj   ;AlexF array of matched words, with adjusted capitalization. This is what user sees. 
   global prefs_ArrowKeyMethod
   global prefs_ListBoxRows
   
   IfEqual, g_OldMatch, 
   {
      IfEqual, prefs_ArrowKeyMethod, Off
      {
         g_MatchPos = 
         g_MatchStart = 1
      } else {
         g_MatchPos = 1
         g_MatchStart = 1
      }
   } else IfEqual, prefs_ArrowKeyMethod, Off
   {
      g_MatchPos = 
      g_MatchStart = 1
   } else IfEqual, prefs_ArrowKeyMethod, LastPosition
   {
      IfGreater, g_OldMatch, %g_MatchTotal%
      {
         g_MatchStart := g_MatchTotal - (prefs_ListBoxRows - 1)
         IfLess, g_MatchStart, 1
            g_MatchStart = 1
         g_MatchPos := g_MatchTotal
      } else {
         g_MatchStart := g_OldMatchStart
         If ( g_MatchStart > (g_MatchTotal - (prefs_ListBoxRows - 1) ))
         {
            g_MatchStart := g_MatchTotal - (prefs_ListBoxRows - 1)
            IfLess, g_MatchStart, 1
               g_MatchStart = 1
         }
         g_MatchPos := g_OldMatch
      }
   
   } else IfEqual, prefs_ArrowKeyMethod, LastWord
   {
      ListPosition =
      Loop, %g_MatchTotal%
      {
         if ( g_OldMatch == g_SingleMatchAdj[A_Index] )
         {
            ListPosition := A_Index
            Break
         }
      }
      IfEqual, ListPosition, 
      {
         g_MatchPos = 1
         g_MatchStart = 1
      } Else {
         g_MatchStart := ListPosition - (prefs_ListBoxRows - 1)
         IfLess, g_MatchStart, 1
            g_MatchStart = 1
         g_MatchPos := ListPosition
      }
   } else {
      g_MatchPos = 1
      g_MatchStart = 1
   }
             
   g_OldMatch = 
   g_OldMatchStart = 
   Return
}

RebuildMatchList()
{
   global g_Match                ;AlexF concatenation of all the lines in the listbox, separated by g_DelimiterChar
   global g_MatchLongestLength
   global g_MatchPos
   global g_MatchStart           ;AlexF position of the first word (match) to be shown in the listbox
   global g_MatchTotal           ;AlexF count of matched words
   global g_OriginalMatchStart
   global prefs_ListBoxRows
   
   g_Match = 
   g_MatchLongestLength =
   
   if (!g_MatchPos)
   {
      ; do nothing
   } else if (g_MatchPos < g_MatchStart)
   {
      g_MatchStart := g_MatchPos
   } else if (g_MatchPos > (g_MatchStart + (prefs_ListBoxRows - 1)))
   {
      g_MatchStart := g_MatchPos - (prefs_ListBoxRows -1)
   }
   
   g_OriginalMatchStart := g_MatchStart
   
   MaxLength := ComputeListBoxMaxLength() ; AlexF: max count of characters (approximately?) to fit width of listbox
   
   ; Actually fill out the listbox
   Loop, %g_MatchTotal%
   {
      CurrentLength := AddToMatchList(A_Index, MaxLength, LongestBaseLength)
      IfGreater, CurrentLength, %g_MatchLongestLength%
         g_MatchLongestLength := CurrentLength      
   }
   StringTrimRight, g_Match, g_Match, 1        ; Get rid of the last linefeed 
   Return
}

;AlexF 
; Upends g_Match - a string of all the matches - with word at given position. 
;   position - index of matched word in g_SingleMatchAdj
;   MaxLength - max count of characters (approximately?) to fit width of listbox
;   LongestBaseLength - max count of characters in the matched words
; Returns number of characters (width) of the added line (I guess).
AddToMatchList(position, MaxLength, LongestBaseLength)
{
   global g_DelimiterChar
   global g_Match                ;AlexF output, concatenation of all the lines in the listbox, separated by g_DelimiterChar
   global g_MatchStart           ;AlexF position of the first word (match) to be shown in the listbox
   global g_NumKeyMethod
   global g_SingleMatchAdj       ;AlexF array of matched words, with adjusted capitalization. This is what user sees. 
   global prefs_ListBoxFontFixed
   
   blankprefix = `t
   
   IfEqual, g_NumKeyMethod, Off
   {
      prefix := blankprefix
   } else IfLess, position, %g_MatchStart%
   {
      prefix := blankprefix
   } else if ( position > ( g_MatchStart + 9 ) )
   {
      prefix := blankprefix
   } else {
      prefix := Mod(position - g_MatchStart +1,10) . "`t"
   }
   
   prefixlen := 2
   
   CurrentMatch := g_SingleMatchAdj[position]
   BaseLength := MaxLength
   
   CurrentMatchLength := StrLen(CurrentMatch) + prefixlen
   
   if (CurrentMatchLength > BaseLength)
   {
      CompensatedBaseLength := BaseLength - prefixlen
      ; remove 3 characters so we can add the ellipsis
      StringLeft, CurrentMatch, CurrentMatch, CompensatedBaseLength - 3
      CurrentMatch .= "..."
   
      CurrentMatchLength := StrLen(CurrentMatch) + prefixlen
   }
   
   Iterations := 0
   Tabs = 
   Remainder := 0
   
   g_Match .= prefix . CurrentMatch
   
   g_Match .= g_DelimiterChar
   Return, CurrentMatchLength
}

;------------------------------------------------------------------------

; AlexF: this is for WIDTH of the listbox in CHARACTERS, not HEIGHT in rows
; find out the longest length we can use in the listbox.
; Any changes to this function probably need to be reflected in ShowListBox() or ForceWithinMonitorBounds
ComputeListBoxMaxLength()
{
   global g_ListBoxCharacterWidthComputed ;AlexF average width of one character, pixels
   global g_MatchTotal ; AlexF: number of matches
   global g_SM_CMONITORS
   global g_SM_CXFOCUSBORDER
   global g_SM_CXVSCROLL
   global prefs_ListBoxMaxWidth
   
   ; grab the width of a vertical scrollbar

   Rows := GetRows() ; AlexF: number of rows
   
   IfGreater, g_MatchTotal, %Rows%
   {
      SysGet, ScrollBarWidth, %g_SM_CXVSCROLL%
      if ScrollBarWidth is not integer
         ScrollBarWidth = 17
   } else ScrollBarWidth = 0

   ; Grab the internal border width of the ListBox box
   SysGet, BorderWidthX, %g_SM_CXFOCUSBORDER%
   If BorderWidthX is not integer
      BorderWidthX = 1
   
   ;Use g_ListBoxCharacterWidthComputed pixels for each character in width
   ListBoxBaseSizeX := g_ListBoxCharacterWidthComputed + ScrollBarWidth + (BorderWidthX * 2)
   
   ListBoxPosX := HCaretX()
   ListBoxPosY := HCaretY()
   
   SysGet, NumMonitors, %g_SM_CMONITORS%

   IfLess, NumMonitors, 1
      NumMonitors =1
         
   Loop, %NumMonitors%
   {
      SysGet, Mon, Monitor, %A_Index%
      IF ( ( ListBoxPosX < MonLeft ) || (ListBoxPosX > MonRight ) || ( ListBoxPosY < MonTop ) || (ListBoxPosY > MonBottom ) )
         Continue
      
      MonWidth := MonRight - MonLeft
      break
   }
   
   if !prefs_ListBoxMaxWidth
   {
      Width := MonWidth
   } else if (prefs_ListBoxMaxWidth < MonWidth)
   {
      Width := prefs_ListBoxMaxWidth
   } else 
   {
      Width := MonWidth
   }
   
   return Floor((Width-ListBoxBaseSizeX)/ g_ListBoxCharacterWidthComputed)
}
   

;Show matched values
; Any changes to this function may need to be reflected in ComputeListBoxMaxLength()
ShowListBox()
{
   global

   IfNotEqual, g_Match,
   {
      Local BorderWidthX
      Local ListBoxActualSize
      Local ListBoxActualSizeH
      Local ListBoxActualSizeW
      Local ListBoxPosY
      Local ListBoxSizeX
      Local ListBoxThread
      Local MatchEnd
      Local Rows
      Local ScrollBarWidth
      static ListBox_Old_Cursor

      Rows := GetRows()
      
      IfGreater, g_MatchTotal, %Rows%
      {
         SysGet, ScrollBarWidth, %g_SM_CXVSCROLL%
         if ScrollBarWidth is not integer
            ScrollBarWidth = 17
      } else ScrollBarWidth = 0
   
      ; Grab the internal border width of the ListBox box
      SysGet, BorderWidthX, %g_SM_CXFOCUSBORDER%
      If BorderWidthX is not integer
         BorderWidthX = 1
      
      ;Use 8 pixels for each character in width
      ListBoxSizeX := g_ListBoxCharacterWidthComputed * g_MatchLongestLength + g_ListBoxCharacterWidthComputed + ScrollBarWidth + (BorderWidthX * 2)
      
      g_ListBoxPosX := HCaretX()
      ListBoxPosY := HCaretY()
      
      ; In rare scenarios, the Cursor may not have been detected. In these cases, we just won't show the ListBox.
      IF (!(g_ListBoxPosX) || !(ListBoxPosY))
      {
         return
      }
      
      MatchEnd := g_MatchStart + (prefs_ListBoxRows - 1)
      
      Loop, %prefs_ListBoxRows%
      { 
         IfEqual, A_Index, %Rows%
         {
            GuiControl, ListBoxGui: -Redraw, g_ListBox%A_Index%
            GuiControl, ListBoxGui: Move, g_ListBox%A_Index%, w%ListBoxSizeX%
            GuiControl, ListBoxGui: ,g_ListBox%A_Index%, %g_DelimiterChar%%g_Match%
            IfNotEqual, g_MatchPos,
            {
               GuiControl, ListBoxGui: Choose, g_ListBox%A_Index%, %MatchEnd%
               GuiControl, ListBoxGui: Choose, g_ListBox%A_Index%, %g_MatchPos%
            }
            GuiControl, ListBoxGui: +AltSubmit +Redraw, g_ListBox%A_Index%
            GuiControl, ListBoxGui: Show, g_ListBox%A_Index%
            GuiControlGet, ListBoxActualSize, ListBoxGui: Pos, g_ListBox%A_Index%
            Continue
         }
      
         GuiControl, ListBoxGui: Hide, g_ListBox%A_Index%
         GuiControl, ListBoxGui: -Redraw, g_ListBox%A_Index%
         GuiControl, ListBoxGui: , g_ListBox%A_Index%, %g_DelimiterChar%
      }
      
      ForceWithinMonitorBounds(g_ListBoxPosX,ListBoxPosY,ListBoxActualSizeW,ListBoxActualSizeH)
      
      g_ListBoxContentWidth := ListBoxActualSizeW - ScrollBarWidth - BorderWidthX
      
      IfEqual, g_ListBox_Id,
      {
         
         if prefs_DisabledAutoCompleteKeys not contains L
         {
            if (!ListBox_Old_Cursor)
            {
               ListBox_Old_Cursor := DllCall(g_SetClassLongFunction, "Uint", g_ListBoxHwnd%Rows%, "int", g_GCLP_HCURSOR, "int", g_cursor_hand)
            }
            
            DllCall(g_SetClassLongFunction, "Uint", g_ListBoxHwnd%Rows%, "int", g_GCLP_HCURSOR, "int", g_cursor_hand)
            
         ; we only need to set it back to the default cursor if we've ever unset the default cursor
         } else if (ListBox_Old_Cursor)
         {
            DllCall(g_SetClassLongFunction, "Uint", g_ListBoxHwnd%Rows%, "int", g_GCLP_HCURSOR, "int", ListBox_Old_Cursor)
         }
            
      }
      
      Gui, ListBoxGui: Show, NoActivate X%g_ListBoxPosX% Y%ListBoxPosY% H%ListBoxActualSizeH% W%ListBoxActualSizeW%, Word List Appears Here.
      Gui, ListBoxGui: +LastFound +AlwaysOnTop
      
      IfEqual, g_ListBox_Id,
      {
         
         EnableKeyboardHotKeys()   
      }
      
      WinGet, g_ListBox_Id, ID, Word List Appears Here.
      
      ListBoxThread := DllCall("GetWindowThreadProcessId", "Ptr", g_ListBox_Id, "Ptr", g_NULL)
      if (g_ScrollEventHook && (ListBoxThread != g_ScrollEventHookThread))
      {
         DllCall("UnhookWinEvent", "Uint", g_ScrollEventHook)
         g_ScrollEventHook =
         g_ScrollEventHookThread =
         MaybeCoUninitialize()
      }
         
      if (!g_ScrollEventHook) {
         MaybeCoInitializeEx()
         g_ScrollEventHook := DllCall("SetWinEventHook", "Uint", g_EVENT_SYSTEM_SCROLLINGSTART, "Uint", g_EVENT_SYSTEM_SCROLLINGEND, "Ptr", g_NULL, "Uint", g_ListBoxScrollCallback, "Uint", g_PID, "Uint", ListBoxThread, "Uint", g_NULL)
         g_ScrollEventHookThread := ListBoxThread
      }
      
      OnMessage(g_WM_LBUTTONUP, "ListBoxClickItem")
      OnMessage(g_WM_LBUTTONDBLCLK, "ListBoxClickItem")
      
      IfNotEqual, prefs_ListBoxOpacity, 255
         WinSet, Transparent, %prefs_ListBoxOpacity%, ahk_id %g_ListBox_Id%
   }
}

; Any changes to this function may need to be reflected in ComputeListBoxMaxLength()
ForceWithinMonitorBounds(ByRef ListBoxPosX,ByRef ListBoxPosY,ListBoxActualSizeW,ListBoxActualSizeH)
{
   global g_ListBoxFlipped
   global g_SM_CMONITORS
   global g_ListBoxCharacterWidthComputed ; AlexF average width of one character, pixels
   global g_ListBoxOffsetComputed
   global g_ListBoxMaxWordHeight
   ;Grab the number of non-dummy monitors
   SysGet, NumMonitors, %g_SM_CMONITORS%
   
   IfLess, NumMonitors, 1
      NumMonitors =1
         
   Loop, %NumMonitors%
   {
      SysGet, Mon, Monitor, %A_Index%
      IF ( ( ListBoxPosX < MonLeft ) || (ListBoxPosX > MonRight ) || ( ListBoxPosY < MonTop ) || (ListBoxPosY > MonBottom ) )
         Continue
      
      if (ListBoxActualSizeH > g_ListBoxMaxWordHeight) {
         g_ListBoxMaxWordHeight := ListBoxActualSizeH
      }
      
      ; + g_ListBoxOffsetComputed Move ListBox down a little so as not to hide the caret. 
      ListBoxPosY := ListBoxPosY + g_ListBoxOffsetComputed
      if (g_ListBoxFlipped) {
         ListBoxMaxPosY := HCaretY() - g_ListBoxMaxWordHeight
         
         if (ListBoxMaxPosY < MonTop) {
            g_ListBoxFlipped =
         } else {
            ListBoxPosY := HCaretY() - ListBoxActualSizeH
         }
      }
      
      ; make sure we don't go below the screen.
      If ( (ListBoxPosY + g_ListBoxMaxWordHeight ) > MonBottom )
      {
         ListBoxPosY := HCaretY() - ListBoxActualSizeH
         g_ListBoxFlipped := true
      }
      
      ; make sure we don't go above the top of the screen.
      If (ListBoxPosY < MonTop) {
         ListBoxPosY := MonTop
         ; Try to move over horizontal position to leave some space, may get overridden later.
         ListBoxPosX += g_ListBoxCharacterWidthComputed
      }
      
      If ( (ListBoxPosX + ListBoxActualSizeW ) > MonRight )
      {
         ListBoxPosX := MonRight - ListBoxActualSizeW
         If ( ListBoxPosX < MonLeft )
            ListBoxPosX := MonLeft
      }
         
         
      Break
   }

   Return      
}

;------------------------------------------------------------------------

; Returns count of rows in the to-be-shown listbox. AlexF.
GetRows()
{
   global g_MatchTotal
   global prefs_ListBoxRows
   IfGreater, g_MatchTotal, %prefs_ListBoxRows%
      Rows := prefs_ListBoxRows
   else Rows := g_MatchTotal
   
   Return, Rows
}
;------------------------------------------------------------------------

; function to grab the X position of the caret for the ListBox
HCaretX() 
{
   global g_DpiAware
   global g_DpiScalingFactor
   global g_Helper_Id
   global g_Process_DPI_Unaware
    
   WinGetPos, HelperX,,,, ahk_id %g_Helper_Id% 
   if HelperX !=
   { 
      return HelperX
   } 
   if ( CheckIfCaretNotDetectable() )
   { 
      chromeWindow := WinExist("ahk_class Chrome_WidgetWin_1")
      if (chromeWindow) {
         Acc_Caret := Acc_ObjectFromWindow(chromeWindow, OBJID_CARET := 0xFFFFFFF8)
         Caret_Location := Acc_Location(Acc_Caret)
         if Caret_Location.x {
            return Caret_Location.x
         }
      }
      else {
         MouseGetPos, MouseX
         return MouseX
      }
   }
   ; non-DPI Aware
   if (g_DpiAware == g_Process_DPI_Unaware) {
      return (A_CaretX * g_DpiScalingFactor)
   }
   
   return A_CaretX 
} 

;------------------------------------------------------------------------

; function to grab the Y position of the caret for the ListBox
HCaretY() 
{
   global g_DpiAware
   global g_DpiScalingFactor
   global g_Helper_Id
   global g_Process_DPI_Unaware

   WinGetPos,,HelperY,,, ahk_id %g_Helper_Id% 
   if HelperY != 
   { 
      return HelperY
   } 
   if ( CheckIfCaretNotDetectable() )
   { 
      chromeWindow := WinExist("ahk_class Chrome_WidgetWin_1")
      if (chromeWindow) {
         Acc_Caret := Acc_ObjectFromWindow(chromeWindow, OBJID_CARET := 0xFFFFFFF8)
         Caret_Location := Acc_Location(Acc_Caret)
         if Caret_Location.y {
            return Caret_Location.y
         }
      }
      else {
         MouseGetPos, , MouseY
         return MouseY + (20*g_DpiScalingFactor)
      }
   }
   if (g_DpiAware == g_Process_DPI_Unaware) {
      return (A_CaretY * g_DpiScalingFactor)
   }
   
   return A_CaretY 
}

;------------------------------------------------------------------------

; functions to fix issue with suggestions list not following caret in chrome based apps
; Based on https://www.autohotkey.com/boards/viewtopic.php?t=67923
Acc_ObjectFromWindow(hWnd, idObject = -4)
{
    Acc_Init()
    If  DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
    Return  ComObjEnwrap(9,pacc,1)
}

Acc_Init()
{
    Static  h
    If Not  h
        h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}

Acc_Location(Acc, ChildId=0, byref Position="") {
    try Acc.accLocation(ComObj(0x4003,&x:=0), ComObj(0x4003,&y:=0), ComObj(0x4003,&w:=0), ComObj(0x4003,&h:=0), ChildId)
    catch
        return
    Position := "x" NumGet(x,0,"int") " y" NumGet(y,0,"int") " w" NumGet(w,0,"int") " h" NumGet(h,0,"int")
    return  {x:NumGet(x,0,"int"), y:NumGet(y,0,"int"), w:NumGet(w,0,"int"), h:NumGet(h,0,"int")}
}

;------------------------------------------------------------------------

CheckIfCaretNotDetectable()
{
   ;Grab the number of non-dummy monitors
   SysGet, NumMonitors, 80
   
   IfLess, NumMonitors, 1
      NumMonitors = 1
   
   if !(A_CaretX)
   {
      Return, 1
   }
   
   ;if the X caret position is equal to the leftmost border of the monitor +1, we can't detect the caret position.
   Loop, %NumMonitors%
   {
      SysGet, Mon, Monitor, %A_Index%
      if ( A_CaretX = ( MonLeft ) )
      {
         Return, 1
      }
      
   }
   
   Return, 0
}