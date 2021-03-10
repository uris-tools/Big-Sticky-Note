/*
    Big Sticky Note


    This utility is a sticky note:  Instead of using Post-It or other notes, this utility is always available
    and is always with me, even if I work from home or a coffee shop.

    To open the note, Click Win-N.  To close, hit ESC.

    The tool also supports reminders:   Any line that contains a "@Date" or "@Time" indication will 
    notify the user using Tray Tip, at the given day (every 30 minutes) or the given hour (every day,
    twice during the hour).  To use, enter a reminder time like this:
                @Monday Call John
                @16:00 Daily summary  (Note that the reminder is on the hour, not the exact time.  It might 
                                        alert anytime between 16:00 and 17:00)
                @23/03 Send my report   
				@today do something (will remind every hour)

    The note is automatically saved in a text file.  10 historical copies are saved, so you can manually
    revert to older versions, by renaming the file.

    Version History

        1.2 - Addd alerts
        1.3 - Cleanup
		1.4 - fixed Icon
*/

global VERSION:="1.4"


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
#SingleInstance, force
SetBatchLines, -1
#Warn  All, MsgBox
#UseHook

#include <AutoXYWH> 

;Save the Note file for the compiled version.
FileInstall, Note.txt, Note.txt, 0

try menu, tray, Icon,StickyNote.ico
Menu, Tray, NoStandard
Menu, Tray, Tip, Big Sticky Note %VERSION%
Menu, Tray, add, Note, showNoteScreen
Menu, tray, Default, Note
Menu, Tray, add, Exit, closeApp



;--------------------------------------------------------------------------------------------------------------
;

global NOTE_FILE:="NOTE.TXT"
global NOTES_INI_FILE:="StickyNote.ini"

global hwnd_noteGUI
global NoteEditBox
global noteContent

global noteWidth
global noteHeight 
global noteX 
global noteY 

;cursor position inside the note (num of character) used to return to same point after reload.
global cursorPositionInNote:=0


log("Sticky Note Version " VERSION " Starting. ")


; Read note text file
FileRead, NoteContent, %NOTE_FILE%
; Read saved window position and size
IniRead,noteWidth,  %NOTES_INI_FILE%, General, noteWidth,1000
IniRead,noteHeight,  %NOTES_INI_FILE%, General, noteHeight,400
IniRead,noteX,  %NOTES_INI_FILE%, General, noteX,50
IniRead,noteY,  %NOTES_INI_FILE%, General, noteY,50
noteWidth:=Max(noteWidth,100)
noteHeight:=Max(noteHeight,100)
if (noteHeight=="" or noteWidth=="") {
    noteWidth:=300
    noteHeight:=300
    noteX := 20
    noteY := 20
}
log("Read note size " noteWidth "*" noteHeight " in pos " noteX "," noteY)

createNoteScreen()

SetTimer,remindersNotify,% 29*60*1000 ; check for alerts every 29 minutes


;------------------------------------------------------------------
;Check if this is the first time running
IniRead,firstTimeRunning, %NOTES_INI_FILE%, General, firstTimeRunning,Yes
if (firstTimeRunning=="Yes") {
    IniWrite, No, %NOTES_INI_FILE%, General, firstTimeRunning
    MsgBox, 4,Big Sticky Note,Welcome to Big Sticky Note !`n`nWould you BSN to  when you login (so that you could use it later?)`n`nIf you say No, you will need to start the EXE every time you login.
    IfMsgBox, Yes 
    {
        FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%A_ScriptName%.lnk, %A_ScriptDir% 
    }
    showNoteScreen()
}
;------------------------------------------------------------------

;To do on startup:
;showNoteScreen()
;remindersNotify()

;----------------------------------------------------------
; create (but don't show) the note screen
;
createNoteScreen() {

    global NoteEditBox:=
    global noteContent
    global hwnd_noteGUI ,  hwnd_editBox ,  hwnd_windowTitle , hwnd_statusLine

    global noteX, noteY, noteHeight, noteWidth
    
    global cursorPositionInNote

    ;0x400000=WS_DLGFRAME
    ;WS_THICKFRAME	0x40000
    Gui, NoteScreen:New, -Caption +LastFound +OwnDialogs +ToolWindow RESIZE +MinSize200x150 +Hwndhwnd_noteGUI,myNotesPage
    Gui, NoteScreen:Color, FFFFA5, FFFFA5
    log("hwnd_noteGUI=" hwnd_noteGUI)

    Gui, Font, S14 CDefault , Verdana
    Gui, NoteScreen:Add, Text, w%noteWidth% vhwnd_windowTitle gdragWindowByTitle BackgroundTrans 0x200,% "     Big Sticky Note"

    Gui, Font, S12 CDefault , Verdana
    Gui, NoteScreen:Add, Edit,% " w" (noteWidth-28) " h" (noteHeight-80) " vNoteEditBox WantTab Wrap hwndhwnd_editBox", %noteContent%
    Gui, Font, S8 CDefault , Verdana
    Gui, NoteScreen:Add, Text, w%noteWidth% vhwnd_statusLine gdragWindowByTitle BackgroundTrans 0x200,% "Windows-N to open.  Use @Monday, @13/10 , or @16:00 for reminders.   Version " VERSION
    
    GuiControl, Focus, NoteEditBox
    SendMessage, 0xB1, %cursorPositionInNote%, %cursorPositionInNote%,,  ahk_id %hwnd_editBox% ; ; EM_SETSEL
    OnMessage( 0x111, "HandleFocusEvents" )
   
    return

    ;Closing the windos. Save the note
    NoteScreenGuiEscape:
    NoteScreenGuiClose:
		saveWindowPosition() 
		GUI, NoteScreen:Default
		Gui, Submit
		saveNoteFile()
		Gui,hide
		return

	
    ;Adapt to Window resize
	NoteScreenGuiSize:
		If (A_EventInfo = 1) ; The window has been minimized.
			Return
		AutoXYWH("hw", "NoteEditBox")	
		AutoXYWH("y", "hwnd_statusLine")	
        
        noteWidth:=A_GuiWidth
        noteHeight:=A_GuiHeight
		return

    ;Move the window when dragged
    dragWindowByTitle:
        PostMessage, 0xA1, 2
        return
}

;Show the note screen, when the hotkey is pressed.
showNoteScreen() {
    global noteX, noteY, noteHeight, noteWidth
    global NoteEditBox
	
    GUI, NoteScreen:Default
    
    Gui, NoteScreen:Show, Y%noteY% X%noteX% w%noteWidth% h%noteHeight%

    ;Activate Auto-saved every 30 sec.   The timer will be turned off when the window is closed.
    setTimer,autoSave,30000 

}


/*
  Auto save note every 30 seconds, if it has changed, and the GUI is open.
  If the GUI is closed, cancel the timer.
*/
autoSave() {
    global hwnd_noteGUI
    log("Autosave - hwnd_noteGUI=" hwnd_noteGUI)
    if WinActive("ahk_id " hwnd_noteGUI) {
        saveNoteFile()
        saveWindowPosition()
        SplashTextOn,,,Auto Saved,
        Sleep,500
        SplashTextOff
        log("Autosave-saved")

    } else {
        ; Note is not open, cancel the timer
        setTimer,, off
        log("Autosave-cancel timer")
    }
}


/*
 Save the content of the note file.  Also, keep 10 copies by rolling over the historical files.
 Save also the cursor position in the note, to be able to return to it next time.
*/
saveNoteFile() {

    global NOTE_FILE
    global noteContent
    global NoteEditBox
    global hwnd_editBox
    global cursorPositionInNote
    	
    Gui, NoteScreen:Submit, NoHide

    ; check if note content has changed
    if (noteContent==NoteEditBox) {
        log("Content note not changed. not saving")
        return
    }
    Loop 10 {
        _x:=11-A_Index 
        if (FileExist(NOTE_FILE "." _x)) {
            _xx:=_x+1
            FileMove, %NOTE_FILE%.%_x%, %NOTE_FILE%.%_xx%, 1
        }    
    }
    if (FileExist(NOTE_FILE)) {    
            FileMove, %NOTE_FILE%, %NOTE_FILE%.1, 1 
    }
    
    log("Saving")
    fileHandle:=FileOpen(NOTE_FILE, "w")       
    fileHandle.write(NoteEditBox)
    fileHandle.close()

    noteContent:=NoteEditBox
    
    Edit_GetSelection(cursorPositionInNote,endOfSelection, , "ahk_id" hwnd_editBox )
    log("current character in note=" cursorPositionInNote)
    IniWrite, %cursorPositionInNote%, %NOTES_INI_FILE%, General, cursorPosition
    
}


 

log(message) {
	FormatTime, now , , HH:mm:ss
    FileAppend,%now% - %message%`r`n, StickyNote.log
}


; Gets the start and end offset of the current selection.
;
Edit_GetSelection(ByRef start, ByRef end, Control="", WinTitle="") { 
    VarSetCapacity(start, 4), VarSetCapacity(end, 4)
    SendMessage, 0xB0, &start, &end, %Control%, %WinTitle%  ; EM_GETSEL
    if (ErrorLevel="FAIL")
        return false
    start := NumGet(start), end := NumGet(end)
    return true
}



/*
    When closing the window, save it's size and position, to return to it next time.
*/
saveWindowPosition() {
    global noteX, noteY, noteHeight, noteWidth
    global hwnd_noteGUI
	
	WinGetPos,noteX,noteY,,,ahk_id %hwnd_noteGUI%
	
	log("HWND " hwnd_noteGUI " - Saving note size " noteWidth "*" noteHeight " in pos " noteX "," noteY)

	IniWrite, %noteWidth%, %NOTES_INI_FILE%, General, noteWidth
	IniWrite, %noteHeight%, %NOTES_INI_FILE%, General, noteHeight
	IniWrite, %noteX%, %NOTES_INI_FILE%, General, noteX
	IniWrite, %noteY%, %NOTES_INI_FILE%, General, noteY
		
	return

}


/******************************************************************************
  remindersNotify() - Go over the note file, and identify @reminders.  If it's time, notify the user
  Timer is ran every hour.
******************************************************************************
*/
remindersNotify() {
    global NOTE_FILE
    
	static dayAndMonth := "(\d{1,2})[^a-zA-Z0-9:.]+(\d{1,2})"

    FileRead, NoteContent, %NOTE_FILE%
    log("Checking for reminders")
    Loop, read, %NOTE_FILE% 
    {
        FoundPos:=RegExMatch(A_LoopReadLine,"O)@(\S*)" , obj)
        if (obj.Count()>0) {
            timeIndication:=obj.value(1)        

            ;a "@xxxxx" was found.  Try different formats, to determine the time it represents.
            parsedTime:=""
            parsedDate:=""
            today:=A_YYYY "-" A_MM "-" A_DD

            ;10:30   10:30pm  22:30
            if (RegExMatch(timeIndication
                        , "i)(\d{1,2})"					;hours
                            . ":(\d{1,2})"		;minutes
                            . "(?:\s*([ap]m))?"
                        , t)) {
                h := (t3=="pm" ? t1+12 : t1)
                m := t2
                parsedTime:= h ":" m 
            }

            if (!parsedTime) {
                ;Not time indication. check for Date indication

                ;Day (Sun/Mon/Tue)
                weekday:=InSTR("SuMoTuWeThFrSa",Substr(timeIndication,1,2),CaseSensitive := false)
                if (weekday) {
                    weekday:=round((weekday+1)/2)
                    log("Weekday for " timeIndication " is " weekday ", today is " A_WDay)
                    if (weekday==A_WDay) {
                        parsedDate:=today
                    }
                }
				if(Substr(timeIndication,1,3)="tod") {
					parsedDate:=today
				}
                

                ;31/12 (without year)
                If (Regexmatch(timeIndication, "i)" . dayAndMonth, d)) {   
                    y := A_YYYY
                    d := (StrLen(d1) == 1 ? "0" . d1 : d1)
                    m := (StrLen(d2) == 1 ? "0" . d2 : d2)
                    parsedDate := y "-" m "-" d
                }

                ;31/12/2014 or 2013/31/12
                if (!parsedDate) {
                    If Regexmatch(timeIndication, "i)(\d{4})[^a-zA-Z0-9:.]+" . dayAndMonth, d) {   ;2004/22/03
                        y := (StrLen(d1) == 2 ? "20" . d1: d1)
                        m := (StrLen(d3) == 1 ? "0" . d3 : d3)
                        d := (StrLen(d2) == 1 ? "0" . d2 : d2)
                        parsedDate := y "-" m "-" d
                    }
                    Else If Regexmatch(timeIndication                ;22/03/2004 or 22/03/04
                                        , "i)" 
                                            . dayAndMonth 
                                            . "(?:[^a-zA-Z0-9:.]+((?:\d{4}|\d{2})))?"
                                        , d) {  
                        y := (StrLen(d3) == 2 ? "20" . d3: d3)
                        m := (StrLen(d2) == 1 ? "0" . d2 : d2)
                        d := (StrLen(d1) == 1 ? "0" . d1 : d1)
                        parsedDate := y "-" m "-" d
                    }
                }
            }


            if (parsedTime or parsedDate) {
                log("reminders:                  Parsed date: " parsedDate ", time " parsedTime)
            }
            
            if (parsedDate==today) {
                log("REMINDER TODAY: " A_LoopReadLine)
                TrayTip,Big Sticky Note, % A_LoopReadLine   
            }
            if(Substr(parsedTime,1,2)==A_Hour)  {
                log("REMINDER Now: " A_LoopReadLine)                
                TrayTip,Big Sticky Note, % A_LoopReadLine   
            }
        }
    }
}

closeApp() {
		saveNoteFile()
		ExitApp 
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;END OF THE SCRIPT - Hotkey definitions

#n::
	log("BigStickyNote - Hotkey pressed")
	if (WinActive("ahk_id" hwnd_noteGUI)) {
		saveWindowPosition() 
		GUI, NoteScreen:Default
		Gui, NoteScreen:Submit
		saveNoteFile()
		Gui, NoteScreen:hide
		
	} else {
		showNoteScreen()
	}
	return
	;


#If WinActive("myNotesPage")
	;Ctrl-R Reloads the script, but only when GUI active
	^r::
		saveWindowPosition() 
        saveNoteFile()
		Reload  
		return
	;Alt-Ctrl-X exists, only when GUI Active
	^!x::
		saveNoteFile()
		ExitApp 
		return
#If

