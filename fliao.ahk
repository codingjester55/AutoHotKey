#SingleInstance force

;///////////////////////////////////////////////////////////////////////////////
; ! = Alt
; + = Shift
; ^ = Ctrl
; # = Win
;///////////////////////////////////////////////////////////////////////////////


;//////////////////////////////////////////////////////////////////////////////
;////////////////////////// REMAPPING KEYs ////////////////////////////////////
;
;maps CTRL+SHIFT+n to numlock (always turns it off)
^+n::
Send {NumLock}
SetNumLockState, Off
Return
;
;maps CAPLOCK to CTRL space for Launchy
CAPSLOCK::Send, {CTRLDOWN}{SPACE}{CTRLUP}
;
;Re-map Outlook quick-tasks to drop the SHIFT
IfWinActive, ahk_class rctrl_renwnd32
{
^1:: ;PUT THE KEYSTROKE YOU WANT TO USE HERE - IN THIS CASE Ctrl + 1.
Send {Shift down}{Ctrl down}1 ;THIS WILL FIRE WHICHEVER QUICKSTEP IS MAPPED TO CTRL+SHIFT+1.
Send {Shift up}{Ctrl up}
return

^2:: 
Send {Shift down}{Ctrl down}2 
Send {Shift up}{Ctrl up}
Return

^3:: 
Send {Shift down}{Ctrl down}3 
Send {Shift up}{Ctrl up}
Return
}

;//////////////////////////////////////////////////////////////////////////////


;//////////////////////////////////////////////////////////////////////////////
;////////////////////////// OUTLOOK SHORTCUTS /////////////////////////////////
;
;create a new email message with ctrl+alt+m
^!m::
Run, C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE /c ipm.note
return
;
;create a new Outlook task with ctrl+alt+t
^!t::
Run, C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE /c ipm.task
WinWaitActive, ahk_class rctrl_renwnd32
send, This is a New Task
return
;
;create a new appointment with ctrl+alt+a
^!a::
Run, C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE /c ipm.appointment
return
;
;;;;Automated emails GUI;;;
;create template email message with ctrl+alt+e
^!e::
	Gui, +AlwaysOnTop
	Gui, Add, Text,, Which automated email do you wish to send?
	Gui, Add, Button, X20 Y30, &Ergo Team Meeting Topics
	Gui, Add, Button, X20 Y60, &Status Update Tracker
	Gui, Add, Button, X20 Y90, &Large Fix Log
	Gui, Show,, Automated Emails
	return

	ButtonErgoTeamMeetingTopics: ;use this format: "ButtonButtonPromptWithoutSpaces"
	{
		Gui, Destroy ;this kills the GUI prompt. Optional if you want to always keep the window open.
		Run, C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE /c ipm.note
		WinWaitActive, Untitled - Message (HTML),,10
		if ErrorLevel
		{
			MsgBox, WinWait timed out.
			return
		}
		else 
		Send Vikas Chowdhry; Kyle Reger; Steve Cubinski ;the to: line of the email address
		Sleep, 500
		Send {Tab}{Tab}{Tab}
		Sleep, 500  
		Send Ergo Team Meeting Topics ;the subject line of the email address
		Sleep, 500
		Send {Tab} 
		Send Do you have any topics for this week's team meeting?
		Send {enter}
		Send {enter}
		Send {Raw}onenote:///F:\Reporting`\Predictive`%`2`0Analytics`\Predictive`%`2`0Modeling`\Ergo`\Team`%`2`0Meeting`%`2`0Notes.one#section-id={91EC3843-FBF8-423C-89B1-0835DCEA1167}&end
		Send {enter}
	return 
	}
	ButtonStatusUpdateTracker: ;use this format: "ButtonButtonPromptWithoutSpaces"
	{
		Gui, Destroy ;this kills the GUI prompt. Optional if you want to always keep the window open.
		{
		ErrorLevel=0
		InputBox,input,Input Box,If you need to send the email to different folks each time, you can enter their names here.:
		if ErrorLevel
		Return
		else
		names=%input%
		}
		Run, C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE /c ipm.note
		WinWaitActive, Untitled - Message (HTML),,10
		if ErrorLevel
		{
			MsgBox, WinWait timed out.
			return
		}
		else
 
		Send Crescentia Stegner-Freitag ;the to: line of the email address
		Sleep, 500
		Send {Tab}{Tab}{tab} 
		Sleep, 500  
		Send weekly status update ;the subject line of the email address
		Sleep, 500
		Send {Tab} 
		Send This is the test of the "status update tracker" email address.
		Send Here is where I can use the names of the people entered in the input box earlier: ^b%names%^b
		Send {enter}
		Send ^uPRJ Links: ^u 
		Send {enter}
		Send Here is how I can embed an EMC2 link into the email (it{'}s not pretty, but it works. KPI: ^kemc2://TRACK/PRJ/134134?action=EDIT{enter}
	return 
	}
	ButtonLargeFixLog: 
	{
		Gui, Destroy
		Run, C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE /c ipm.note
		WinWaitActive, Untitled - Message (HTML),,10
		if ErrorLevel
		{
			MsgBox, WinWait timed out.
			return
		}
		else
		Send {tab}{tab}{tab} ;skip the to:, cc:, and bcc: lines because their content varies a lot
		Send FYI: large fix logs of the week ;subject line of the email
		Send {tab}
		Send Here is the text of my "large fix log" email.
	return 
	}
	return
;//////////////////////////////////////////////////////////////////////////////


;;;;;;;;;;;;;Volume Control;;;;;;;;;;;;;;;
;Volume Up	 Windows+PgUp
#PgUp::Send {Volume_Up 3}
;Volume Down	 Windows+PgDn
#PgDn::Send {Volume_Down 3}
;Volume Mute	 Windows+End
#End::Send {Volume_Mute}
;Volume Max	 Windows+Home
;#Home::Send {Volume_Up 100}

;;;;;;;;;Text-only paste with Win+V;;;;;;;;;;;
#v::                           ; Text-only paste from ClipBoard
 Clip0 = %ClipBoardAll%        ; Save original clipboard
 ClipBoard = %ClipBoard%       ; Convert to text
 ; Replace Win1252 characters
 char := Chr(145) ;0x2018 Left Single Quotation Mark
 StringReplace, ClipBoard, ClipBoard, %char%, `', All
 char := Chr(146) ;0x2019 Right Single Quotation Mark
 StringReplace, ClipBoard, ClipBoard, %char%, `', All
 char := Chr(147) ;0x201c Left Double Quotation Mark
 StringReplace, ClipBoard, ClipBoard, %char%, `", All
 char := Chr(148) ;0x201d Right Double Quotation Mark
 StringReplace, ClipBoard, ClipBoard, %char%, `", All
 char := Chr(150) ;0x2013 En Dash
 StringReplace, ClipBoard, ClipBoard, %char%, `-, All
 char := Chr(151) ;0x2014 Em Dash
 StringReplace, ClipBoard, ClipBoard, %char%, `-, All
 Send ^v
 Sleep 50                      ; Don't change clipboard while it is pasted! (Sleep > 0)
 ClipBoard = %Clip0%           ; Restore original ClipBoard
 VarSetCapacity(Clip0, 0)      ; Free memory
 Return

;;;;;;;;;CTRL+right-click to open IE links from Chrome;;;;;;;;;;;
^RButton::
ifwinactive, ahk_class Chrome_WidgetWin_1
{
ClipSaved := ClipboardAll ;don't lose what's on the clipboard
sendinput {RButton}
sleep, 500
send, e
sleep 200 
URL:=clipboard
Run, "C:\Program Files (x86)\Internet Explorer\iexplore.exe" %URL%
WinActivate, ahk_class IEFrame
sleep, 700
sendinput {enter}
IfWinNotActive, ahk_class IEFrame
{
WinActivate ahk_class IEFrame
sendinput {^w}
}
Clipboard := ClipSaved
ClipSaved = ;empties the variable to save space
}
return


;/////////////// OLD RETIRED SCRIPTS /////////////////////////////
;Map CTRL+SHIFT+Z to CTRL+SHIFT+V for ArsClip
;^+V::Send {CTRL DOWN}{Shift+Z}{CTRL UP}

;Re-map Copy and Paste for Reflections Sessions
;CTRL+SHIFT+C -> Copy (CTRL+Insert)
;^+c::
;Send {CTRL DOWN}{Ins}{CTRL UP}
;CTRL+SHIFT+V -> Paste (SHIFT+Insert)
;^+v::Send {SHIFTDOWN}{Ins}{SHIFTUP}

