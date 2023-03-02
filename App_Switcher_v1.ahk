#NoEnv    ;Recommended for performance and compatibility with future AutoHotKey releases.
; #Warn   ; Enable warnings to assist with detecting common errors.
SendMode Input     ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%    ; Ensures a consistent starting directory.
#SingleInstance, Force ;Only launch 1 instance of this script.


;These codes will launch the app specified in the script. Use AHK Windows Spy to check what is the ahk_exe code of the app.
;If the app already exist, it will activate the app and cycle on all its opened instances.


;==========================CTRL COMMANDS===============================================================

#IfWinActive

^Numpad1::						;Ctrl + Numpad1 = Run Calc
	IfWinExist, ahk_class ApplicationFrameWindow	
	{	
		IfWinActive, ahk_class ApplicationFrameWindow
			{
			GroupAdd, HaroldCalc, ahk_class ApplicationFrameWindow
			GroupActivate, HaroldCalc, r
			}
		Else 
			WinActivate, ahk_class ApplicationFrameWindow
		Return
	}	
	Else
		Run, calc.exe	
	Return


^Numpad2::						;Ctrl + Numpad2 = Notepad
	IfWinExist, ahk_class Notepad	
	{	
		IfWinActive, ahk_class Notepad
			{
			GroupAdd, HaroldNotepad, ahk_class Notepad
			GroupActivate, HaroldNotepad, r
			}
		Else 
			WinActivate, ahk_class Notepad
		Return
	}	
	Else
		Run, notepad.exe	
	Return



^Numpad3::						;Ctrl + Numpad3 = SAP B1
	IfWinExist, ahk_class TMFrameClass	
	{	
		IfWinActive, ahk_class TMFrameClass
			{
			GroupAdd, HaroldSAP, ahk_class TMFrameClass
			GroupActivate, HaroldSAP, r
			}
		Else 
			WinActivate, ahk_class TMFrameClass
		Return
	}	
	Else
		Run, SAP Business One.exe	
	Return


^Numpad4::						;Ctrl + Numpad4 = Excel
	IfWinExist, ahk_class XLMAIN	
	{	
		IfWinActive, ahk_class XLMAIN
			{
			GroupAdd, HaroldExcel, ahk_class XLMAIN
			GroupActivate, HaroldExcel, r
			}
		Else 
			WinActivate, ahk_class XLMAIN
		Return
	}	
	Else
		Run, EXCEL.EXE	
	Return



^Numpad5::						;Ctrl + Numpad5 = PowerPoint
	IfWinExist, ahk_class PPTFrameClass	
	{	
		IfWinActive, ahk_class PPTFrameClass
			{
			GroupAdd, HaroldPPT, ahk_class PPTFrameClass
			GroupActivate, HaroldPPT, r
			}
		Else 
			WinActivate, ahk_class PPTFrameClass
		Return
	}	
	Else
		Run, POWERPNT.EXE	
	Return



^Numpad6::						;Ctrl + Numpad6 = Word
	IfWinExist, ahk_class OpusApp	
	{	
		IfWinActive, ahk_class OpusApp
			{
			GroupAdd, HaroldWord, ahk_class OpusApp
			GroupActivate, HaroldWord, r
			}
		Else 
			WinActivate, ahk_class OpusApp
		Return
	}	
	Else
		Run, WINWORD.EXE	
	Return



^Numpad7::						;Ctrl + Numpad7 = Outlook
	IfWinExist, ahk_class rctrl_renwnd32	
	{	
		IfWinActive, ahk_class rctrl_renwnd32
			{
			GroupAdd, HaroldOutlook, ahk_class rctrl_renwnd32
			GroupActivate, HaroldOutlook, r
			}
		Else 
			WinActivate, ahk_class rctrl_renwnd32
		Return
	}	
	Else
		Run, OUTLOOK.EXE	
	Return




^Numpad8::						;Ctrl + Numpad8 = Edge
	IfWinExist, ahk_class Chrome_WidgetWin_1	
	{	
		IfWinActive, ahk_class Chrome_WidgetWin_1
			{
			GroupAdd, HaroldEdge, ahk_class Chrome_WidgetWin_1
			GroupActivate, HaroldEgde, r
			}
		Else 
			WinActivate, ahk_class Chrome_WidgetWin_1
		Return
	}	
	Else
		Run, brave.exe	
	Return


^Numpad9::						;Ctrl + Numpad9 = OneNote
	IfWinExist, ahk_class Framework::CFrame	
	{	
		IfWinActive, ahk_class Framework::CFrame
			{
			GroupAdd, HaroldOneNote, ahk_class Framework::CFrame
			GroupActivate, HaroldOneNote, r
			}
		Else 
			WinActivate, ahk_class Framework::CFrame
		Return
	}	
	Else
		Run, ONENOTE.EXE	
	
Return



^NumpadDot::						;Ctrl + NumpadDot = Messenger
	IfWinExist, ahk_class Messenger0
	{	
		IfWinActive, ahk_class Messenger0
			{
			GroupAdd, HaroldMesenger, ahk_class Messenger0
			GroupActivate, HaroldMessenger, r
			}
		Else 
			WinActivate, ahk_class Messenger0
		Return
	}	
	Else
		Send, #1
	Return



^Numpad0::						;Ctrl + Numpad0 = File Explorer
	IfWinExist, ahk_class CabinetWClass
	{	
		IfWinActive, ahk_class CabinetWClass
			{
			GroupAdd, HaroldFile, ahk_class CabinetWClass
			GroupActivate, HaroldFile, r
			}
		Else 
			WinActivate, ahk_class CabinetWClass
		Return
	}	
	Else
		Run, explorer.exe	
	
Return



;==========================SHIFT COMMANDS===============================================================

#IfWinActive

+NumpadAdd::				;Excell Zoom In
	Send ^{WheelUp}
Return


+NumpadSub::				;Excell Zoom In
	Send ^{WheelDown}
Return


;==========================WIN COMMANDS===============================================================


#IfWinActive


#Numpad7::Run, http://20.1.1.247:81/			;Run HRC



#Numpad8::						;English to Japanse
	IfWinExist, Google Translate
		{
		Send, ^c
		WinActivate, Google Translate
		}
	Else
		{
		Send, ^c
		Run, https://translate.google.com/?sl=ja&tl=en&op=translate	
		Sleep, 2000
		Send, ^v
		}
	Return


#Numpad9::Run, https://rcbconline-corporate.com/fo/login


#Numpad4::
	IfWinExist, Accounting PIC (Public)
		WinActivate,Accounting PIC (Public)
	Else
		Run, \\20.1.1.4\Accounting PIC (Public)
	Return


#Numpad5::
	IfWinExist, Files
		WinActivate, Files
	Else
		Run, C:\Users\htsicat\Desktop\Desktop Files\Files
	Return


#Numpad6::					;Autologin RCBC
	;Declare Variables
	RCBCsite:= "https://rcbconline-corporate.com/fo/login"
	CorpCode:= "********"
	UserID:= "*********"
	PW:= "***********"

	Run, %RCBCsite%
	Msgbox, 4097,ROC is Loading,Press OK if done loading

	IfMsgBox OK
		{
		send, {Tab}
		send, {Tab}
		Send, %CorpCode%
		send {Tab}
		Send, %UserID%
		send, {Tab}
		send, %PW% 
		Sleep, 50
		Send, {Enter}
		}
	Else
		{
		WinActivate, Business Banking		;Change Active Window name if necessary
		send, ^w				;Close RCBC if canceled
		}
	Return
	
	
#Numpad1::				;Google Search
	{
	Run, https://www.google.com/search?q=%clipboard%
	Return
	}


#NumLock::Send {Volume_Mute}

#NumpadDiv::Send {Volume_Down}

#NumpadMult::Send {Volume_Up}
	






