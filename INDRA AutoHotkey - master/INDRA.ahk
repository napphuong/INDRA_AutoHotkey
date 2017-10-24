;===================================
;GUI
;===================================
Gui,+AlwaysOnTop
Gui, Add, CheckBox, x10 y5 w100 gCheck vMyCheckBox, Wrap up! (F1)
MyCheckBox:= 0
Gui, Add, CheckBox, x115 y5 w110 gUpdateMode vUpdateMode, Update Mode
UpdateMode:= 0
Gui, Add, Button, x10 y25 w100 Default gStartDownload, Start Download
Gui, Add, Button, x115 y25 w110 Default gContinueDownload, Continue Download
Gui, Add, Edit, x10 y50 w100 ReadOnly, Download Folder
Gui, Add, Edit, x115 y50 w110 vDownloadFolder, E:\INDRA\
Gui, Add, Edit, x10 y75 w100 ReadOnly, Download Item
Gui, Add, Edit, x115 y75 w110 vDownloadItem,
Gui, Add, Edit, x10 y100 w100 ReadOnly, Delay Timer
Gui, Add, Edit, x115 y100 w110 vDelayTime,
xLocation:=A_ScreenWidth-235
Gui, Show, x%xLocation% y0 w235 h130, INDRA
return

GuiEscape: 
GuiClose: 
ExitApp
return

UpdateMode:
UpdateMode:= !UpdateMode
return

Check:
MyCheckBox:= !MyCheckBox
return

F1::
Gui, Submit, NoHide
MyCheckBox:= !MyCheckBox
if MyCheckBox
	Guicontrol,,MyCheckBox, 1
else
	Guicontrol,,MyCheckBox, 0
return

Pause::Pause





;===================================
;CORE FEATURES
;===================================
StartDownload:
	Gosub, CheckInitialCondition

NextItem:
	GuiControlGet, DownloadFolder
    
	; check every 15 minutes if idle time is too long reset all.
	SetTimer, trapper, 30000

	Loop 
	{
		WinActivate, id.xlsx - Excel
		WinWaitActive, id.xlsx - Excel, , 5
		if (!ErrorLevel) 
		{
			break
		}
	}
	Send, ^{Left}
	Loop {
		Clipboard =
		Send ^c
		ClipWait 1
		If (StrLen(Clipboard) > 3) 
		{
			MyFolder0 = 
			MyFolder0:= Clipboard
			StringTrimRight, MyFolder0, MyFolder0, 2
			MyFolder0= %MyFolder0%
		}
		Send, {Right}

		Clipboard =
		MyFolder1 = 
		Send ^c
		ClipWait 1
		MyFolder1:= Clipboard
		StringTrimRight, MyFolder1, MyFolder1, 2
		MyFolder1= %MyFolder1%
		; MyFolder1, vd: "000 - Documents common to whole plant"

		If (StrLen(MyFolder1) < 2) 
		{
			break
		}
		Loop {
			Loop {
				WinActivate, id.xlsx - Excel
				WinWaitActive, id.xlsx - Excel, , 5
				if (!ErrorLevel) 
				{
					Send, {Right}
					break
				}
			}
			
			Clipboard =
			CurrentItem =
			; CurrentItem, vd: D-000-1225xxxxxxx, "VP215A-501xxxxxx
			Send ^c
			ClipWait 1
			CurrentItem := Clipboard
			If (StrLen(CurrentItem) < 5) 
			{
				break
			}
			else
			{
				StringTrimRight, CurrentItem, CurrentItem, 2
				UpdateItem = %CurrentItem%
				; UpdatItem, vd: D-000-1314-1001V ,VV2158-901-A-001
			}
			
			MyFolder3 := SubStr(CurrentItem, 1 , 10)
			; MyFolder3, vd: "D-000-1225", "VP215A-501"
			MyFolder2 := SubStr(MyFolder3, 1 , InStr(MyFolder3, "-", , 4) - 1)
			; MyFolder2, vd: "D-000", "VP215A"

			if (UpdateMode = 1) 
			{
				if (InStr(MyFolder1, "-") = 4 and UpdateFolder1 != MyFolder1)
				{
					UpdateFolder1 := MyFolder1
					DownloadItem= %DownloadFolder%%MyFolder0%\%MyFolder1%\
					FileAppend, %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadItem%`n, %DownloadFolder%Updated.txt
					GuiControl,,DownloadItem, %DownloadItem%
					IfNotExist, %DownloadItem%
					{
						FileCreateDir, %DownloadItem%
					}
					FileDelete,  %DownloadItem%*.csv
					Clipboard:= SubStr(MyFolder1, 1 , 2)
					Gosub, DownloadTextFile
				}				

				if (UpdateFolder2 != MyFolder2)
				{
					UpdateFolder2 := MyFolder2
					; MyFolder2, vd: "D-000", "VP215A"
					DownloadItem= %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\
					FileAppend, %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadItem%`n, %DownloadFolder%Updated.txt
					GuiControl,,DownloadItem, %DownloadItem%
					IfNotExist, %DownloadItem%
					{
						FileCreateDir, %DownloadItem%
					}
					FileDelete,  %DownloadItem%*.csv
					Clipboard= %MyFolder2%
					Gosub, DownloadTextFile
				}
				if (UpdateFolder3 != MyFolder3)
				{
					UpdateFolder3 := MyFolder3
					; MyFolder3, vd: "D-000-1225", "VP215A-501"	
					DownloadItem= %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\%MyFolder3%\
					FileAppend, %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadItem%`n, %DownloadFolder%Updated.txt
					GuiControl,,DownloadItem, %DownloadItem%
					IfNotExist, %DownloadItem%
					{
						FileCreateDir, %DownloadItem%
					}
					FileDelete,  %DownloadItem%*.csv
					Clipboard= %MyFolder3%
					Gosub, DownloadTextFile
				}
				DownloadItem= %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\%MyFolder3%\
				FileAppend, %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadItem%%UpdateItem%`n, %DownloadFolder%Updated.txt
				GuiControl,,DownloadItem,%DownloadItem%
				FileDelete,  %DownloadItem%%UpdateItem%*
				Clipboard= %UpdateItem%
				Gosub, SearchAndDownloadAll2
			} 
			else ;normal download mode (if exist skip, if not exist, download all Folder3)
			{

				DownloadItem= %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\
				IfNotExist, %DownloadItem% 
				{
					FileCreateDir, %DownloadItem%
					Clipboard= %MyFolder2%
					FileAppend, %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadItem%`n, %DownloadFolder%Downloaded.txt					
					GuiControl,,DownloadItem, %DownloadItem%
					Gosub, DownloadTextFile
				}
				DownloadItem= %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\%MyFolder3%\
				IfNotExist, %DownloadItem% 
				{
					FileCreateDir, %DownloadItem%
					Clipboard= %MyFolder3%
					FileAppend, %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadItem%`n, %DownloadFolder%Downloaded.txt
					GuiControl,,DownloadItem, %DownloadItem%
					Gosub, DownloadTextFile
					Gosub, SearchAndDownloadAll2
				}
			}

			If (MyCheckBox) 
			{
				SetTimer, trapper, Off
				return
			}
			Loop 
			{
				WinActivate, id.xlsx - Excel
				WinWaitActive, id.xlsx - Excel, , 5
				if (!ErrorLevel) 
				{
					break
				}
			}
		}
		Send, {Left}
		Send, ^{Left}
		Send, {Down}
		Send, {Left}
	}

	SetTimer, trapper, Off
return

ContinueDownload:
	; seach the item, navigate to the page you want to continue.
	Gosub, CheckInitialCondition
	Loop {
		GuiControlGet, DownloadItem
		If (DownloadItem = "") {
			Msgbox, Please give path to Current Download Item
			return
		}Else{
			break
		}
	}
	Gosub, DownloadWithOutSearching
	If (!MyCheckBox) {
		Gosub, NextItem
	}
return

;===================================
;SUB ROUTINES
;===================================





;Load INDRA
;==========
LoadINDRA:
	Run, iexplore.exe https://wpms.jgc.com

	Loop {
		WinActivate, Login for PMS (Specifying User ID & Password) - Internet Explorer
		WinWaitActive, Login for PMS (Specifying User ID & Password) - Internet Explorer, , 10
		If (!ErrorLevel) {
			send KX493
			send {tab}
			send Phuongmai124
			send, {enter}
			send, {enter}
			break
		}
	}
	sleep, 500
	Loop {
		WinActivate, Specifying Project ID/Group ID & Printer ID - Internet Explorer
		WinWaitActive, Specifying Project ID/Group ID & Printer ID - Internet Explorer, , 10
		If (!ErrorLevel) {
			loop, 5 
			{
				send, {enter}
			}
		break
		}
	}
	sleep, 500
	Loop {
		WinActivate, PMS Integrated Menu - Internet Explorer
		WinWaitActive, PMS Integrated Menu - Internet Explorer, , 5
		If (!ErrorLevel) {
			;Wait for loading
			Loop {
				PixelSearch, Px, Py, 29,209, 38,214, 0xC0C0C0, 3, RGB
				if (!ErrorLevel) {
					click 17, 209
					break
				}
				sleep, 50
			}
			sleep, 1000
			Loop {
				PixelSearch, Px, Py, 43,226,55,238, 0x008080, 3, RGB
				if (!ErrorLevel) 
				{
					click 130, 230
					break
				}
				sleep, 50
			}
			break
		}
		sleep, 50
	}
	sleep, 1000
	Loop {
		WinActivate, PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		If !ErrorLevel 
		{
			;Wait for loading
			Loop {
				PixelSearch, Px, Py, 39,165,49,169, 0xC0C0C0, 3, RGB
				if (!ErrorLevel) 
				{
					click 26, 167
					break
				}
				sleep, 50
			}
			sleep, 1000
			Loop {
				PixelSearch, Px, Py, 54,208,62,218, 0xC0C0C0, 3, RGB
				if (!ErrorLevel) 
				{
					click 156, 212 
					break
				}
				sleep, 50
			}
			break
		}
		sleep, 50
	}

	; close all sub windows if any
	Gosub, CloseSubWindows

return





;Check Inital Condition
;======================
CheckInitialCondition:
	; INDRA is not loaded.
	Loop {
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 0
		If (!ErrorLevel) {
			break
		} else {
			Gosub, CloseAllINDRAWindows    
			sleep, 500
			Gosub, LoadINDRA
		}
	}

	;id.xlsx is not opened.
	Loop 
	{
		WinActivate, id.xlsx - Excel
		WinWaitActive, id.xlsx - Excel, , 1
		If (!ErrorLevel) 
		{
			break
		} 
		else 
		{
			MsgBox, 1, INDRA, Please open id.xlsx and select the first cell.
			IfMsgBox Cancel
				return
		}
	}

	; close all sub windows if any
	Gosub, CloseSubWindows

return






; Input data to INDRA
;====================

InputData2:
	;switch to INDRA
	Loop 
	{
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if !ErrorLevel 
		{
			PixelSearch, Px, Py, 943, 64, 995, 89, 0xFF0000, 3, RGB
			if (!ErrorLevel) 
			{
				;load INDRA	
				Click 148, 211
				break
			}
		}
	}

	;Wait for IE window become blank.
	Loop 
	{
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if !ErrorLevel 
		{
			PixelSearch, Px, Py, 398, 125, 419, 137, 0xFFFFFF, 3, RGB
			if !ErrorLevel
				break
		}
	}

	;Wait for IE window loaded.
	Loop 
	{
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if !ErrorLevel 
		{
			PixelSearch, Px, Py, 361, 495, 395, 507, 0x666666, 3, RGB
			if !ErrorLevel
				break
		}
	}

	;input search value and option
	Loop {
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if !ErrorLevel
			break
	}
	Loop, 3 
	{
		Send, {Tab}
	}
	Send, ^v
	Send, `%

	Loop, 17 
	{
		Send, {Tab}
	}	
	Send, {Down}
	Send, {Down}

	Loop, 27 
	{
		Send, {Shift down}{Tab}{Shift up}
	}
	Send, {Right}
	Send, {Right}

	Loop, 2	
	{
		Send, {Shift down}{Tab}{Shift up}
	}
	Send, {Left}
	Send, {Shift down}{Tab}{Shift up}
	Send, {Right}
	Send, {Shift down}{Tab}{Shift up}
return





; Download Text File
;===================
DownloadTextFile:
	Gosub, InputData2

	;download text file
	Loop 
	{
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if (!ErrorLevel) 
		{
			Click 358, 199
			break
		}
		sleep, 50
	}

	Gosub, Download2
	Gosub, FinishDownload
	sleep 100
return





;Search and download all items (include text file)
;=================================================
SearchAndDownloadAll2:
	Gosub, InputData2

	;search file
	Loop 
	{
		Gosub, CloseSubWindows
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if !ErrorLevel 
		{
			Click 271, 201 ; search button
			; waiting for IE to be blank
			PixelSearch, Px, Py, 398, 125, 419, 137, 0xFFFFFF, 3, RGB
			if !ErrorLevel
				break
		}
		; sub windows error
		WinActivate, Message from webpage
		WinWaitActive, Message from webpage, , 2
		If (ErrorLevel = 0) 
		{
			Loop 
			{
				click 416, 138
				sleep, 50
				WinActivate, Message from webpage
				WinWaitActive, Message from webpage, , 2
				if (ErrorLevel) 
				{	
					FileAppend, Message from webpage activated %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadItem%`n, %DownloadFolder%Download.txt
					break
				}
			}
		}
		sleep, 50
	}

DownloadWithOutSearching:
	Loop 
	{
		Gosub, DownloadOnePage
		; exit loop and minimize
		if (toValue = totalRecordsNumber) 
		{
			WinMinimize , PMS INDRA - Internet Explorer
			return
		}

		Loop {
			WinActivate PMS INDRA - Internet Explorer
			WinWaitActive, PMS INDRA - Internet Explorer, , 5
			if (!ErrorLevel) 
			{
				Click 301, 219 ; next button
				break
			}
			sleep, 50
		}

		;Wait for IE window become blank.
		Loop 
		{
			WinActivate PMS INDRA - Internet Explorer
			WinWaitActive, PMS INDRA - Internet Explorer, , 5
			if !ErrorLevel 
			{
				PixelSearch, Px, Py, 398, 125, 419, 137, 0xFFFFFF, 3, RGB
				if !ErrorLevel
					break
			}
			sleep, 50
		}
		sleep, 50
	}
return





;Download 1 Page
;===============
DownloadOnePage:

	Loop 
	{
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if !ErrorLevel 
		{
			; check if IE is loaded.
			PixelSearch, Px, Py, 811, 280, 870, 299, 0x666666 , 5, RGB
			if !ErrorLevel
				break
		}
		sleep, 50
	}

	;get number of record
	Loop
	{
		Clipboard =
		MyString =
		Loop 
		{
			WinActivate PMS INDRA - Internet Explorer
			WinWaitActive, PMS INDRA - Internet Explorer, , 5
			if !ErrorLevel
				break
		}
		Send ^a
		Send ^c
		ClipWait 1
		MyString := Clipboard

		pos0:= InStr(MyString, "e-File List")
		pos1:= InStr(MyString,"to")
		pos2:= InStr(MyString,"of")
		pos3:= InStr(MyString,"Records")

		fromStart:= pos0 + 29
		fromLength:= pos1 - pos0 - 30
		fromValue:= SubStr(MyString, fromStart, fromLength)

		toStart:= pos1 + 4
		toLength:= pos2 - pos1 - 5
		toValue:= SubStr(MyString, toStart, toLength)

		pageRecordValue:= (toValue - fromValue + 1)/2

		totalRecordsStart:= pos2 + 2
		totalRecordsLength:= pos3 - pos2 -3
		totalRecordsNumber:= SubStr(MyString, totalRecordsStart, totalRecordsLength)
		if pageRecordValue > 0
			break
	}

	;download all if pageRecordValue <31
	if (pageRecordValue <31) 
	{
		Gosub, DownThemAll
	}
	else 
	{
		;download each 25 records if pageRecordValue >=31
		skipItems:= 0
		unselectItems:= 0
		selectItems:= 25
		Gosub, DownloadItems
		unselectItems:= 25
		Loop {
			pageRecordValue:= pageRecordValue - 25
			if (pageRecordValue <= 0) 
			{
				break
			} 
			else if (pageRecordValue < 25) 
			{
				selectItems:= pageRecordValue
			}
			Gosub, DownloadItems
			skipItems:= skipItems + 25
		}
	}
return





;Select All and download
;=======================
DownThemAll:
	Loop 
	{
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if !ErrorLevel
			break
	}
	Click 257, 130
	Click 314, 239
	Click 257, 130
	Click 317, 175
	Gosub, Download1
	sleep 100
	Gosub, FinishDownload
return





;Select and download some items in 1 Page
;========================================
DownloadItems:
	Loop 
	{
		WinActivate PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 5
		if !ErrorLevel
			break
	}
	Click 753, 418
	Loop, 10 
	{
		Send {PgUp}
	}
	Sleep 500
	Click 245, 317
	Send, {space}
	Sleep 200
	Loop, %skipItems% 
	{
		Send, {Tab}
	}
	Sleep 100
	Loop, %unselectItems% 
	{
		Send, {Space}
		Send, {Tab}
	}
	Sleep 100
	Loop, %selectItems%	
	{
		Send, {Space}
		Send, {Tab}
	}
	Send, {Space}
	Send, {Tab}
	Sleep 100
	Click 257, 130
	Click 317, 175
	Gosub, Download1
	sleep 100
	Gosub, FinishDownload
return





;Download after selecting and click download
;===========================================
Download1:
	Loop 
	{
		WinWaitActive, Download E-File - Internet Explorer, , 0
		If (ErrorLevel = 0) 
		{
			break
		}
		sleep, 50
	}

	; wait for loading
	Loop 
	{
		PixelSearch, Px, Py, 108, 301, 115, 309, 0x212121, 5, RGB
		if !ErrorLevel
			break
		sleep, 50
	}

	;download as PDF
	;Click 156, 203

	;wait untill the "Execute" button is shown
	MouseMove 175, 412
	Loop 
	{
		PixelSearch, Px, Py, 140, 394, 211, 421, 0xA6F4FF, 5, RGB
		if !ErrorLevel
			break
		sleep, 50
	}
	Click 177, 408
 
Download2:
	Loop 
	{
		WinActivate, File Transfer for PMS - Internet Explorer
		WinWaitActive, File Transfer for PMS - Internet Explorer, , 0
		If (ErrorLevel = 0) 
		{
			break
		}

		; more than 3000 items
		WinActivate, Message from webpage
		WinWaitActive, Message from webpage, , 0
		If (ErrorLevel = 0) 
		{
			Loop 
			{
				click 319, 139
				sleep, 50
				WinActivate, Message from webpage
				WinWaitActive, Message from webpage, , 0
				if ErrorLevel
					break
			}
			break
		}
	}

	; wait for loading
	Loop 
	{
		WinActivate, File Transfer for PMS - Internet Explorer
		PixelSearch, Px, Py, 26, 80, 40, 96, 0x212121, 5, RGB
		if !ErrorLevel
			break
		sleep, 50
	}
	Click 408, 67

	; wait for "Download" button is shown
	Loop 
	{
		WinActivate, File Transfer for PMS - Internet Explorer
		PixelSearch, Px, Py, 208, 156, 269, 170, 0x212121, 5, RGB
		if !ErrorLevel
				break
		sleep, 50
	}
	Click 245, 163
	Clipboard:= DownloadItem

	; wait for "Save" button is shown
	Loop
	{
		WinActivate, File Transfer for PMS - Internet Explorer
		PixelSearch, Px, Py, 317, 608, 325, 614, 0x000000, 5, RGB
		if !ErrorLevel
			break
		sleep, 50
	}

	; loop until SaveAs window appear
	ClickSaveAs:
	Loop 
	{
		WinActivate, File Transfer for PMS - Internet Explorer
		WinWaitActive, File Transfer for PMS - Internet Explorer, , 1
		Mousemove 320, 610
		Click 320, 610
		Send, {down}
		Send, {enter}

		WinWait, Save As, , 1
		If (ErrorLevel = 0)
		{
			break
		}
	}

	; do Save As actions.
	Loop 
	{
		WinActivate, Save As
		WinWaitActive, Save As, , 1
		If (ErrorLevel = 0) 
		{
			Click 398, 48
			Send, ^v
			send {enter}
			Click 515, 447
			WinWait, Confirm Save As, , 1
			If !ErrorLevel 
			{
				WinActivate, Confirm Save As
				WinWaitActive, Confirm Save As, , 1
				send, !n
				WinClose , Save As
				; Check if SaveAs window disappeared.
				WinWait, Confirm Save As, , 0
				If (ErrorLevel = 1) 
				{
					WinWait, Save As, , 0
					If (ErrorLevel = 1) 
					{
						FileAppend, Confirm Save As activated %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\%MyFolder3%`n, %DownloadFolder%Downloaded.txt
						Gosub, ClickSaveAs
						return
					}
				}
			}
			break	
		}
		sleep, 50	
	}

	return





;Close 2 sub windows of INDRA
;============================
CloseSubWindows:
	Loop 
	{
		IfWinExist, Download E-File - Internet Explorer
		{
			WinClose , Download E-File - Internet Explorer
			WinWait, Download E-File - Internet Explorer, , 1
			If (ErrorLevel = 1)
				break
			sleep, 50
		}
		else
		{
			break
		}
	}

	Loop 
	{
		IfWinExist, File Transfer for PMS - Internet Explorer
		{
			WinClose , File Transfer for PMS - Internet Explorer
			WinWait, File Transfer for PMS - Internet Explorer, , 1
			If (ErrorLevel = 1)
				break
			sleep, 50
		}
		else
		{
			break
		}
	}

return





;Close all INDRA Windows
;=======================
CloseAllINDRAWindows:
	Gosub, CloseSubWindows

	Loop 
	{
		IfWinExist, Appserver Error - Internet Explorer
		{
			WinClose , Appserver Error - Internet Explorer
			WinWait, Download E-File - Internet Explorer, , 1
			If (ErrorLevel = 1)
				break
			sleep, 50
		}
		else
		{
			break
		}
	} 

	Loop 
	{
		IfWinExist, PMS INDRA - Internet Explorer
		{
			WinClose , PMS INDRA - Internet Explorer
			WinWait, Download E-File - Internet Explorer, , 1
			If (ErrorLevel = 1)
				break
			sleep, 50
		}
		else
		{
			break
		}
	}  

	Loop 
	{
		IfWinExist, PMS Integrated Menu - Internet Explorer
		{
			WinClose , PMS Integrated Menu - Internet Explorer
			WinWait, Download E-File - Internet Explorer, , 1
			If (ErrorLevel = 1)
				break
			sleep, 50
		}
		else
		{
			break
		}
	} 

	Loop 
	{
		IfWinExist, JGC PMS - Internet Explorer
		{
			WinClose , JGC PMS - Internet Explorer
			WinWait, Download E-File - Internet Explorer, , 1
			If (ErrorLevel = 1)
				break
			sleep, 50
		}
		else
		{
			break
		}
	}

	Loop 
	{
		IfWinExist, File Cleaning for PMS - Internet Explorer
		{
			WinClose , File Cleaning for PMS - Internet Explorer
			WinWait, Download E-File - Internet Explorer, , 1
			If (ErrorLevel = 1)
				break
			sleep, 50
		}
		else
		{
			break
		}
	}

	Loop 
	{
		IfWinExist, Calling JGCFTRAN Cleaner - Internet Explorer
		{
			WinClose , Calling JGCFTRAN Cleaner - Internet Explorer
			WinWait, Download E-File - Internet Explorer, , 1
			If (ErrorLevel = 1)
				break
			sleep, 50
		}
		else
		{
			break
		}
	}

return






;Close sub-windows and set delay timer after download 1 package of items
;=======================================================================
FinishDownload:
	Gosub, CloseSubWindows
	Gosub, CloseSubWindows
	Gosub, CloseSubWindows

	;set delay timer

	IfWinExist, ahk_class #32770
	{
		Loop 
		{
			WinActivate, ahk_class #32770
			WinWaitActive, ahk_class #32770, , 1
			if !ErrorLevel
				break
			sleep, 50
		}
	}
	else
	{
		Loop 
		{
			WinActivate, PMS INDRA - Internet Explorer
			WinWaitActive, PMS INDRA - Internet Explorer, , 1
			if !ErrorLevel
			{
				send, ^j
				break
			}
			sleep, 50
		}
	}

	WinGetTitle, DownloadWindowTitle, A
	If InStr(DownloadWindowTitle, "ownload")
	{ 
		WinMove, A,, 1306, 691
	}
	DelayTime:=0
	If InStr(DownloadWindowTitle, "downloads in progress") 
	{
		NumberOfDownload:= SubStr(DownloadWindowTitle, 1, InStr(DownloadWindowTitle, "downloads in progress") - 2)
		If (NumberOfDownload > 6) 
		{
			DelayTime:=120/(1+EXP(-(NumberOfDownload-6)))
			Loop 
			{
				GuiControl,,DelayTime,%DelayTime%
				Sleep, 1000
				DelayTime:= DelayTime - 1
				If (DelayTime < 0) 
				{
					break
				}
			}
		}	
	}

	;View Downloads - Internet Explorer
	;15.5 MB of 20170716004252_0000528911.zip downloaded
	;3 downloads in progress

	Loop 
	{
		WinActivate, PMS INDRA - Internet Explorer
		WinWaitActive, PMS INDRA - Internet Explorer, , 0
		if !ErrorLevel
			break
		sleep, 50
	}
return


;If perform tasks if idle time is too long
;=========================================
trapper:
	If (A_TimeIdle > 30000 AND WinExist ("View Downloads - Internet Explorer")) 
	{
		sleep, 1000
		FileAppend, trapper activated %A_YYYY% %A_MMM% %A_DD% %A_Hour%:%A_Min%:%A_Sec% %DownloadItem%`n, %DownloadFolder%Downloaded.txt
		Gosub, ReloadINDRAandResetDownload
	}
return





;Reload INDRA and deleted lasted download item 
; if idle time is more than 10 minutes
; and all download is done. 
;==============================================
ReloadINDRAandResetDownload:
	Gosub, CloseAllINDRAWindows    
	Gosub, CloseAllINDRAWindows    
	Gosub, CloseAllINDRAWindows    

	Gosub, LoadINDRA
	Loop 
	{
		If (UpdateMode = 1)
			break
		IfExist, %DownloadItem%
			FileRemoveDir, %DownloadItem%, 1
		IfNotExist, %DownloadItem%
			break
		sleep 50
	}
	Gosub, StartDownload
return