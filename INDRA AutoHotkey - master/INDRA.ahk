;===================================
;GUI
;===================================
Gui,+AlwaysOnTop
Gui, Add, CheckBox, gCheck vMyCheckBox, Stop when this item finished (F1)
Gui, Show, w200 h80, INDRA
MyCheckBox:= 0
Gui, Add, Button, Default gStartDownload, Start Download
Gui, Add, Button, Default gContinueDownload, Continue Download
return

GuiEscape: 
GuiClose: 
ExitApp
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

^7::
send, {right} {down}
send, =LEFT(A2,10)
send, {enter} {up}

send, ^+{down}
;fill
click 102, 59
sleep 500
click 1115, 38
sleep 500
click 50, 14
sleep 500

;copy paste
click 102, 59
sleep 500
click 77, 39
sleep 500
click 25, 65
sleep 500
click 17, 123
sleep 500

;remove duplicate
click 379, -16
sleep 500
click 889, 37
sleep 500
WinWaitActive, Remove Duplicates
click 29, 122
sleep 500
click 240, 158
Sleep 1000
WinWaitActive, Remove Duplicates
click 315, 254
sleep 500
WinWaitActive, Microsoft Excel
click 218, 97
sleep 500
return

StartDownload:
    Gosub, CheckInitialCondition
    InputBox, DownloadFolder, Download Folder, Please enter download folder name , , , , , , , , F:\INDRA\

NextItem:
    WinActivate, id.xlsx - Excel
    WinWaitActive, id.xlsx - Excel
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
        If (StrLen(MyFolder1) < 3)
            break

        Loop {
            WinActivate, id.xlsx - Excel
            WinWaitActive, id.xlsx - Excel
            Send, {Right}

            Clipboard =
            MyFolder3 = 
            ; MyFolder3, vd: "D-000-1225", "VP215A-501"
            Send ^c
            ClipWait 1
            MyFolder3 := Clipboard
            If (StrLen(MyFolder3) < 5)
                break
            MyFolder3 := SubStr(MyFolder3, 1 , 10)

            MyFolder2 := SubStr(MyFolder3, 1 , InStr(MyFolder3, "-", , 4) - 1)
            ; MyFolder2, vd: "D-000", "VP215A"

            IfNotExist, %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\
            {
                FileCreateDir, %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\
                Clipboard= %MyFolder2%
                DownloadItem= %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\
                Gosub, DownloadTextFile2
            }

            IfNotExist, %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\%MyFolder3%\
            {
                FileCreateDir, %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\%MyFolder3%\
                Clipboard= %MyFolder3%
                DownloadItem= %DownloadFolder%%MyFolder0%\%MyFolder1%\%MyFolder2%\%MyFolder3%\
                Gosub, SearchAndDownloadAll2
            }
            If MyCheckBox
                return
        }
        WinActivate, id.xlsx - Excel
        WinWaitActive, id.xlsx - Excel
        Send, {Left}
        Send, ^{Left}
        Send, {Down}
        Send, {Left}
    }
return

ContinueDownload:
    Gosub, CheckInitialCondition
    If !MyCheckBox
        InputBox, DownloadFolder, Download Folder, Please enter download folder name , , , , , , , , F:\INDRA\
    InputBox, DownloadItem, Destination, Please enter download folder name , , , , , , , , F:\INDRA\
    Gosub, DownloadWithOutSearching
    If !MyCheckBox
        Gosub, NextItem
return

;===================================
;SUB ROUTINES
;===================================

CheckInitialCondition:
If ErrorLevel
    return
Loop {
WinActivate PMS INDRA - Internet Explorer
WinWaitActive, PMS INDRA - Internet Explorer, , 0
If !ErrorLevel
    break
else
    {
    MsgBox, 1, INDRA, Please open INDRA in IE.
    IfMsgBox Cancel
        return
    }
}

Loop {
WinActivate, id.xlsx - Excel
WinWaitActive, id.xlsx - Excel, , 0
If !ErrorLevel
    break
else
    {
    MsgBox, 1, INDRA, Please open id.xlsx and select the first cell.
    IfMsgBox Cancel
        return
    }
}
return

SearchAndDownloadAll2:
Gosub, DownloadTextFile2
;search file
WinActivate PMS INDRA - Internet Explorer
WinWaitActive, PMS INDRA - Internet Explorer
Click 271, 201
DownloadWithOutSearching:
Loop
{
Gosub, DownloadOnePage
    ; exit loop and minimize
    if (toValue = totalRecordsNumber){
        Gosub, FinishDownload0
        return
    }
    sleep, 50
    WinActivate PMS INDRA - Internet Explorer
    WinWaitActive, PMS INDRA - Internet Explorer
    Click 301, 219
    ;Wait for IE window become blank.
    Loop
    {
    PixelSearch, Px, Py, 398, 125, 419, 137, 0xFFFFFF, 3, RGB
    if !ErrorLevel
        break
    sleep, 50
    }
}
return

InputData2:
;switch to INDRA
WinActivate PMS INDRA - Internet Explorer
WinWaitActive, PMS INDRA - Internet Explorer
sleep 500
Loop
{
PixelSearch, Px, Py, 943, 64, 995, 89, 0xFF0000, 3, RGB
if !ErrorLevel
    break
sleep, 50
}
;load INDRA	
Click 148, 211
;Wait for IE window become blank.
Loop
{
PixelSearch, Px, Py, 398, 125, 419, 137, 0xFFFFFF, 3, RGB
if !ErrorLevel
    break
sleep, 50
}
WinActivate PMS INDRA - Internet Explorer
WinWaitActive, PMS INDRA - Internet Explorer
;Wait for IE window loaded.
Loop
{
PixelSearch, Px, Py, 361, 495, 395, 507, 0x666666, 3, RGB
if !ErrorLevel
    break
sleep, 50
}
WinActivate PMS INDRA - Internet Explorer
WinWaitActive, PMS INDRA - Internet Explorer
;input search value and option
Loop, 3
{
    Send, {Tab}
}	
Send, ^v
Send, `%
Loop, 10
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

DownloadTextFile2:
;skip prepare data
Gosub, InputData2
;download text file
WinActivate PMS INDRA - Internet Explorer
Click 358, 199
Gosub, Download2
Gosub, FinishDownload2
sleep 100
return

DownloadOnePage:
    ; check if IE is loaded.
    Loop
    {
        PixelSearch, Px, Py, 811, 280, 870, 299, 0x666666 , 5, RGB
        if !ErrorLevel
            break
        sleep, 50
    }
    ;get number of record

    Loop {
    Clipboard =
    MyString =
    WinActivate PMS INDRA - Internet Explorer
    WinWaitActive, PMS INDRA - Internet Explorer
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

    pageRecordValue:= toValue - fromValue + 1

    totalRecordsStart:= pos2 + 2
    totalRecordsLength:= pos3 - pos2 -3
    totalRecordsNumber:= SubStr(MyString, totalRecordsStart, totalRecordsLength)
    if pageRecordValue > 0
        break
    }

    ;download all if pageRecordValue <31
    if (pageRecordValue <31) {
        Gosub, DownThemAll
    }
    else {
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

DownloadItems:
WinActivate PMS INDRA - Internet Explorer
Click 753, 418
Loop, 10 {
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
Sleep 100
Click 257, 130
Click 317, 175
Gosub, Download1
sleep 100
Gosub, FinishDownload2
Gosub, FinishDownload1
return

DownThemAll:
WinActivate PMS INDRA - Internet Explorer
Click 257, 130
Click 314, 239
Click 257, 130
Click 317, 175
Gosub, Download1
sleep 100
Gosub, FinishDownload2
Gosub, FinishDownload1
return

Download1:
Loop {
    WinWaitActive, Download E-File - Internet Explorer, , 15
    If (ErrorLevel = 0)
        break
    WinActivate, This page can’t be displayed - Internet Explorer
    WinWaitActive, This page can’t be displayed - Internet Explorer, , 0
    If (ErrorLevel = 0)
        {
        Gosub, FinishDownload1ByError
        WinActivate PMS INDRA - Internet Explorer
        Click 257, 130
        Click 317, 175
        }
}
; wait for loading
Loop
{
  PixelSearch, Px, Py, 108, 301, 115, 309, 0x212121, 5, RGB
  if !ErrorLevel
     break
  sleep, 50
}
Click 156, 203

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
WinWaitActive, File Transfer for PMS - Internet Explorer, , 0
If (ErrorLevel = 0)
    break
WinWaitActive, Message from webpage, , 0
If (ErrorLevel = 0)
    {
    click 319, 139
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
sleep 500
; loop until SaveAs window appear
Loop {
  Loop {
  WinActivate, File Transfer for PMS - Internet Explorer
  WinWaitActive, File Transfer for PMS - Internet Explorer, , 0
  Click 320, 610
  Sleep 100
  WinWaitActive, File Transfer for PMS - Internet Explorer, , 0
  if !ErrorLevel
     break
  }  
  Send, {down}
  Send, {enter}
  WinWait, Save As, , 1
  If (ErrorLevel = 0)
    break
}
; loop until SaveAs window dis-appear
Loop {
WinActivate, Save As
WinWaitActive, Save As, , 0
If (ErrorLevel = 0) {
Click 398, 48
Send, ^v
send {enter}
Click 515, 447
}
WinWait, Save As, , 0
If (ErrorLevel = 1)
    break
}
return

FinishDownload2:
WinActivate, File Transfer for PMS - Internet Explorer
WinWaitActive, File Transfer for PMS - Internet Explorer
MouseMove 449, 15
Loop
{
  PixelSearch, Px, Py, 429, 3, 549, 89, 0xFFFFFF, 5, RGB
  if !ErrorLevel
     break
  sleep, 50
}
Click 449, 15
sleep, 200
return

FinishDownload1:
WinActivate, Download E-File - Internet Explorer
WinWaitActive, Download E-File - Internet Explorer
FinishDownload1ByError:
MouseMove 420, 17
Loop
{
  PixelSearch, Px, Py, 397, 2, 439, 27, 0xE81123, 5, RGB
  if !ErrorLevel
     break
  sleep, 50
}
Click 420, 17
sleep, 200
return

FinishDownload0:
WinActivate, PMS INDRA - Internet Explorer
WinWaitActive, PMS INDRA - Internet Explorer
Send, {Alt down}
Send, {space}
Send, {Alt up}
Sleep 50
Send, n	
return