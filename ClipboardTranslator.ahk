#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; @Author Warthog
; 12/3/17
;
; Monitors the clipboard, and tries to translate it when detect change in it.
; For use with Capture2Text http://capture2text.sourceforge.net/
; Or ABBYY screenshot reader
;
; Google Translation script modified from Benny-D, p3trus - https://autohotkey.com/boards/viewtopic.php?t=21105
; Log Window code taken and modified from gamax92 - https://autohotkey.com/board/topic/49104-hot-to-make-a-scrolling-log-window/
; Resize function (AutoXYWH()) Taken from tmplinshi, toralf - https://autohotkey.com/boards/viewtopic.php?f=6&t=1079
;
; Issues:
; May not give translation or gives translation of previous phrase, particularly if its a long phrase, just try again
; If script not exit properly, an instance of IE might be left running

; IE instance for google translate
IE := ComObjCreate("InternetExplorer.Application")
; IE.visible := true
autoTranslateFlag = 1
pronounciationFlag = 1

OnExit("ExitFunction")

Gui, +AlwaysOnTop
Gui, +Resize -MaximizeBox
Gui, Add, Edit, Readonly x10 y10 w400 h270 vTranslationLog
GUI, Add, Edit, Readonly x10 y285 w400 h25 vClipboardBox
Gui, Add, Text, Readonly x10 y360 w225 vStatus, Loading...
Gui, Add, DropDownList, Sort x10 y315 w100 gUpdateSelection vLangIn, auto||en|ja|zh-CN|zh-TW
Gui, Add, DropDownList, Sort x120 y315 w100 gUpdateSelection vLangOut, en||ja|zh-CN|zh-TW
Gui, Add, Button, x250 y315 w80 h40 gTranslateClipboard vTranslateBtn, &Translate
Gui, Add, Button, x330 y315 w80 h40 gClearLog vClearBtn, &Clear
Gui, Add, Checkbox, x10 y340 gToggleAutoTranslate Checked vAutoTranslateCheckbox, Auto translate clipboard
Gui, Add, CheckBox, x145 y340 gTogglePronounciation Checked vPronounciationCheckbox, Pronounciation
Gui, Show, w420 h380, Translation Log

UpdateClipboardBox()
UpdateSelection()
return

OnClipboardChange:
UpdateClipboardBox()
if(autoTranslateFlag == 1) {
	TranslateClipboard()
}
return

; Translates whatever's in the clipboard
TranslateClipboard()
{
	global LangIn
	global LangOut
	global ClipboardBox
	phrase = %ClipboardBox%
	LogAppend(phrase)

	ChangeStatus("Translating..." . " (" LangIn . " -> " . LangOut . ")")
	LogAppend(GoogleTranslate(phrase,LangIn,LangOut))
	LogAppend("")

	UpdateSelection()
	return
}

GoogleTranslate(phrase,LangIn,LangOut)
{
	global IE
	global pronounciationFlag

	base := "https://translate.google.com.tw/?hl=en&tab=wT#"
	path := base . LangIn . "/" . LangOut . "/" . phrase	
	IE.Navigate(path)

	While IE.readyState!=4 || IE.document.readyState!="complete" || IE.busy
			Sleep 50

	; Give Google time to translate (workaround)
	Sleep 500

	Result := IE.document.all.result_box.innertext
	
	if(pronounciationFlag == 1) {
		PhrasePronounciation := IE.document.getElementById("src-translit").innertext
		ResultPronounciation := IE.document.getElementById("res-translit").innertext
		if(StrLen(PhrasePronounciation)!=0){
			Result = %PhrasePronounciation%`n%Result%
		}
		if(ResultPronounciation){
			Result = %Result%`n%ResultPronounciation%
		}
	}
	return Result
}

LogAppend(Data)
{
	GuiControlGet, TranslationLog
	GuiControl,, TranslationLog, %TranslationLog%%Data%`r`n
	ControlSend,Edit1,^{End},Translation Log
	return
}

ChangeStatus(String)
{
	GuiControlGet, Status
	GuiControl,, Status, %String%
	return
}

UpdateClipboardBox()
{
	ChangeStatus("Reading clipboard...")
	GuiControlGet, ClipboardBox
	; Get contents in clipboard as plain text
	GuiControl,, ClipboardBox, %clipboard%
	UpdateSelection()
	return
}

UpdateSelection()
{
	Gui, Submit, NoHide
	GuiControlGet, LangIn
	GuiControlGet, LangOut
	ChangeStatus("Ready. " . " (" LangIn . " -> " . LangOut . ")")
	return
}

; Used because checkbox isn't updating variables fast enough
ToggleAutoTranslate()
{
	global autoTranslateFlag
	autoTranslateFlag := !autoTranslateFlag
	return
}

TogglePronounciation()
{
	global pronounciationFlag
	pronounciationFlag := !pronounciationFlag
	return
}

ClearLog()
{
	GuiControlGet, TranslationLog
	GuiControl,, TranslationLog,
	ChangeStatus("Cleared translation log.")
	return
}

ExitFunction()
{
	global IE
	Gui, Destroy
	IE.Quit
}

GuiClose:
ExitApp
return

GuiSize:
	If (A_EventInfo = 1) ; The window has been minimized.
		Return
	AutoXYWH("wh", "TranslationLog")
	AutoXYWH("yw", "ClipboardBox")
	AutoXYWH("xy", "TranslateBtn")
	AutoXYWH("xy", "ClearBtn")
	AutoXYWH("y", "LangIn")
	AutoXYWH("y", "LangOut")
	AutoXYWH("y", "AutoTranslateCheckbox")
	AutoXYWH("y", "PronounciationCheckbox")
	AutoXYWH("yw", "Status")
return

; Taken from https://autohotkey.com/boards/viewtopic.php?f=6&t=1079

; =================================================================================
; Function: AutoXYWH
;   Move and resize control automatically when GUI resizes.
; Parameters:
;   DimSize - Can be one or more of x/y/w/h  optional followed by a fraction
;             add a '*' to DimSize to 'MoveDraw' the controls rather then just 'Move', this is recommended for Groupboxes
;   cList   - variadic list of ControlIDs
;             ControlID can be a control HWND, associated variable name, ClassNN or displayed text.
;             The later (displayed text) is possible but not recommend since not very reliable 
; Examples:
;   AutoXYWH("xy", "Btn1", "Btn2")
;   AutoXYWH("w0.5 h 0.75", hEdit, "displayed text", "vLabel", "Button1")
;   AutoXYWH("*w0.5 h 0.75", hGroupbox1, "GrbChoices")
; ---------------------------------------------------------------------------------
; Version: 2015-5-29 / Added 'reset' option (by tmplinshi)
;          2014-7-03 / toralf
;          2014-1-2  / tmplinshi
; requires AHK version : 1.1.13.01+
; =================================================================================
AutoXYWH(DimSize, cList*){       ; http://ahkscript.org/boards/viewtopic.php?t=1079
  static cInfo := {}
 
  If (DimSize = "reset")
    Return cInfo := {}
 
  For i, ctrl in cList {
    ctrlID := A_Gui ":" ctrl
    If ( cInfo[ctrlID].x = "" ){
        GuiControlGet, i, %A_Gui%:Pos, %ctrl%
        MMD := InStr(DimSize, "*") ? "MoveDraw" : "Move"
        fx := fy := fw := fh := 0
        For i, dim in (a := StrSplit(RegExReplace(DimSize, "i)[^xywh]")))
            If !RegExMatch(DimSize, "i)" dim "\s*\K[\d.-]+", f%dim%)
              f%dim% := 1
        cInfo[ctrlID] := { x:ix, fx:fx, y:iy, fy:fy, w:iw, fw:fw, h:ih, fh:fh, gw:A_GuiWidth, gh:A_GuiHeight, a:a , m:MMD}
    }Else If ( cInfo[ctrlID].a.1) {
        dgx := dgw := A_GuiWidth  - cInfo[ctrlID].gw  , dgy := dgh := A_GuiHeight - cInfo[ctrlID].gh
        For i, dim in cInfo[ctrlID]["a"]
            Options .= dim (dg%dim% * cInfo[ctrlID]["f" dim] + cInfo[ctrlID][dim]) A_Space
        GuiControl, % A_Gui ":" cInfo[ctrlID].m , % ctrl, % Options
} } }