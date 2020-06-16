; text-capture-acc.ahk - main file
;
; Charset for this file must be UTF 8 with BOM.
; it may not function properly otherwise.
;
; Script written for AHK_L v1.1.28 Unicode.
;
; Disclaimer: this script is provided "as is", without any kind of warranty.
; The author(s) shall not be liable for any damage caused by using
; this script or its derivatives,  et cetera.
;
; =====================
; GENERAL OVERVIEW
; =====================
;
; Compilation directives; include files in binary and set file properties
; ===========================================================
;
;@Ahk2Exe-SetName Text-Capture-ACC
;@Ahk2Exe-SetCopyright Marius Şucan (2017-2018)
;@Ahk2Exe-SetCompanyName sucan.ro

;================================================================
; Section 0. Auto-exec.
;================================================================

; Script Initialization

 #SingleInstance Force
 #NoEnv
 #MaxMem 128
 #ClipboardTimeout 3000
 DetectHiddenWindows, On
 ; #Warn Debug
 ComObjError(false)
 SetTitleMatchMode, 2
 SetBatchLines, -1
 ListLines, Off
 SetWorkingDir, %A_ScriptDir%
 Critical, On

; Default Settings

 Global IniFile           := "text-capture-acc.ini"
 , Copy2Clip              := 1 
 , showACCdetails         := 1

; OSD settings
 , DisplayTimeUser        := 3     ; in seconds
 , GuiX                   := 40
 , GuiY                   := 250
 , FontName               := (A_OSVersion="WIN_XP" && FileExist(A_WinDir "\Fonts\ARIALUNI.TF")) ? "Arial Unicode MS" : "Arial"
 , FontSize               := 19
 , PrefsLargeFonts        := 0
 , OSDbgrColor            := "131209"
 , OSDtextColor           := "FFFEFA"
 , OSDalpha               := 230
 , OSDmarginTop           := 20
 , OSDmarginBottom        := 20
 , OSDmarginSides         := 25
 , maxMainLength         := 65

; Script's own global shortcuts (hotkeys)
 , GlobalKBDhotkeys       := 1     ; Enable system-wide shortcuts (hotkeys)
 , KBDCapText             := "Pause"
 , KBDCapTextConstant     := "^Pause"

 , ShowPreview            := 0     ; let it be a persistent setting
 , ThisFile               := A_ScriptName

; Release info
 , Version                := "0.6"
 , ReleaseDate            := "2018 / 10 / 26"
 , ScriptInitialized, FirstRun := 1

; Check if INIT previously failed or if KP is running and then load settings.
; These functions are in Section 8.

    INIaction(0, "FirstRun", "SavedSettings")
    If (FirstRun=0)
    {
        INIsettings(0)
    } Else
    {
        CheckSettings()
        INIsettings(1)
    }

; Initialization variables. Altering these may lead to undesired results.

Global Debug := 0    ; for testing purposes
 , DisplayTime := DisplayTimeUser*1000
 , MainGuiVisible := 0
 , AccTextCaptureActive := 0
 , LastMainQuoteDisplay := 1    ; timer to keep track of OSD redraws
 , Tickcount_start := 0               ; timer to count repeated key presses
 , MousePosition := ""
 , DoNotRepeatTimer := 0
 , lastMsgDisplayied := ""
 , PrefOpen := 0
 , FontList := []
 , LargeUIfontValue := 13
 , InstKBDsWinOpen, CurrentTab, AnyWindowOpen := 0
 , PreviewWindowText := "Preview " Lola "window... " Lola2
 , GlobalKBDsList := "KBDCapText,KBDCapTextConstant"
 , KeysComboList := "(Disabled)|(Restore Default)|[[ 0-9 / Digits ]]|[[ Letters ]]|Right|Left|Up|Down|Home|End
    |Page_Down|Page_Up|Backspace|Space|Tab|Delete|Enter|Escape|Insert|CapsLock|NumLock|ScrollLock|L_Click
    |M_Click|R_Click|PrintScreen|Pause|Break|CtrlBreak|AppsKey|F1|F2|F3|F4|F5|F6|F7|F8|F9|F10|F11|F12
    |Nav_Back|Nav_Favorites|Nav_Forward|Nav_Home|Nav_Refresh|Nav_Search|Nav_Stop|Help|Launch_App1
    |Launch_App2|Launch_Mail|Launch_Media|Media_Next|Media_Play_Pause|Media_Prev|Media_Stop|Pad0|Pad1
    |Pad2|Pad3|Pad4|Pad5|Pad6|Pad7|Pad8|Pad9|PadClear|PadDel|PadDiv|PadDot|PadHome|PadEnd|PadEnter
    |PadIns|PadLeft|PadRight|PadAdd|PadSub|PadMult|PadPage_Down|PadPage_Up|PadUp|PadDown|Sleep
    |Volume_Mute|Volume_Up|Volume_Down|WheelUp|WheelDown|WheelLeft|WheelRight|[[ VK nnn ]]|[[ SC nnn ]]"
 , hMainOSD, ColorPickerHandles
 , hMain := A_ScriptHwnd
 , CCLVO := "-E0x200 +Border -Hdr -Multi +ReadOnly Report AltSubmit gsetColors"
 , BaseURL := "http://marius.sucan.ro/media/files/blog/ahk-scripts/"
 , hWinMM := DllCall("kernel32\LoadLibraryW", "Str", "winmm.dll", "Ptr")
 , ScriptelSuspendel := 0
 , ForceUpdate := 0     ; this will be used when major changes require full update

; Initializations of the core components and functionality

Sleep, 5
CreateGlobalShortcuts()
InitializeTray()

hCursM := DllCall("user32\LoadCursorW", "Ptr", NULL, "Int", 32646, "Ptr")  ; IDC_SIZEALL
hCursH := DllCall("user32\LoadCursorW", "Ptr", NULL, "Int", 32649, "Ptr")  ; IDC_HAND
OnMessage(0x404, "AHK_NOTIFYICON")
OnMessage(0x200, "MouseMove")    ; WM_MOUSEMOVE
Sleep, 5
ScriptInitialized := 1      ; the end of the autoexec section and INIT
Return

;================================================================
; Section 1. The OSD GUI - CreateOSDGUI()
; - GetTextExtentPoint() and GuiGetSize() are constantly used
;   to determine text and window sizes.
;================================================================

stripText(txt) {
    StringReplace, txt, txt, %A_SPACE%%A_SPACE%, %A_SPACE%, All
    StringReplace, txt, txt, `r`n, %A_Space%, All
    StringReplace, txt, txt, `n, %A_Space%, All
    StringReplace, txt, txt, `r, %A_Space%, All
    StringReplace, txt, txt, `f, %A_Space%, All
    StringReplace, txt, txt, %A_TAB%, %A_SPACE%, All
    StringReplace, txt, txt, %A_SPACE%%A_SPACE%, %A_SPACE%, All
    txt := RegExReplace(txt, "\s+", A_Space)
    If (txt=A_Space)
       txt := ""

    Return txt
}

CreateMainGUI(msg2Display) {
    Critical, On
    msg2Display := stripText(msg2Display)
    If msg2Display
       lastMsgDisplayied := msg2Display
    Else 
       Return
    msg2Display := ST_wordWrap(msg2Display, maxMainLength)
    msg2Display := ST_LineWrap(msg2Display, maxMainLength+1)
    Gui, MainGui: Destroy
    Sleep, 25
    If (PrefOpen=0)
       Global LastMainQuoteDisplay := A_TickCount
    HorizontalMargins := OSDmarginTop
    Gui, MainGui: -DPIScale -Caption +Owner +ToolWindow +HwndhMainOSD
    Gui, MainGui: Margin, %OSDmarginSides%, %HorizontalMargins%
    Gui, MainGui: Color, %OSDbgrColor%
    If (FontChangedTimes>190)
       Gui, MainGui: Font, c%OSDtextColor% s%FontSize% Bold,
    Else
       Gui, MainGui: Font, c%OSDtextColor% s%FontSize% Bold, %FontName%
    Gui, MainGui: Add, Text, hwndhMainTxt, %msg2Display%
    Gui, MainGui: Show, NoActivate AutoSize x%GuiX% y%GuiY%, MainWin
    WinSet, Transparent, %OSDalpha%, MainWin
    WinSet, AlwaysOnTop, On, MainWin
    MainGuiVisible := 1
    quoteDisplayTime := (PrefOpen=1) ? DisplayTime*1.5 : StrLen(msg2Display) * 100 + 1000
    If (PrefOpen!=1)
       SetTimer, DestroyMainGui, % -quoteDisplayTime
}

ST_LineWrap(string, column= 56, indentChar= "") {
; String Things - Common String & Array Functions, 2014
; by tidbit https://autohotkey.com/board/topic/90972-string-things-common-text-and-array-functions/


    CharLength := StrLen(indentChar)
    , columnSpan := column - CharLength
    , Ptr := A_PtrSize ? "Ptr" : "UInt"
    , NewLineType := A_IsUnicode ? "UShort" : "UChar"
    , UnicodeModifier := A_IsUnicode ? 2 : 1
    , VarSetCapacity(out, (StrLen(string) + (Ceil(StrLen(string) / columnSpan) * (column + CharLength + 1))) * UnicodeModifier, 0)
    , A := &out
     
    loop, parse, string, `n, `r
        If ((FieldLength := StrLen(ALoopField := A_LoopField)) > column)
        {
            DllCall("RtlMoveMemory", Ptr, A, Ptr, &ALoopField, "UInt", column * UnicodeModifier)
            , A += column * UnicodeModifier
            , NumPut(10, A+0, 0, NewLineType)
            , A += UnicodeModifier
            , Pos := column
             
            While (Pos < FieldLength)
            {
                If CharLength
                    DllCall("RtlMoveMemory", Ptr, A, Ptr, &indentChar, "UInt", CharLength * UnicodeModifier)
                    , A += CharLength * UnicodeModifier
                 
                If (Pos + columnSpan > FieldLength)
                    DllCall("RtlMoveMemory", Ptr, A, Ptr, &ALoopField + (Pos * UnicodeModifier), "UInt", (FieldLength - Pos) * UnicodeModifier)
                    , A += (FieldLength - Pos) * UnicodeModifier
                    , Pos += FieldLength - Pos
                Else
                    DllCall("RtlMoveMemory", Ptr, A, Ptr, &ALoopField + (Pos * UnicodeModifier), "UInt", columnSpan * UnicodeModifier)
                    , A += columnSpan * UnicodeModifier
                    , Pos += columnSpan
                 
                NumPut(10, A+0, 0, NewLineType)
                , A += UnicodeModifier
            }
        } Else
            DllCall("RtlMoveMemory", Ptr, A, Ptr, &ALoopField, "UInt", FieldLength * UnicodeModifier)
            , A += FieldLength * UnicodeModifier
            , NumPut(10, A+0, 0, NewLineType)
            , A += UnicodeModifier
     
    VarSetCapacity(out, -1)
    Return SubStr(out,1, -1)
}

ST_wordWrap(string, column=56, indentChar="") {
; String Things - Common String & Array Functions, 2014
; by tidbit https://autohotkey.com/board/topic/90972-string-things-common-text-and-array-functions/
; fixed by Marius Șucan, such that it does not give Continuable Exception Error on some systems

    indentLength := StrLen(indentChar)
    Loop, Parse, string, `n
    {
        If (StrLen(A_LoopField) > column)
        {
            pose := 1
            Loop, Parse, A_LoopField, %A_Space%
            {
                loopLength := StrLen(A_LoopField)
                If (pose + loopLength <= column)
                {
                   out .= (A_Index = 1 ? "" : " ") A_LoopField
                   pose += loopLength + 1
                } Else
                {
                   pose := loopLength + 1 + indentLength
                   out .= "`n" indentChar A_LoopField
                }
            }
            out .= "`n"
        } Else
            out .= A_LoopField "`n"
    }
    result := SubStr(out, 1, -1)
    Return result
}

DestroyMainGui() {
  Gui, MainGui: Destroy
  MainGuiVisible := 0
}

MouseMove(wP, lP, msg, hwnd) {
; Function by Drugwash
  Global
  Local A
  SetFormat, Integer, H
  hwnd+=0, A := WinExist("A"), hwnd .= "", A .= ""
  SetFormat, Integer, D

  If InStr(hMainOSD, hwnd) && (A_TickCount - LastMainQuoteDisplay>700) && (A_TimeIdle<200)
  {
        If (PrefOpen=0)
           DestroyMainGui()
        DllCall("user32\SetCursor", "Ptr", hCursM)
        If !(wP&0x13)    ; no LMR mouse button is down, we hover
        {
           If A not in %hMainOSD%
              hAWin := A
        } Else If (wP&0x1)  ; L mouse button is down, we're dragging
        {
           SetTimer, DestroyMainGui, Off
           While GetKeyState("LButton", "P")
           {
              PostMessage, 0xA1, 2,,, ahk_id %hMainOSD%
              DllCall("user32\SetCursor", "Ptr", hCursM)
           }
           SetTimer, trackMouseDragging, -1
           Sleep, 0
        } Else If ((wP&0x2) || (wP&0x10))
           DestroyMainGui()
  } Else If ColorPickerHandles
  {
     If hwnd in %ColorPickerHandles%
        DllCall("user32\SetCursor", "Ptr", hCursH)
  }
}

trackMouseDragging() {
; Function by Drugwash
  Global
  WinGetPos, NewX, NewY,,, ahk_id %hMainOSD%

  GuiX := !NewX ? "2" : NewX
  GuiY := !NewY ? "2" : NewY

  If hAWin
  {
     If hAWin not in %hMainOSD%
        WinActivate, ahk_id %hAWin%
  }
  saveGuiPositions()
}

saveGuiPositions() {
; function called after dragging the OSD to a new position

  If (PrefOpen=0)
  {
     Sleep, 700
     SetTimer, DestroyMainGui, -1500
     INIaction(1, "GuiX", "OSDprefs")
     INIaction(1, "GuiY", "OSDprefs")
  } Else If (PrefOpen=1)
  {
     GuiControl, SettingsGUIA:, GuiX, %GuiX%
     GuiControl, SettingsGUIA:, GuiY, %GuiY%
  }
}

;================================================================
; Section 5. features invoked by keyboard shortcuts
; - The hotkeys registered replace the system default
;================================================================

RegisterGlobalShortcuts(HotKate,destination,apriori) {
   testHotKate := RegExReplace(HotKate, "i)^(\!|\^|\#|\+)$", "")
   If (InStr(HotKate, "disa") || StrLen(HotKate)<1)
   {
      HotKate := "(Disabled)"
      Return HotKate
   }

   If (GlobalKBDsNoIntercept=1 || InStr(HotKate, "button"))
   {
      HotKate := "~" HotKate
      apriori := "~" apriori
   }

   Hotkey, %HotKate%, %destination%, UseErrorLevel
   If (ErrorLevel!=0)
   {
      Hotkey, %apriori%, %destination%, UseErrorLevel
      Return apriori
   }
   Return HotKate
}

CreateGlobalShortcuts() {
    If (GlobalKBDhotkeys=1)
    {
       KBDCapText := RegisterGlobalShortcuts(KBDCapText,"AccCaptureTextNow", "Pause")
       KBDCapTextConstant := RegisterGlobalShortcuts(KBDCapTextConstant,"ToggleAccCaptureText", "^Pause")
    }
}

SuspendScriptNow() {
  SuspendScript(0)
}

SuspendScript(partially:=0) {
   Suspend, Permit
   Thread, Priority, 150
   Critical, On

   If (SecondaryTypingMode=1)
      Return

   If (PrefOpen=1 && A_IsSuspended=1)
   {
      SoundBeep, 300, 900
      Return
   }
 
   If (AccTextCaptureActive=1)
      ToggleAccCaptureText()
   Sleep, 50
   Menu, Tray, UseErrorLevel
   Menu, Tray, Rename, &Text Capture ACC activated,&Text Capture ACC deactivated
   If (ErrorLevel=1)
   {
      Menu, Tray, Rename, &Text Capture ACC deactivated,&Text Capture ACC activated
      Menu, Tray, Check, &Text Capture ACC activated
   }
   Menu, Tray, Uncheck, &Text Capture ACC deactivated

   friendlyName := A_IsSuspended ? "activated" : "deactivated"
   CreateMainGUI("Text Capture ACC " friendlyName)
   Suspend
}

AccCaptureTextNow() {
  Static lastInvoked := 1, timesInvoked := 0, prevLastMsgDisplayied
  
  If (InStr(KBDCapText, "Button") && A_TimeIdle<1200 && MainGuiVisible=0 && AccTextCaptureActive=0)
  {
     SetTimer, AccCaptureTextNow, -1500
     Return
  }

  If (A_TickCount - lastInvoked < 400) && (timesInvoked>1)
  || (A_TickCount - lastInvoked < 400) && (AccTextCaptureActive=1)
  {
     SoundBeep
     ToggleAccCaptureText()
     timesInvoked := 0
     Return
  }

  If (A_TickCount - lastInvoked < 400) ; 
  {
     iF (prevLastMsgDisplayied=lastMsgDisplayied && StrLen(lastMsgDisplayied)>1) && (A_TickCount - DoNotRepeatTimer > 200)
     {
        SoundBeep
        Clipboard := lastMsgDisplayied
     }
     timesInvoked++
     Return
  }

  Global DoNotRepeatTimer := A_TickCount
  GetAccInfo(1)
  prevLastMsgDisplayied := lastMsgDisplayied
  lastInvoked := A_TickCount
}

ReloadScriptNow() {
    ReloadScript(0)
}

;================================================================
; Section 6. Tray menu and related functions.
;================================================================

InitializeTray() {
    Menu, PrefsMenu, Add, &Customize, ShowOSDsettings
    Menu, PrefsMenu, Add
    Menu, PrefsMenu, Add, L&arge UI fonts, ToggleLargeFonts
    Menu, PrefsMenu, Add, R&un in Admin Mode, RunAdminMode
    Menu, PrefsMenu, Add

    If A_IsAdmin
    {
       Menu, PrefsMenu, Check, R&un in Admin Mode
       Menu, PrefsMenu, Disable, R&un in Admin Mode
    }

    If (PrefsLargeFonts=1)
       Menu, PrefsMenu, Check, L&arge UI fonts

    RunType := A_IsCompiled ? "" : " [script]"
    Menu, Tray, NoStandard
    Menu, Tray, Add, &Preferences, :PrefsMenu
    Menu, Tray, Add
    Menu, Tray, Add, Mouse text collector, ToggleAccCaptureText
    Menu, Tray, Add
    Menu, Tray, Add, &Text Capture ACC activated, SuspendScriptNow
    Menu, Tray, Check, &Text Capture ACC activated
    Menu, Tray, Add, &Restart, ReloadScriptNow
    Menu, Tray, Add
    Menu, Tray, Add, &About, AboutWindow
    Menu, Tray, Add
    Menu, Tray, Add, E&xit, KillScript, P50
    Menu, Tray, Tip, Text Capture ACC v%Version%%RunType%
    Menu, Tray, Default, Mouse text collector
}

AHK_NOTIFYICON(wParam, lParam, uMsg, hWnd) {
  Critical, off
  Static lastInvoked := 1

  If (PrefOpen=1 || AccTextCaptureActive=1 || A_IsSuspended || NeverDisplayOSD=1)
  || (A_TickCount - lastInvoked < 900)
     Return
  CreateMainGUI("Text capture tray icon")
  lastInvoked := A_TickCount
}


ToggleAccCaptureText() {
    AccTextCaptureActive := !AccTextCaptureActive
    Menu, Tray, % (AccTextCaptureActive=0 ? "Uncheck" : "Check"), Mouse text collector
    If (AccTextCaptureActive=1)
    {
       CreateMainGUI("Text Capture activated")
       SetTimer, GetAccInfo, 200, 50
    } Else
    {
       SetTimer, GetAccInfo, off
       CreateMainGUI("Text Capture deactivated")
    }
    Sleep, 1000
}

ToggleLargeFonts() {
    PrefsLargeFonts := !PrefsLargeFonts
    INIaction(1, "PrefsLargeFonts", "SavedSettings")
    Menu, PrefsMenu, % (PrefsLargeFonts=0 ? "Uncheck" : "Check"), L&arge UI fonts
    Sleep, 200
}

ReloadScript(silent:=1) {
    Thread, Priority, 50
    Critical, On

    If (PrefOpen=1)
    {
       CloseSettings()
       Return
    }

    If FileExist(ThisFile)
    {
        Cleanup()
        Reload
        Sleep, 50
        ExitApp
    } Else
    {
        CreateMainGUI("FATAL ERROR: Main file missing. Execution terminated.")
        SoundBeep
        Sleep, 2000
        Cleanup() ; if you don't do it HERE you're not doing it right, Run %i% will force the script to close before cleanup
        MsgBox, 4,, Do you want to choose another file to execute?
        IfMsgBox, Yes
        {
            FileSelectFile, i, 2, %A_ScriptDir%\%A_ScriptName%, Select a different script to load, AutoHotkey script (*.ahk; *.ah1u)
            If !InStr(FileExist(i), "D")  ; we can't run a folder, we need to run a script
               Run, %i%
        } Else (Sleep, 500)
        ExitApp
    }
}

RunAdminMode() {
  If !A_IsAdmin
  {
      Try {
         Cleanup()
         If A_IsCompiled
            Run *RunAs "%A_ScriptFullPath%" /restart
         Else
            Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
         ExitApp
      }
  }
}

DeleteSettings() {
    MsgBox, 4,, Are you sure you want to delete the stored settings?
    IfMsgBox, Yes
    {
       FileSetAttrib, -R, %IniFile%
       FileDelete, %IniFile%
       Cleanup()
       Reload
    }
}

KillScript(showMSG:=1) {
   Thread, Priority, 50
   Critical, On
   If (ScriptInitialized!=1)
      ExitApp

   PrefOpen := 0
   If (FileExist(ThisFile) && showMSG)
   {
      INIsettings(1)
      CreateMainGUI("Bye byeee :-)")
      Sleep, 350
   } Else If showMSG
   {
      CreateMainGUI("Adiiooosss :-(((")
      Sleep, 950
   }
   Cleanup()
   ExitApp
}

;================================================================
; Section 7. Settings window.
; - In this section you can find each preferences window
;   or any other window based on SettingsGUI() and 
;   various functions used in the UI.
;================================================================

SettingsGUI() {
   Global
   Gui, SettingsGUIA: Destroy
   Sleep, 15
   Gui, SettingsGUIA: Default
   Gui, SettingsGUIA: -MaximizeBox
   Gui, SettingsGUIA: -MinimizeBox
   Gui, SettingsGUIA: Margin, 15, 15
}

initSettingsWindow() {
    Global ApplySettingsBTN
    If (PrefOpen=1)
    {
        SoundBeep, 300, 900
        doNotOpen := 1
        Return doNotOpen
    }

    If (A_IsSuspended!=1)
       SuspendScript(1)

    PrefOpen := 1
    SettingsGUI()
}

SwitchPreferences(forceReopenSame:=0) {
    testPrefWind := (forceReopenSame=1) ? "lol" : CurrentPrefWindow
    GuiControlGet, CurrentPrefWindow
    If (testPrefWind=CurrentPrefWindow)
       Return

    PrefOpen := 0
    GuiControlGet, ApplySettingsBTN, Enabled
    Gui, Submit
    Gui, SettingsGUIA: Destroy
    Sleep, 25
    SettingsGUI()
    CheckSettings()
    If (CurrentPrefWindow=5)
    {
       ShowOSDsettings()
       VerifyOsdOptions(ApplySettingsBTN)
    }
}

ApplySettings() {
    Gui, SettingsGUIA: Submit, NoHide
    CheckSettings()
    PrefOpen := 0
    INIsettings(1)
    Sleep, 100
    ReloadScript()
}

CloseWindow() {
    AnyWindowOpen := 0
    Gui, SettingsGUIA: Destroy
}

CloseSettings() {
   GuiControlGet, ApplySettingsBTN, Enabled
   GuiControlGet, CurrentTab
   PrefOpen := 0
   CloseWindow()
   If (ApplySettingsBTN=0)
   {
      Sleep, 25
      SuspendScript()
      Return
   }
   Sleep, 100
   ReloadScript()
}

SettingsGUIAGuiEscape:
   If (PrefOpen=1)
      CloseSettings()
   Else
      CloseWindow()
Return

SettingsGUIAGuiClose:
   If (PrefOpen=1)
      CloseSettings()
   Else
      CloseWindow()
Return


AddKBDmods(HotKate, HotKateRaw) {
    Global
    modBtnWidth := (PrefsLargeFonts=1) ? 45 : 32
    reused := "x+0 +0x1000 w" modBtnWidth " hp gGenerateHotkeyStrS "
    C%HotKate% := InStr(HotKateRaw, "^")
    S%HotKate% := InStr(HotKateRaw, "+")
    A%HotKate% := InStr(HotKateRaw, "!")
    W%HotKate% := InStr(HotKateRaw, "#")

    Gui, Add, Checkbox, % reused " Checked" C%HotKate% " vCtrl" HotKate, Ctrl
    Gui, Add, Checkbox, % reused " Checked" A%HotKate% " vAlt" HotKate, Alt
    Gui, Add, Checkbox, % reused " Checked" S%HotKate% " vShift" HotKate, Shift
    Gui, Add, Checkbox, % reused " Checked" W%HotKate% " vWin" HotKate, Win
}

AddKBDcombo(HotKate, HotKateRaw) {
    Global
    col2width := (PrefsLargeFonts=1) ? 140 : 90
    ComboChoice := ProcessChoiceKBD(HotKateRaw)
    Gui, Add, ComboBox, % "x+0 w"col2width " gProcessComboKBD vCombo" HotKate, %KeysComboList%|%ComboChoice%||
}

GenerateHotkeyStrS(enableApply:=1) {
  GuiControlGet, ApplySettingsBTN

  kW1 := "disa"
  kW2 := "resto"
  kWa := "(Disabled)"
  kWb := "(Restore Default)"

  Loop, Parse, GlobalKBDsList, CSV
  {
     GuiControlGet, Combo%A_LoopField%
     GuiControlGet, Ctrl%A_LoopField%
     GuiControlGet, Shift%A_LoopField%
     GuiControlGet, Alt%A_LoopField%
     GuiControlGet, Win%A_LoopField%
     %A_LoopField% := ""
     %A_LoopField% .= Ctrl%A_LoopField%=1 ? "^" : ""
     %A_LoopField% .= Shift%A_LoopField%=1 ? "+" : ""
     %A_LoopField% .= Alt%A_LoopField%=1 ? "!" : ""
     %A_LoopField% .= Win%A_LoopField%=1 ? "#" : ""
     %A_LoopField% .= ProcessChoiceKBD2(Combo%A_LoopField%)
     If InStr(Combo%A_LoopField%, kW1)
        %A_LoopField% := kWa

     If InStr(Combo%A_LoopField%, kW2)
        %A_LoopField% := kWb
  }

  keywords := "i)(disa|resto)"
  KBDsTestDuplicate := KBDCapText "&" KBDCapTextConstant
  For each, kbd2test in StrSplit(KBDsTestDuplicate, "&")
  {
      countDuplicate := 0
      Loop, Parse, KBDsTestDuplicate, &
      {
          If RegExMatch(A_LoopField, keywords)
             Continue
          If (kbd2test=A_LoopField)
             countDuplicate++
      }
      If (countDuplicate>1)
         disableButtons := 1
  }

  If (disableButtons=1)
  {
     ToolTip, Detected duplicate keyboard shorcuts...
     SoundBeep, 300, 900
     GuiControl, Disable, ApplySettingsBTN
     GuiControl, Disable, CurrentPrefWindow
     GuiControl, Disable, CancelBTN
     SetTimer, DupeHotkeysToolTipDummy, -1500
  } Else
  {
     GuiControl, % (!enableApply ? "Disable" : "Enable"), ApplySettingsBTN
     GuiControl, Enable, CurrentPrefWindow
     GuiControl, Enable, CancelBTN
  }
}

DupeHotkeysToolTipDummy() {
  ToolTip
}

ProcessComboKBD(enableApply:=1) {
  forbiddenChars := "(\~|\*|\!|\+|\^|\#|\$|\<|\>|\&)"
  keywords := "i)(\(.|^([\p{Z}\p{P}\p{S}\p{C}\p{N}].)|disa|resto|\s|\[\[|\]\])"
  GuiControlGet, activeCtrl, FocusV
  Loop, Parse, GlobalKBDsList, CSV
  {
      GuiControlGet, CbEdit%A_LoopField%,, Combo%A_LoopField%
      If RegExMatch(CbEdit%A_LoopField%, forbiddenChars)
         GuiControl,, Combo%A_LoopField%, | %KeysComboList%
      If RegExMatch(CbEdit%A_LoopField%, keywords)
         SwitchStateKBDbtn(A_LoopField, 0, 0)
  }

  StringReplace, activeCtrl, activeCtrl, ComboK, K
  If (RegExMatch(CbEdit%activeCtrl%, keywords) || StrLen(CbEdit%activeCtrl%)<1)
     SwitchStateKBDbtn(activeCtrl, 0, 0)
  Else
     SwitchStateKBDbtn(activeCtrl, 1, 0)
  
  GuiControl, % (!enableApply ? "Disable" : "Enable"), ApplySettingsBTN
  GenerateHotkeyStrS(enableApply)
}

ProcessChoiceKBD(strg) {
     Loop, Parse, % "^~#&!+<>$*"
         StringReplace, strg, strg, %A_LoopField%
     If !strg
        strg := "(Disabled)"
     Return strg
}

ProcessChoiceKBD2(strg) {
     StringReplace, strg, strg,Pad,Numpad
     StringReplace, strg, strg,Page_Up,PgUp
     StringReplace, strg, strg,Page_Down,PgDn
     StringReplace, strg, strg,Nav_,Browser_
     StringReplace, strg, strg,_Click,Button
     StringReplace, strg, strg,numnumpad,Numpad
     Return strg
}

SwitchStateKBDbtn(HotKate, do, noCombo:=1) {
    action := (do=0) ? "Disable" : "Enable"
    If (noCombo=1)
       GuiControl, %action%, Combo%HotKate%
    GuiControl, %action%, Ctrl%HotKate%
    GuiControl, %action%, Shift%HotKate%
    GuiControl, %action%, Alt%HotKate%
    GuiControl, %action%, Win%HotKate%
}

hexRGB(c) {
; unknown source
  r := ((c&255)<<16)+(c&65280)+((c&0xFF0000)>>16)
  c := "000000"
  DllCall("msvcrt\sprintf", "AStr", c, "AStr", "%06X", "UInt", r, "CDecl")
  Return c
}

Dlg_Color(Color,hwnd) {
; Function by maestrith 
; from: [AHK 1.1] Font and Color Dialogs 
; https://autohotkey.com/board/topic/94083-ahk-11-font-and-color-dialogs/
; Modified by Marius Șucan and Drugwash

  Static
  If !cpdInit {
     VarSetCapacity(CUSTOM,64,0), cpdInit:=1, size:=VarSetCapacity(CHOOSECOLOR,9*A_PtrSize,0)
  }

  Color := "0x" hexRGB(InStr(Color, "0x") ? Color : Color ? "0x" Color : 0x0)
  NumPut(size,CHOOSECOLOR,0,"UInt"),NumPut(hwnd,CHOOSECOLOR,A_PtrSize,"Ptr")
  ,NumPut(Color,CHOOSECOLOR,3*A_PtrSize,"UInt"),NumPut(3,CHOOSECOLOR,5*A_PtrSize,"UInt")
  ,NumPut(&CUSTOM,CHOOSECOLOR,4*A_PtrSize,"Ptr")
  If !ret := DllCall("comdlg32\ChooseColorW","Ptr",&CHOOSECOLOR,"UInt")
     Exit

  SetFormat, Integer, H
  Color := NumGet(CHOOSECOLOR,3*A_PtrSize,"UInt")
  SetFormat, Integer, D
  Return Color
}

setColors(hC, event, c, err=0) {
; Function by Drugwash
; Critical MUST be disabled below! If that's not done, script will enter a deadlock !
  Static
  oc := A_IsCritical
  Critical, Off
  If (event != "Normal")
     Return
  g := A_Gui, ctrl := A_GuiControl
  r := %ctrl% := hexRGB(Dlg_Color(%ctrl%, hC))
  Critical, %oc%
  GuiControl, %g%:+Background%r%, %ctrl%
  GuiControl, Enable, ApplySettingsBTN
  Sleep, 100
  OSDpreview()
}

UpdateFntNow() {
  Global
  Fnt_DeleteFont(hfont)
  fntOptions := "s" FontSize " Bold Q5"
  hFont := Fnt_CreateFont(FontName,fntOptions)
  ; Fnt_SetFont(hOSDctrl,hfont,true)
  Fnt_SetFont(hMainTxt,hfont,true)
}

OSDpreview() {
    Static LastBorderState, lastFnt := FontName

    Gui, SettingsGUIA: Submit, NoHide
    If (ShowPreview=0)
    {
       DestroyMainGui()
       Return
    }

    CreateMainGUI(PreviewWindowText)
    Sleep, 25
    If (lastFnt!=FontName)
    {
       FontChangedTimes++
       lastFnt := FontName
    }
    ; ToolTip, nr. %FontChangedTimes%
    If (FontChangedTimes>190)
       UpdateFntNow()
}

editsOSDwin() {
  If (A_TickCount-DoNotRepeatTimer<1000)
     Return
  VerifyOsdOptions()
}


ShowOSDsettings() {
    doNotOpen := initSettingsWindow()
    If (doNotOpen=1)
       Return

    Global CurrentPrefWindow := 5
    Global DoNotRepeatTimer := A_TickCount
    Global positionB, editF1, editF2, editF3, editF4, editF5, editF6, Btn1, editF60
         , editF7, editF8, editF9, editF10, editF11, editF13, editF35, editF36, editF37, Btn2
    columnBpos1 := columnBpos2 := 125
    KBDcol1width := 290
    editFieldWid := 220
    If (PrefsLargeFonts=1)
    {
       Gui, Font, s%LargeUIfontValue%
       editFieldWid := 285
       KBDcol1width := 430
       columnBpos1 := columnBpos2 := columnBpos2 + 125
    }
    columnBpos1b := columnBpos1 + 70

    Gui, Add, Tab3,, General|OSD options

    Gui, Tab, 1 ; general
    Gui, Add, Checkbox, x+15 y+15 gVerifyOsdOptions Checked%Copy2Clip% vCopy2Clip, Copy to clipboard the text `n (applies only when text is captured only once)
    Gui, Add, Checkbox, y+10 Section gVerifyOsdOptions Checked%showACCdetails% vshowACCdetails, Show extensive ACC details
    Gui, Add, Checkbox, y+10 gVerifyOsdOptions Checked%GlobalKBDhotkeys% vGlobalKBDhotkeys, Global keyboard shortcuts

    Gui, Add, Text, xs+15 y+5 w%KBDcol1width%, Capture text only once
    Gui, Add, Text, xs+50 y+1 w100, 
    AddKBDcombo("KBDCapText", KBDCapText)
    AddKBDmods("KBDCapText", KBDCapText)

    Gui, Add, Text, xs+15 y+1 w%KBDcol1width%, Constantly capture text
    Gui, Add, Text, xs+50 y+1 w100, 
    AddKBDcombo("KBDCapTextConstant", KBDCapTextConstant)
    AddKBDmods("KBDCapTextConstant", KBDCapTextConstant)

    Gui, Tab, 2 ; size/position
    Gui, Add, Text, x+15 y+15 Section, OSD position (x, y)
    Gui, Add, Edit, xs+%columnBpos2% ys w65 geditsOSDwin r1 limit4 -multi number -wantCtrlA -wantReturn -wantTab -wrap veditF1, %GuiX%
    Gui, Add, UpDown, vGuiX gVerifyOsdOptions 0x80 Range-9995-9998, %GuiX%
    Gui, Add, Edit, x+5 w65 geditsOSDwin r1 limit4 -multi number -wantCtrlA -wantReturn -wantTab -wrap veditF2, %GuiY%
    Gui, Add, UpDown, vGuiY gVerifyOsdOptions 0x80 Range-9995-9998, %GuiY%

    Gui, Add, Text, xm+15 ys+30 Section, Margins (horizontal, vertical)
    Gui, Add, Edit, xs+%columnBpos2% ys+0 Section w65 geditsOSDwin r1 limit3 -multi number -wantCtrlA -wantReturn -wantTab -wrap veditF11, %OSDmarginTop%
    Gui, Add, UpDown, gVerifyOsdOptions vOSDmarginTop Range1-900, %OSDmarginTop%
    Gui, Add, Edit, x+5 w65 geditsOSDwin r1 limit3 -multi number -wantCtrlA -wantReturn -wantTab -wrap veditF13, %OSDmarginSides%
    Gui, Add, UpDown, gVerifyOsdOptions vOSDmarginSides Range1-900, %OSDmarginSides%

    Gui, Add, Text, xm+15 y+10 Section, Font name
    Gui, Add, Text, xs yp+30, OSD colors and opacity
    Gui, Add, Text, xs yp+30, Font size
    Gui, Add, Text, xs yp+30, Display time (in sec.)
    Gui, Add, Text, xs yp+30, Maximum line length

    Gui, Add, DropDownList, xs+%columnBpos2% ys+0 section w205 gVerifyOsdOptions Sort Choose1 vFontName, %FontName%
    Gui, Add, ListView, xp+0 yp+30 w55 h25 %CCLVO% Background%OSDtextColor% vOSDtextColor hwndhLV1,
    Gui, Add, ListView, x+5 yp w55 h25 %CCLVO% Background%OSDbgrColor% vOSDbgrColor hwndhLV2,
    Gui, Add, Edit, x+5 yp+0 w55 hp geditsOSDwin r1 limit3 -multi number -wantCtrlA -wantReturn -wantTab -wrap veditF10, %OSDalpha%
    Gui, Add, UpDown, vOSDalpha gVerifyOsdOptions Range25-250, %OSDalpha%
    Gui, Add, Edit, xp-120 yp+30 w55 geditsOSDwin r1 limit3 -multi number -wantCtrlA -wantReturn -wantTab -wrap veditF5, %FontSize%
    Gui, Add, UpDown, gVerifyOsdOptions vFontSize Range12-295, %FontSize%
    Gui, Add, Edit, xp+0 yp+30 w55 hp geditsOSDwin r1 limit2 -multi number -wantCtrlA -wantReturn -wantTab -wrap veditF6, %DisplayTimeUser%
    Gui, Add, UpDown, vDisplayTimeUser gVerifyOsdOptions Range1-99, %DisplayTimeUser%
    Gui, Add, Edit, xp+0 yp+30 w55 hp geditsOSDwin r1 limit3 -multi number -wantCtrlA -wantReturn -wantTab -wrap veditF60, %maxMainLength%
    Gui, Add, UpDown, vmaxMainLength gVerifyOsdOptions Range10-130, %maxMainLength%

    If !FontList._NewEnum()[k, v]
    {
        Fnt_GetListOfFonts()
        FontList := trimArray(FontList)
    }

    Loop, % FontList.MaxIndex() {
        fontNameInstalled := FontList[A_Index]
        If (fontNameInstalled ~= "i)(@|oem|extb|symbol|marlett|wst_|glyph|reference specialty|system|terminal|mt extra|small fonts|cambria math|this font is not|fixedsys|emoji|hksc| mdl|wingdings|webdings)") || (fontNameInstalled=FontName)
           Continue
        GuiControl, , FontName, %fontNameInstalled%
    }

    Gui, Tab
    Gui, Font, Bold
    Gui, Add, Checkbox, y+8 gVerifyOsdOptions Checked%ShowPreview% vShowPreview, Show preview window
    Gui, Add, Edit, x+7 gVerifyOsdOptions w%editFieldWid% limit980 r1 -multi -wantReturn -wantTab -wrap vPreviewWindowText, %PreviewWindowText%
    Gui, Font, Normal

    Gui, Add, Button, xm+0 y+10 w70 h30 Default gApplySettings vApplySettingsBTN, A&pply
    Gui, Add, Button, x+8 wp hp gCloseSettings, C&ancel
    Gui, Add, Button, x+8 w160 hp gDeleteSettings, R&estore defaults
    Gui, Show, AutoSize, Customize: Text Capture ACC
    VerifyOsdOptions(0)
    ColorPickerHandles := hLV1 "," hLV2 "," hLV3 "," hLV5 "," hTXT
}

VerifyOsdOptions(EnableApply:=1) {
    GuiControlGet, ShowPreview
    GuiControlGet, GlobalKBDhotkeys

    GuiControl, % (EnableApply=0 ? "Disable" : "Enable"), ApplySettingsBTN
    GuiControl, % (ShowPreview=1 ? "Enable" : "Disable"), PreviewWindowText

    If (GlobalKBDhotkeys=0)
    {
       SwitchStateKBDbtn("KBDCapText", 0)
       SwitchStateKBDbtn("KBDCapTextConstant", 0)
    } Else
    {
       SwitchStateKBDbtn("KBDCapText", 1)
       SwitchStateKBDbtn("KBDCapTextConstant", 1)
    }

    ProcessComboKBD()

    Static LastInvoked := 1
    If (A_TickCount - LastInvoked>200) || (MainGuiVisible=0 && ShowPreview=1)
    || (MainGuiVisible=1 && ShowPreview=0)
    {
       LastInvoked := A_TickCount
       OSDpreview()
    }
}

trimArray(arr) { ; Hash O(n) 
; Function by errorseven from:
; https://stackoverflow.com/questions/46432447/how-do-i-remove-duplicates-from-an-autohotkey-array
    hash := {}, newArr := []
    For e, v in arr
        If (!hash.Haskey(v))
            hash[(v)] := 1, newArr.push(v)
    Return newArr
}

DonateNow() {
   Run, https://www.paypal.me/MariusSucan/15
   CloseWindow()
}

AboutWindow() {
    If (PrefOpen=1)
    {
        SoundBeep, 300, 900
        Return
    }

    If (AnyWindowOpen=1)
    {
       CloseWindow()
       Return
    }

    SettingsGUI()
    AnyWindowOpen := 1
    btnWid := 100
    txtWid := 360
    Global btn1
    Gui, Font, s20 Bold, Arial, -wrap
    Gui, Add, Text, x+7 y15, Text Capture ACC v%Version%
    Gui, Font
    If (PrefsLargeFonts=1)
    {
       btnWid := btnWid + 50
       txtWid := txtWid + 105
       Gui, Font, s%LargeUIfontValue%
    }
    Gui, Add, Link, y+4, Developed by <a href="http://marius.sucan.ro">Marius Şucan</a> on AHK_H v1.1.28.
    Gui, Add, Text, y+10 w%txtWid% Section, This application contains code from various entities. You can find more details in the source code.
    Gui, Font, Bold
    Gui, Add, Link, xp+25 y+10, To keep the development going, `n<a href="https://www.paypal.me/MariusSucan/15">please donate</a> or <a href="mailto:marius.sucan@gmail.com">send me feedback</a>.
    Gui, Font, Normal
    Gui, Add, Button, xs+0 y+20 w75 Default gCloseWindow, &Close
    Gui, Add, Button, x+5 w75 Default gShowOSDsettings, &Settings
    Gui, Add, Text, x+8 hp +0x200, Released: %ReleaseDate%
    Gui, Show, AutoSize, About
    ColorPickerHandles := hDonateBTN "," hIcon
    Sleep, 25
}

;================================================================
; Section 8. Other functions:
; - Updater, file existence checks.
; - Load, verify and save settings
;================================================================

INIaction(act, var, section) {
  varValue := %var%
  If (act=1)
     IniWrite, %varValue%, %IniFile%, %section%, %var%
  Else
     IniRead, %var%, %IniFile%, %section%, %var%, %varValue%
}

INIsettings(a) {
  FirstRun := 0
  If (a=1) ; a=1 means save into INI
  {
     INIaction(1, "FirstRun", "SavedSettings")
     INIaction(1, "ReleaseDate", "SavedSettings")
     INIaction(1, "Version", "SavedSettings")
  }
  INIaction(a, "PrefsLargeFonts", "SavedSettings")
  INIaction(a, "Copy2Clip", "SavedSettings")
  INIaction(a, "showACCdetails", "SavedSettings")

; OSD settings
  INIaction(a, "DisplayTimeUser", "OSDprefs")
  INIaction(a, "FontName", "OSDprefs")
  INIaction(a, "FontSize", "OSDprefs")
  INIaction(a, "GuiY", "OSDprefs")
  INIaction(a, "GuiX", "OSDprefs")
  INIaction(a, "OSDbgrColor", "OSDprefs")
  INIaction(a, "OSDtextColor", "OSDprefs")
  INIaction(a, "OSDalpha", "OSDprefs")
  INIaction(a, "OSDmarginTop", "OSDprefs")
  INIaction(a, "OSDmarginBottom", "OSDprefs")
  INIaction(a, "OSDmarginSides", "OSDprefs")
  INIaction(a, "maxMainLength", "OSDprefs")

; Hotkey settings
  INIaction(a, "GlobalKBDhotkeys", "Hotkeys")
  INIaction(a, "KBDCapText", "Hotkeys")

  If (a=0) ; a=0 means to load from INI
     CheckSettings()
}

BinaryVar(ByRef givenVar, defy) {
    givenVar := (Round(givenVar)=0 || Round(givenVar)=1) ? Round(givenVar) : defy
}

HexyVar(ByRef givenVar, defy) {
   If (givenVar ~= "[^[:xdigit:]]") || (StrLen(givenVar)!=6)
      givenVar := defy
}

MinMaxVar(ByRef givenVar, miny, maxy, defy) {
    testNumber := givenVar
    If (testNumber ~= "i)^(\-[\p{N}])")
       StringReplace, testNumber, testNumber, -

    If testNumber is not digit
    {
       givenVar := defy
       Return
    }

    givenVar := (Round(givenVar) < miny) ? miny : Round(givenVar)
    givenVar := (Round(givenVar) > maxy) ? maxy : Round(givenVar)
}

CheckSettings() {

; verify check boxes
    BinaryVar(Copy2Clip, 1)
    BinaryVar(showACCdetails, 1)
    BinaryVar(GlobalKBDhotkeys, 1)
    BinaryVar(PrefsLargeFonts, 0)

; correct contradictory settings

; verify numeric values: min, max and default values
    MinMaxVar(DisplayTimeUser, 1, 99, 3)
    MinMaxVar(FontSize, 10, 300, 20)
    MinMaxVar(GuiX, -9999, 9999, 40)
    MinMaxVar(GuiY, -9999, 9999, 250)
    MinMaxVar(OSDmarginTop, 1, 900, 20)
    MinMaxVar(OSDmarginBottom, 1, 900, 20)
    MinMaxVar(OSDmarginSides, 1, 900, 25)
    MinMaxVar(maxMainLength, 10, 130, 55)
    MinMaxVar(OSDalpha, 24, 252, 230)

; verify HEX values

   HexyVar(OSDbgrColor, "131209")
   HexyVar(OSDtextColor, "FFFEFA")

   FontName := (StrLen(FontName)>2) ? FontName
             : (A_OSVersion!="WIN_XP") ? "Arial"
             : FileExist(A_WinDir "\Fonts\ARIALUNI.TTF") ? "Arial Unicode MS" : "Arial"
}


;================================================================
; Section 9. Functions not written by Marius Sucan.
; Here, I placed only the functions I was unable to decide
; where to place within the code structure. Yet, they had 
; one thing in common: written by other people.
;
; Please note, some of the functions borrowed may or may not
; be modified/adapted/transformed by Marius Șucan or other people.
;================================================================

GetPhysicalCursorPos(ByRef mX, ByRef mY) {
; function from: https://github.com/jNizM/AHK_DllCall_WinAPI/blob/master/src/Cursor%20Functions/GetPhysicalCursorPos.ahk
; by jNizM, modified by Marius Șucan
    Static POINT, init := VarSetCapacity(POINT, 8, 0) && NumPut(8, POINT, "Int")
    If !(DllCall("user32.dll\GetPhysicalCursorPos", "Ptr", &POINT))
       Return MouseGetPos, mX, mY
;       Return DllCall("kernel32.dll\GetLastError")
    mX := NumGet(POINT, 0, "Int")
    mY := NumGet(POINT, 4, "Int")
    Return
}

;================================================================
; Functions by Drugwash. Direct contribuitor to this script. Many thanks!
; ===============================================================

Cleanup() {
    OnMessage(0x4a, "")
    OnMessage(0x200, "")
    OnMessage(0x102, "")
    OnMessage(0x103, "")
    DllCall("wtsapi32\WTSUnRegisterSessionNotification", "Ptr", hMain)
    func2exec := "ahkThread_Free"
    Sleep, 10
    a := "Acc_Init"
    If IsFunc(a)
       %a%(1)

    Gui, OSD: Destroy
    DllCall("kernel32\FreeLibrary", "Ptr", hWinMM)

    Fnt_DeleteFont(hFont)
}
; ------------------------------------------------------------- ; from Drugwash

;================================================================
; The following functions were extracted from Font Library 3.0 for AHK
; ===============================================================

Fnt_SetFont(hControl,hFont:="",p_Redraw:=False) {
    Static Dummy30050039
          ,DEFAULT_GUI_FONT:= 17
          ,OBJ_FONT        := 6
          ,WM_SETFONT      := 0x30

    ;-- If needed, get the handle to the default GUI font
    If (DllCall("gdi32\GetObjectType","Ptr",hFont)<>OBJ_FONT)
        hFont:=DllCall("gdi32\GetStockObject","Int",DEFAULT_GUI_FONT)

    ;-- Set font
    SendMessage WM_SETFONT,hFont,p_Redraw,,ahk_id %hControl%
}

Fnt_CreateFont(p_Name:="",p_Options:="") {
    Static Dummy34361446

          ;-- Misc. font constants
          ,LOGPIXELSY:=90
          ,CLIP_DEFAULT_PRECIS:=0
          ,DEFAULT_CHARSET    :=1
          ,DEFAULT_GUI_FONT   :=17
          ,OUT_TT_PRECIS      :=4

          ;-- Font family
          ,FF_DONTCARE  :=0x0
          ,FF_ROMAN     :=0x1
          ,FF_SWISS     :=0x2
          ,FF_MODERN    :=0x3
          ,FF_SCRIPT    :=0x4
          ,FF_DECORATIVE:=0x5

          ;-- Font pitch
          ,DEFAULT_PITCH :=0
          ,FIXED_PITCH   :=1
          ,VARIABLE_PITCH:=2

          ;-- Font quality
          ,DEFAULT_QUALITY       :=0
          ,DRAFT_QUALITY         :=1
          ,PROOF_QUALITY         :=2  ;-- AutoHotkey default
          ,NONANTIALIASED_QUALITY:=3
          ,ANTIALIASED_QUALITY   :=4
          ,CLEARTYPE_QUALITY     :=5

          ;-- Font weight
          ,FW_DONTCARE:=0
          ,FW_NORMAL  :=400
          ,FW_BOLD    :=700

    ;-- Parameters
    ;   Remove all leading/trailing white space
    p_Name   :=Trim(p_Name," `f`n`r`t`v")
    p_Options:=Trim(p_Options," `f`n`r`t`v")

    ;-- If both parameters are null or unspecified, return the handle to the
    ;   default GUI font.
    If (p_Name="" and p_Options="")
        Return DllCall("gdi32\GetStockObject","Int",DEFAULT_GUI_FONT)

    ;-- Initialize options
    o_Height   :=""             ;-- Undefined
    o_Italic   :=False
    o_Quality  :=PROOF_QUALITY  ;-- AutoHotkey default
    o_Size     :=""             ;-- Undefined
    o_Strikeout:=False
    o_Underline:=False
    o_Weight   :=FW_DONTCARE

    ;-- Extract options (if any) from p_Options
    Loop Parse,p_Options,%A_Space%
        {
        If A_LoopField is Space
            Continue

        If (SubStr(A_LoopField,1,4)="bold")
            o_Weight:=FW_BOLD
        Else If (SubStr(A_LoopField,1,6)="italic")
            o_Italic:=True
        Else If (SubStr(A_LoopField,1,4)="norm")
            {
            o_Italic   :=False
            o_Strikeout:=False
            o_Underline:=False
            o_Weight   :=FW_DONTCARE
            }
        Else If (A_LoopField="-s")
            o_Size:=0
        Else If (SubStr(A_LoopField,1,6)="strike")
            o_Strikeout:=True
        Else If (SubStr(A_LoopField,1,9)="underline")
            o_Underline:=True
        Else If (SubStr(A_LoopField,1,1)="h")
            {
            o_Height:=SubStr(A_LoopField,2)
            o_Size  :=""  ;-- Undefined
            }
        Else If (SubStr(A_LoopField,1,1)="q")
            o_Quality:=SubStr(A_LoopField,2)
        Else If (SubStr(A_LoopField,1,1)="s")
            {
            o_Size  :=SubStr(A_LoopField,2)
            o_Height:=""  ;-- Undefined
            }
        Else If (SubStr(A_LoopField,1,1)="w")
            o_Weight:=SubStr(A_LoopField,2)
        }

    ;-- Convert/Fix invalid or
    ;-- unspecified parameters/options
    If p_Name is Space
        p_Name:=Fnt_GetFontName()   ;-- Font name of the default GUI font

    If o_Height is not Integer
        o_Height:=""                ;-- Undefined

    If o_Quality is not Integer
        o_Quality:=PROOF_QUALITY    ;-- AutoHotkey default

    If o_Size is Space              ;-- Undefined
        o_Size:=Fnt_GetFontSize()   ;-- Font size of the default GUI font
     Else
        If o_Size is not Integer
            o_Size:=""              ;-- Undefined
         Else
            If (o_Size=0)
                o_Size:=""          ;-- Undefined

    If o_Weight is not Integer
        o_Weight:=FW_DONTCARE       ;-- A font with a default weight is created

    ;-- If needed, convert point size to em height
    If o_Height is Space        ;-- Undefined
        If o_Size is Integer    ;-- Allows for a negative size (emulates AutoHotkey)
            {
            hDC:=DllCall("gdi32\CreateDCW","Str","DISPLAY","Ptr",0,"Ptr",0,"Ptr",0)
            o_Height:=-Round(o_Size*DllCall("gdi32\GetDeviceCaps","Ptr",hDC,"Int",LOGPIXELSY)/72)
            DllCall("gdi32\DeleteDC","Ptr",hDC)
            }

    If o_Height is not Integer
        o_Height:=0                 ;-- A font with a default height is created

    ;-- Create font
    hFont:=DllCall("gdi32\CreateFontW"
        ,"Int",o_Height                                 ;-- nHeight
        ,"Int",0                                        ;-- nWidth
        ,"Int",0                                        ;-- nEscapement (0=normal horizontal)
        ,"Int",0                                        ;-- nOrientation
        ,"Int",o_Weight                                 ;-- fnWeight
        ,"UInt",o_Italic                                ;-- fdwItalic
        ,"UInt",o_Underline                             ;-- fdwUnderline
        ,"UInt",o_Strikeout                             ;-- fdwStrikeOut
        ,"UInt",DEFAULT_CHARSET                         ;-- fdwCharSet
        ,"UInt",OUT_TT_PRECIS                           ;-- fdwOutputPrecision
        ,"UInt",CLIP_DEFAULT_PRECIS                     ;-- fdwClipPrecision
        ,"UInt",o_Quality                               ;-- fdwQuality
        ,"UInt",(FF_DONTCARE<<4)|DEFAULT_PITCH          ;-- fdwPitchAndFamily
        ,"Str",SubStr(p_Name,1,31))                     ;-- lpszFace

    Return hFont
}

Fnt_DeleteFont(hFont) {
    If not hFont  ;-- Zero or null
        Return True

    Return DllCall("gdi32\DeleteObject","Ptr",hFont) ? True:False
}

Fnt_GetFontName(hFont:="") {
    Static Dummy87890484
          ,DEFAULT_GUI_FONT    :=17
          ,HWND_DESKTOP        :=0
          ,OBJ_FONT            :=6
          ,MAX_FONT_NAME_LENGTH:=32     ;-- In TCHARS

    ;-- If needed, get the handle to the default GUI font
    If (DllCall("gdi32\GetObjectType","Ptr",hFont)<>OBJ_FONT)
        hFont:=DllCall("gdi32\GetStockObject","Int",DEFAULT_GUI_FONT)

    ;-- Select the font into the device context for the desktop
    hDC      :=DllCall("user32\GetDC","Ptr",HWND_DESKTOP)
    old_hFont:=DllCall("gdi32\SelectObject","Ptr",hDC,"Ptr",hFont)

    ;-- Get the font name
    VarSetCapacity(l_FontName,MAX_FONT_NAME_LENGTH*(A_IsUnicode ? 2:1))
    DllCall("gdi32\GetTextFaceW","Ptr",hDC,"Int",MAX_FONT_NAME_LENGTH,"Str",l_FontName)

    ;-- Release the objects needed by the GetTextFace function
    DllCall("gdi32\SelectObject","Ptr",hDC,"Ptr",old_hFont)
        ;-- Necessary to avoid memory leak

    DllCall("user32\ReleaseDC","Ptr",HWND_DESKTOP,"Ptr",hDC)
    Return l_FontName
}

Fnt_GetFontSize(hFont:="") {
    Static Dummy64998752

          ;-- Device constants
          ,HWND_DESKTOP:=0
          ,LOGPIXELSY  :=90

          ;-- Misc.
          ,DEFAULT_GUI_FONT:=17
          ,OBJ_FONT        :=6

    ;-- If needed, get the handle to the default GUI font
    If (DllCall("gdi32\GetObjectType","Ptr",hFont)<>OBJ_FONT)
        hFont:=DllCall("gdi32\GetStockObject","Int",DEFAULT_GUI_FONT)

    ;-- Select the font into the device context for the desktop
    hDC      :=DllCall("user32\GetDC","Ptr",HWND_DESKTOP)
    old_hFont:=DllCall("gdi32\SelectObject","Ptr",hDC,"Ptr",hFont)

    ;-- Collect the number of pixels per logical inch along the screen height
    l_LogPixelsY:=DllCall("gdi32\GetDeviceCaps","Ptr",hDC,"Int",LOGPIXELSY)

    ;-- Get text metrics for the font
    VarSetCapacity(TEXTMETRIC,A_IsUnicode ? 60:56,0)
    DllCall("gdi32\GetTextMetricsW","Ptr",hDC,"Ptr",&TEXTMETRIC)

    ;-- Convert em height to point size
    l_Size:=Round((NumGet(TEXTMETRIC,0,"Int")-NumGet(TEXTMETRIC,12,"Int"))*72/l_LogPixelsY)
        ;-- (Height - Internal Leading) * 72 / LogPixelsY

    ;-- Release the objects needed by the GetTextMetrics function
    DllCall("gdi32\SelectObject","Ptr",hDC,"Ptr",old_hFont)
        ;-- Necessary to avoid memory leak

    DllCall("user32\ReleaseDC","Ptr",HWND_DESKTOP,"Ptr",hDC)
    Return l_Size
}

Fnt_GetListOfFonts() {
; function stripped down from Font Library 3.0 by jballi
; from https://autohotkey.com/boards/viewtopic.php?t=4379

    Static Dummy65612414
          ,HWND_DESKTOP := 0  ;-- Device constants
          ,LF_FACESIZE := 32  ;-- In TCHARS - LOGFONT constants

    ;-- Initialize and populate LOGFONT structure
    Fnt_EnumFontFamExProc_List := ""
    p_CharSet := 1
    p_Flags := 0x800
    VarSetCapacity(LOGFONT,A_IsUnicode ? 92:60,0)
    NumPut(p_CharSet,LOGFONT,23,"UChar")                ;-- lfCharSet

    ;-- Enumerate fonts
    EFFEP := RegisterCallback("Fnt_EnumFontFamExProc","F")
    hDC := DllCall("user32\GetDC","Ptr",HWND_DESKTOP)
    DllCall("gdi32\EnumFontFamiliesExW"
        ,"Ptr", hDC                                      ;-- hdc
        ,"Ptr", &LOGFONT                                 ;-- lpLogfont
        ,"Ptr", EFFEP                                    ;-- lpEnumFontFamExProc
        ,"Ptr", p_Flags                                  ;-- lParam
        ,"UInt", 0)                                      ;-- dwFlags (must be 0)

    DllCall("user32\ReleaseDC","Ptr",HWND_DESKTOP,"Ptr",hDC)
    DllCall("GlobalFree", "Ptr", EFFEP)
    Return Fnt_EnumFontFamExProc_List
}

Fnt_EnumFontFamExProc(lpelfe,lpntme,FontType,p_Flags) {
    Fnt_EnumFontFamExProc_List := 0
    Static Dummy62479817
           ,LF_FACESIZE := 32  ;-- In TCHARS - LOGFONT constants

    l_FaceName := StrGet(lpelfe+28,LF_FACESIZE)
    FontList.Push(l_FaceName)    ;-- Append the font name to the list
    Return True  ;-- Continue enumeration
}
; ------------------------------------------------------------- ; Font Library

dummy() {
    Return
}



; ================================================================

; Accessible Info Viewer
; by Sean and jethrow
; http://www.autohotkey.com/board/topic/77888-accessible-info-viewer-alpha-release-2012-09-20/
; https://dl.dropbox.com/u/47573473/Accessible%20Info%20Viewer/AccViewer%20Source.ahk
; Modified in 2018 by Marius Șucan for KeyPress OSD. 
;

GetAccInfo(skipVerification:=0) {
  DetectHiddenWindows, On
  SendMessage, WM_GETOBJECT := 0x003D, 0, 1, Chrome_RenderWidgetHostHWND1, A
  Acc := Acc_ObjectFromPoint(ChildId)
  ; Acc := Acc_ObjectFromWindow(ChildId)
  UpdateAccInfo(Acc, ChildId)
}

UpdateAccInfo(Acc, ChildId, Obj_Path="") {
  Global InputMsg, AccViewName, AccViewValue, CtrlTextVar, NewCtrlTextVar
  Global uia := UIA_Interface()
  Global Element := uia.ElementFromPoint()
  
  MouseGetPos, , , WinID, hChild, 3
  If (WinID=hMainOSD)
  {
     ; DestroyMainGui()
     Return
  }

  MouseGetPos, , , id, controla, 2
  ControlGetText, NewCtrlTextVar , , ahk_id %controla%
  CtrlTextVar := StrLen(NewCtrlTextVar)>1 || !MainGuiVisible ? NewCtrlTextVar : CtrlTextVar

  NewAccViewName := Element.CurrentName
  If !NewAccViewName
     NewAccViewName := Acc.accName(ChildId)
  AccViewName := StrLen(NewAccViewName)>1 || !MainGuiVisible ? NewAccViewName : AccViewName

  For each, value in [30093,30092,30045] ; lvalue,lname,value
      NewAccViewValue := Element.GetCurrentPropertyValue(value)
  Until r != ""
  If !NewAccViewValue
     NewAccViewValue := Acc.accValue(ChildId)
  AccViewValue := StrLen(NewAccViewValue)>1 || !MainGuiVisible ? NewAccViewValue : AccViewValue
  CtrlTextVar := RegExReplace(CtrlTextVar, "i)^(\s+)")
  AccViewName := RegExReplace(AccViewName, "i)^(\s+)")
  AccViewValue := RegExReplace(AccViewValue, "i)^(\s+)")
  If (StrLen(AccViewName) = StrLen(CtrlTextVar)-1) || (StrLen(AccViewName) = StrLen(CtrlTextVar)+1)
     CtrlTextVar := ""
  If (AccViewName=AccViewValue)
     AccViewValue := ""
  If (AccViewName=CtrlTextVar) || (AccViewValue=CtrlTextVar)
     CtrlTextVar := ""

  otherDetails := Acc_GetRoleText(Acc.accRole(ChildId)) " " Acc_GetStateText(Acc.accState(ChildId)) " " Acc.accDefaultAction(ChildId) " " Acc.accDescription(ChildId) " " Acc.accHelp(ChildId)
  If (showACCdetails=1)
     NewInputMsg := AccViewName " " AccViewValue " " CtrlTextVar " " otherDetails
  Else
     NewInputMsg := AccViewName " " AccViewValue
  StringReplace, NewInputMsg, NewInputMsg, %A_TAB%, %A_SPACE%, All
  StringReplace, NewInputMsg, NewInputMsg, %A_SPACE%%A_SPACE%, %A_SPACE%, All
  ; NewInputMsg .= GetMenu(WinID)

  If (NewInputMsg!=InputMsg || AccTextCaptureActive=0)
  {
     CreateMainGUI(NewInputMsg)
     If (AccTextCaptureActive=0 && Copy2Clip=1)
        Clipboard := NewInputMsg
     InputMsg := NewInputMsg
  }
}

GetClassNN(Chwnd, Whwnd) {
  Global _GetClassNN := {}
  _GetClassNN.Hwnd := Chwnd
  Detect := A_DetectHiddenWindows
  WinGetClass, Class, ahk_id %Chwnd%
  _GetClassNN.Class := Class
  DetectHiddenWindows, On
  EnumAddress := RegisterCallback("GetClassNN_EnumChildProc")
  DllCall("user32\EnumChildWindows", "UInt",Whwnd, "UInt",EnumAddress)
  DetectHiddenWindows, %Detect%
  Return, _GetClassNN.ClassNN, _GetClassNN:=""
}

GetClassNN_EnumChildProc(hwnd, lparam) {
  Static Occurrence
  Global _GetClassNN
  WinGetClass, Class, ahk_id %hwnd%
  If _GetClassNN.Class == Class
    Occurrence++
  If Not _GetClassNN.Hwnd == hwnd
    Return true
  Else {
    _GetClassNN.ClassNN := _GetClassNN.Class Occurrence
    Occurrence := 0
    Return false
  }
}

TV_Expanded(TVid) {
  For Each, TV_Child_ID in TVobj[TVid].Children
    If TVobj[TV_Child_ID].need_children
      TV_BuildAccChildren(TVobj[TV_Child_ID].obj, TV_Child_ID)
}

TV_BuildAccChildren(AccObj, Parent, Selected_Child="", Flag="") {
  TVobj[Parent].need_children := false
  Parent_Obj_Path := Trim(TVobj[Parent].Obj_Path, ",")
  For wach, child in Acc_Children(AccObj) {
    If Not IsObject(child) {
      added := TV_Add("[" A_Index "] " Acc_GetRoleText(AccObj.accRole(child)), Parent)
      TVobj[added] := {is_obj:false, obj:Acc, childid:child, Obj_Path:Parent_Obj_Path}
      If (child = Selected_Child)
        TV_Modify(added, "Select")
    }
    Else {
      added := TV_Add("[" A_Index "] " Acc_Role(child), Parent, "bold")
      TVobj[added] := {is_obj:true, need_children:true, obj:child, childid:0, Children:[], Obj_Path:Trim(Parent_Obj_Path "," A_Index, ",")}
    }
    TVobj[Parent].Children.Insert(added)
    If (A_Index = Flag)
      Flagged_Child := added
  }
  Return Flagged_Child
}

GetAccPath(Acc, byref hwnd="") {
  hwnd := Acc_WindowFromObject(Acc)
  WinObj := Acc_ObjectFromWindow(hwnd)
  WinObjPos := Acc_Location(WinObj).pos
  While Acc_WindowFromObject(Parent:=Acc_Parent(Acc)) = hwnd {
    t2 := GetEnumIndex(Acc) "." t2
    If Acc_Location(Parent).pos = WinObjPos
      Return {AccObj:Parent, Path:SubStr(t2,1,-1)}
    Acc := Parent
  }
  While Acc_WindowFromObject(Parent:=Acc_Parent(WinObj)) = hwnd
    t1.="P.", WinObj:=Parent
  Return {AccObj:Acc, Path:t1 SubStr(t2,1,-1)}
}

GetEnumIndex(Acc, ChildId=0) {
  If Not ChildId {
    ChildPos := Acc_Location(Acc).pos
    For Each, child in Acc_Children(Acc_Parent(Acc))
      If IsObject(child) and Acc_Location(child).pos=ChildPos
        Return A_Index
  } 
  Else {
    ChildPos := Acc_Location(Acc,ChildId).pos
    For Each, child in Acc_Children(Acc)
      If Not IsObject(child) and Acc_Location(Acc,child).pos=ChildPos
        Return A_Index
  }
}

GetAccLocation(AccObj, Child=0, byref x="", byref y="", byref w="", byref h="") {
  AccObj.accLocation(ComObj(0x4003,&x:=0), ComObj(0x4003,&y:=0), ComObj(0x4003,&w:=0), ComObj(0x4003,&h:=0), Child)
  Return  "x" (x:=NumGet(x,0,"Int")) "  "
  .  "y" (y:=NumGet(y,0,"Int")) "  "
  .  "w" (w:=NumGet(w,0,"Int")) "  "
  .  "h" (h:=NumGet(h,0,"Int"))
}

;================================================================
; Acc Library
;================================================================
  Acc_Init(unload:=0) {
    Static h := 0
    If !h
      h:=DllCall("kernel32\LoadLibraryW","Str","oleacc","Ptr")
   If (h && unload)
     Dllcall("kernel32\FreeLibrary", "Ptr", h)
  }

  Acc_ObjectFromEvent(ByRef _idChild_, hWnd, idObject, idChild) {
    Acc_Init()
    If DllCall("oleacc\AccessibleObjectFromEvent", "Ptr", hWnd, "UInt", idObject, "UInt", idChild, "Ptr*", pacc, "Ptr", VarSetCapacity(varChild,8+2*A_PtrSize,0)*0+&varChild)=0
       Return  ComObjEnwrap(9,pacc,1), _idChild_:=NumGet(varChild,8,"UInt")
  }

  Acc_ObjectFromPoint(ByRef _idChild_ = "", x = "", y = "") {
    Acc_Init()
    If  DllCall("oleacc\AccessibleObjectFromPoint", "Int64", x==""||y==""?0*DllCall("user32\GetCursorPos","Int64*",pt)+pt:x&0xFFFFFFFF|y<<32, "Ptr*", pacc, "Ptr", VarSetCapacity(varChild,8+2*A_PtrSize,0)*0+&varChild)=0
    Return  ComObjEnwrap(9,pacc,1), _idChild_:=NumGet(varChild,8,"UInt")
  }

  Acc_ObjectFromWindow(hWnd, idObject = 0) {
    Acc_Init()
    If  DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
    Return  ComObjEnwrap(9,pacc,1)
  }

  Acc_WindowFromObject(pacc) {
    If DllCall("oleacc\WindowFromAccessibleObject", "Ptr", IsObject(pacc)?ComObjValue(pacc):pacc, "Ptr*", hWnd)=0
    Return  hWnd
  }

  Acc_GetRoleText(nRole) {
    nSize := DllCall("oleacc\GetRoleText", "Uint", nRole, "Ptr", 0, "Uint", 0)
    VarSetCapacity(sRole, (A_IsUnicode?2:1)*nSize)
    DllCall("oleacc\GetRoleText", "Uint", nRole, "str", sRole, "Uint", nSize+1)
    Return sRole
  }

  Acc_GetStateText(nState) {
    nSize := DllCall("oleacc\GetStateText", "Uint", nState, "Ptr", 0, "Uint", 0)
    VarSetCapacity(sState, (A_IsUnicode?2:1)*nSize)
    DllCall("oleacc\GetStateText", "Uint", nState, "str", sState, "Uint", nSize+1)
    Return sState
  }

  Acc_Role(Acc, ChildId=0) {
    try Return ComObjType(Acc,"Name")="IAccessible"?Acc_GetRoleText(Acc.accRole(ChildId)):"invalid object"
  }

  Acc_State(Acc, ChildId=0) {
    try Return ComObjType(Acc,"Name")="IAccessible"?Acc_GetStateText(Acc.accState(ChildId)):"invalid object"
  }

  Acc_Children(Acc) {
    If ComObjType(Acc,"Name")!="IAccessible"
      error_message := "Cause:`tInvalid IAccessible Object`n`n"
    Else
    {
      Acc_Init()
      cChildren:=Acc.accChildCount, Children:=[]
      If DllCall("oleacc\AccessibleChildren", "Ptr", ComObjValue(Acc), "Int", 0, "Int", cChildren, "Ptr", VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*", cChildren)=0
      {
        Loop %cChildren%
          i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i), Children.Insert(NumGet(varChildren,i-8)=3?child:Acc_Query(child)), ObjRelease(child)
      Return Children
      }
    }
    error:=Exception("",-1)
    MsgBox, 262148, Acc_Children Failed, % (error_message?error_message:"") "File:`t" (error.file==A_ScriptFullPath?A_ScriptName:error.file) "`nLine:`t" error.line "`n`nContinue Script?"
    IfMsgBox, No
      ExitApp
  }

  Acc_Location(Acc, ChildId=0) {
    try Acc.accLocation(ComObj(0x4003,&x:=0), ComObj(0x4003,&y:=0), ComObj(0x4003,&w:=0), ComObj(0x4003,&h:=0), ChildId)
    catch
    Return
    Return  {x:NumGet(x,0,"Int"), y:NumGet(y,0,"Int"), w:NumGet(w,0,"Int"), h:NumGet(h,0,"Int")
    ,  pos:"x" NumGet(x,0,"Int")" y" NumGet(y,0,"Int") " w" NumGet(w,0,"Int") " h" NumGet(h,0,"Int")}
  }

  Acc_Parent(Acc) {
    try parent:=Acc.accParent
    Return parent?Acc_Query(parent):
  }

  Acc_Child(Acc, ChildId=0) {
    try child:=Acc.accChild(ChildId)
    Return child?Acc_Query(child):
  }

  Acc_Query(Acc) {
    try Return ComObj(9, ComObjQuery(Acc,"{618736e0-3c3d-11cf-810c-00aa00389b71}"), 1)
  }
;================================================================
; Anchor
;================================================================

Anchor(i, a = "", r = false) {
  Static c, cs := 12, cx := 255, cl := 0, g, gs := 8, gl := 0, gpi, gw, gh, z := 0, k := 0xffff
  If z = 0
    VarSetCapacity(g, gs * 99, 0), VarSetCapacity(c, cs * cx, 0), z := true
  If (!WinExist("ahk_id" . i))
  {
    GuiControlGet, t, Hwnd, %i%
    If ErrorLevel = 0
    i := t
    Else ControlGet, i, Hwnd, , %i%
  }
  VarSetCapacity(gi, 68, 0), DllCall("user32\GetWindowInfo", "UInt", gp := DllCall("user32\GetParent", "UInt", i), "Ptr", &gi)
  , giw := NumGet(gi, 28, "Int") - NumGet(gi, 20, "Int"), gih := NumGet(gi, 32, "Int") - NumGet(gi, 24, "Int")
  If (gp != gpi)
  {
    gpi := gp
    Loop, %gl%
      If (NumGet(g, cb := gs * (A_Index - 1)) == gp, "UInt")
      {
        gw := NumGet(g, cb + 4, "Short"), gh := NumGet(g, cb + 6, "Short"), gf := 1
        Break
      }
    If (!gf)
      NumPut(gp, g, gl, "UInt"), NumPut(gw := giw, g, gl + 4, "Short"), NumPut(gh := gih, g, gl + 6, "Short"), gl += gs
  }
  ControlGetPos, dx, dy, dw, dh, , ahk_id %i%
  Loop, %cl%
  If (NumGet(c, cb := cs * (A_Index - 1), "UInt") == i)
  {
    If a =
    {
      cf = 1
      Break
    }
    giw -= gw, gih -= gh, as := 1, dx := NumGet(c, cb + 4, "Short"), dy := NumGet(c, cb + 6, "Short")
    , cw := dw, dw := NumGet(c, cb + 8, "Short"), ch := dh, dh := NumGet(c, cb + 10, "Short")
    Loop, Parse, a, xywh
      If A_Index > 1
        av := SubStr(a, as, 1), as += 1 + StrLen(A_LoopField)
        , d%av% += (InStr("yh", av) ? gih : giw) * (A_LoopField + 0 ? A_LoopField : 1)
    DllCall("user32\SetWindowPos", "UInt", i, "UInt", 0, "Int", dx, "Int", dy
    , "Int", InStr(a, "w") ? dw : cw, "Int", InStr(a, "h") ? dh : ch, "Int", 4)
    If r != 0
      DllCall("user32\RedrawWindow", "UInt", i, "UInt", 0, "UInt", 0, "UInt", 0x0101)
    Return
  }
  If cf != 1
    cb := cl, cl += cs
  bx := NumGet(gi, 48, "UInt"), by := NumGet(gi, 16, "Int") - NumGet(gi, 8, "Int") - gih - NumGet(gi, 52, "UInt")
  If cf = 1
    dw -= giw - gw, dh -= gih - gh
  NumPut(i, c, cb, "UInt"), NumPut(dx - bx, c, cb + 4, "Short"), NumPut(dy - by, c, cb + 6, "Short")
  , NumPut(dw, c, cb + 8, "Short"), NumPut(dh, c, cb + 10, "Short")
  Return, true
}

WinGetAll(Which="Title", DetectHidden="Off"){
O_DHW := A_DetectHiddenWindows, O_BL := A_BatchLines ;Save original states
DetectHiddenWindows, % (DetectHidden != "off" && DetectHidden) ? "on" : "off"
SetBatchLines, -1
    WinGet, all, list ;get all hwnd
    If (Which="Title") ;return Window Titles
    {
        Loop, %all%
        {
            WinGetTitle, WTitle, % "ahk_id " all%A_Index%
            If WTitle ;Prevent to get blank titles
                Output .= WTitle "`n"        
        }
    }
    Else If (Which="Process") ;return Process Names
    {
        Loop, %all%
        {
            WinGet, PName, ProcessName, % "ahk_id " all%A_Index%
            Output .= PName "`n"
        }
    }
    Else If (Which="Class") ;return Window Classes
    {
        Loop, %all%
        {
            WinGetClass, WClass, % "ahk_id " all%A_Index%
            Output .= WClass "`n"
        }
    }
    Else If (Which="hwnd") ;return Window Handles (Unique ID)
    {
        Loop, %all%
            Output .= all%A_Index% "`n"
    }
    Else If (Which="PID") ;return Process Identifiers
    {
        Loop, %all%
        {
            WinGet, PID, PID, % "ahk_id " all%A_Index%
            Output .= PID "`n"        
        }
        Sort, Output, U N ;numeric order and remove duplicates
    }
DetectHiddenWindows, %O_DHW% ;back to original state
SetBatchLines, %O_BL% ;back to original state
    Sort, Output, U ;remove duplicates
    Return Output
}







; =====================================================

; from https://github.com/neptercn/Component_AHK/blob/master/uia.ahk
;~ UI Automation Constants: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671207(v=vs.85).aspx
;~ UI Automation Enumerations: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671210(v=vs.85).aspx
;~ http://www.autohotkey.com/board/topic/94619-ahk-l-screen-reader-a-tool-to-get-text-anywhere/
; by Sean and jethrow

/* Questions:
  - better way to do __properties?
  - support for Constants?
  - if method returns a SafeArray, should we return a Wrapped SafeArray, Raw SafeArray, or AHK Array
  - on UIA Interface conversion methods, how should the data be returned? wrapped/extracted or raw? should raw data be a ByRef param?
  - do variants need cleared? what about SysAllocString BSTRs?
  - do RECT struts need destroyed?
  - if returning wrapped data & raw is ByRef, will the wrapped data being released destroy the raw data?
  - returning varaint data other than vt=3|8|9|13|0x2000
  - Cached Members?
  - UIA Element existance - dependent on window being visible (non minimized)?
  - function(params, ByRef out="……")
*/


class UIA_Base {
  __New(p="", flag=1) {
    ObjInsert(this,"__Type","IUIAutomation" SubStr(this.__Class,5))
    ,ObjInsert(this,"__Value",p)
    ,ObjInsert(this,"__Flag",flag)
  }
  __Get(member) {
    if member not in base,__UIA ; base & __UIA should act as normal
    {  if raw:=SubStr(member,0)="*" ; return raw data - user should know what they are doing
        member:=SubStr(member,1,-1)
      if RegExMatch(this.__properties, "im)^" member ",(\d+),(\w+)", m) { ; if the member is in the properties. if not - give error message
        if (m2="VARIANT")  ; return VARIANT data - DllCall output param different
          return UIA_Hr(DllCall(this.__Vt(m1), "ptr",this.__Value, "ptr",UIA_Variant(out)))? (raw?out:UIA_VariantData(out)):
        else if (m2="RECT") ; return RECT struct - DllCall output param different
          return UIA_Hr(DllCall(this.__Vt(m1), "ptr",this.__Value, "ptr",&(rect,VarSetCapacity(rect,16))))? (raw?out:UIA_RectToObject(rect)):
        else if UIA_Hr(DllCall(this.__Vt(m1), "ptr",this.__Value, "ptr*",out))
          return raw?out:m2="BSTR"?StrGet(out):RegExMatch(m2,"i)IUIAutomation\K\w+",n)?new UIA_%n%(out):out ; Bool, int, DWORD, HWND, CONTROLTYPEID, OrientationType?
      }
      else throw Exception("Property not supported by the " this.__Class " Class.",-1,member)
    }
  }
  __Set(member) {
    throw Exception("Assigning values not supported by the " this.__Class " Class.",-1,member)
  }
  __Call(member) {
    if !ObjHasKey(UIA_Base,member)&&!ObjHasKey(this,member)
      throw Exception("Method Call not supported by the " this.__Class " Class.",-1,member)
  }
  __Delete() {
    this.__Flag? ObjRelease(this.__Value):
  }
  __Vt(n) {
    return NumGet(NumGet(this.__Value+0,"ptr")+n*A_PtrSize,"ptr")
  }
}  

class UIA_Interface extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671406(v=vs.85).aspx
  static __IID := "{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}"
    ,  __properties := "ControlViewWalker,14,IUIAutomationTreeWalker`r`nContentViewWalker,15,IUIAutomationTreeWalker`r`nRawViewWalker,16,IUIAutomationTreeWalker`r`nRawViewCondition,17,IUIAutomationCondition`r`nControlViewCondition,18,IUIAutomationCondition`r`nContentViewCondition,19,IUIAutomationCondition`r`nProxyFactoryMapping,48,IUIAutomationProxyFactoryMapping`r`nReservedNotSupportedValue,54,IUnknown`r`nReservedMixedAttributeValue,55,IUnknown"
  
  CompareElements(e1,e2) {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",e1.__Value, "ptr",e2.__Value, "int*",out))? out:
  }
  CompareRuntimeIds(r1,r2) {
    return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr",ComObjValue(r1), "ptr",ComObjValue(r2), "int*",out))? out:
  }
  GetRootElement() {
    return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr*",out))? new UIA_Element(out):
  }
  ElementFromHandle(hwnd) {
    return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr",hwnd, "ptr*",out))? new UIA_Element(out):
  }
  ElementFromPoint(x="", y="") {
    try {
        return UIA_Hr(DllCall(this.__Vt(7), "ptr",this.__Value, "int64",x==""||y==""?0*DllCall("user32\GetCursorPos","Int64*",pt)+pt:x&0xFFFFFFFF|y<<32, "ptr*",out))? new UIA_Element(out):
    }
  }  
  GetFocusedElement() {
    return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value, "ptr*",out))? new UIA_Element(out):
  }
  ;~ GetRootElementBuildCache   9
  ;~ ElementFromHandleBuildCache   10
  ;~ ElementFromPointBuildCache   11
  ;~ GetFocusedElementBuildCache   12
  CreateTreeWalker(condition) {
    return UIA_Hr(DllCall(this.__Vt(13), "ptr",this.__Value, "ptr",Condition.__Value, "ptr*",out))? new UIA_TreeWalker(out):
  }
  ;~ CreateCacheRequest   20

  CreateTrueCondition() {
    return UIA_Hr(DllCall(this.__Vt(21), "ptr",this.__Value, "ptr*",out))? new UIA_Condition(out):
  }
  CreateFalseCondition() {
    return UIA_Hr(DllCall(this.__Vt(22), "ptr",this.__Value, "ptr*",out))? new UIA_Condition(out):
  }
  CreatePropertyCondition(propertyId, ByRef var, type="Variant") {
    if (type!="Variant")
      UIA_Variant(var,type,var)
    return UIA_Hr(DllCall(this.__Vt(23), "ptr",this.__Value, "int",propertyId, "ptr",&var, "ptr*",out))? new UIA_PropertyCondition(out):
  }
  CreatePropertyConditionEx(propertyId, ByRef var, type="Variant", flags=0x1) { ; NOT TESTED
  ; PropertyConditionFlags_IgnoreCase = 0x1
    if (type!="Variant")
      UIA_Variant(var,type,var)
    return UIA_Hr(DllCall(this.__Vt(24), "ptr",this.__Value, "int",propertyId, "ptr",&var, "uint",flags, "ptr*",out))? new UIA_PropertyCondition(out):
  }
  CreateAndCondition(c1,c2) {
    return UIA_Hr(DllCall(this.__Vt(25), "ptr",this.__Value, "ptr",c1.__Value, "ptr",c2.__Value, "ptr*",out))? new UIA_AndCondition(out):
  }
  CreateAndConditionFromArray(array) { ; ComObj(0x2003)??
  ;->in: AHK Array or Wrapped SafeArray
    if ComObjValue(array)&0x2000
      SafeArray:=array
    else {
      SafeArray:=ComObj(0x2003,DllCall("oleaut32\SafeArrayCreateVector", "uint",13, "uint",0, "uint",array.MaxIndex()),1)
      for i,c in array
        SafeArray[A_Index-1]:=c.__Value, ObjAddRef(c.__Value) ; AddRef - SafeArrayDestroy will release UIA_Conditions - they also release themselves
    }
    return UIA_Hr(DllCall(this.__Vt(26), "ptr",this.__Value, "ptr",ComObjValue(SafeArray), "ptr*",out))? new UIA_AndCondition(out):
  }
  CreateAndConditionFromNativeArray(p*) { ; Not Implemented
    return UIA_NotImplemented()
  /*  [in]           IUIAutomationCondition **conditions,
    [in]           int conditionCount,
    [out, retval]  IUIAutomationCondition **newCondition
  */
    ;~ return UIA_Hr(DllCall(this.__Vt(27), "ptr",this.__Value,
  }
  CreateOrCondition(c1,c2) {
    return UIA_Hr(DllCall(this.__Vt(28), "ptr",this.__Value, "ptr",c1.__Value, "ptr",c2.__Value, "ptr*",out))? new UIA_OrCondition(out):
  }
  CreateOrConditionFromArray(array) {
  ;->in: AHK Array or Wrapped SafeArray
    if ComObjValue(array)&0x2000
      SafeArray:=array
    else {
      SafeArray:=ComObj(0x2003,DllCall("oleaut32\SafeArrayCreateVector", "uint",13, "uint",0, "uint",array.MaxIndex()),1)
      for i,c in array
        SafeArray[A_Index-1]:=c.__Value, ObjAddRef(c.__Value) ; AddRef - SafeArrayDestroy will release UIA_Conditions - they also release themselves
    }
    return UIA_Hr(DllCall(this.__Vt(29), "ptr",this.__Value, "ptr",ComObjValue(SafeArray), "ptr*",out))? new UIA_AndCondition(out):
  }
  CreateOrConditionFromNativeArray(p*) { ; Not Implemented
    return UIA_NotImplemented()
  /*  [in]           IUIAutomationCondition **conditions,
    [in]           int conditionCount,
    [out, retval]  IUIAutomationCondition **newCondition
  */
    ;~ return UIA_Hr(DllCall(this.__Vt(27), "ptr",this.__Value,
  }
  CreateNotCondition(c) {
    return UIA_Hr(DllCall(this.__Vt(31), "ptr",this.__Value, "ptr",c.__Value, "ptr*",out))? new UIA_NotCondition(out):
  }

  ;~ AddAutomationEventHandler   32
  ;~ RemoveAutomationEventHandler   33
  ;~ AddPropertyChangedEventHandlerNativeArray   34
  AddPropertyChangedEventHandler(element,scope=0x1,cacheRequest=0,handler="",propertyArray="") {
    SafeArray:=ComObjArray(0x3,propertyArray.MaxIndex())
    for i,propertyId in propertyArray
      SafeArray[i-1]:=propertyId
    return UIA_Hr(DllCall(this.__Vt(35), "ptr",this.__Value, "ptr",element.__Value, "int",scope, "ptr",cacheRequest,"ptr",handler.__Value,"ptr",ComObjValue(SafeArray)))
  }
  ;~ RemovePropertyChangedEventHandler   36
  ;~ AddStructureChangedEventHandler   37
  ;~ RemoveStructureChangedEventHandler   38
  AddFocusChangedEventHandler(cacheRequest, handler) {
    return UIA_Hr(DllCall(this.__Vt(39), "ptr",this.__Value, "ptr",cacheRequest, "ptr",handler.__Value))
  }
  ;~ RemoveFocusChangedEventHandler   40
  ;~ RemoveAllEventHandlers   41

  IntNativeArrayToSafeArray(ByRef nArr, n="") {
    return UIA_Hr(DllCall(this.__Vt(42), "ptr",this.__Value, "ptr",&nArr, "int",n?n:VarSetCapacity(nArr)/4, "ptr*",out))? ComObj(0x2003,out,1):
  }
/*  IntSafeArrayToNativeArray(sArr, Byref nArr="", Byref arrayCount="") { ; NOT WORKING
    VarSetCapacity(nArr,(sArr.MaxIndex()+1)*4)
    return UIA_Hr(DllCall(this.__Vt(43), "ptr",this.__Value, "ptr",ComObjValue(sArr), "ptr*",nArr, "int*",arrayCount))? arrayCount:
  }
*/
  RectToVariant(ByRef rect, ByRef out="") {  ; in:{left,top,right,bottom} ; out:(left,top,width,height)
    ; in:  RECT Struct
    ; out:  AHK Wrapped SafeArray & ByRef Variant
    return UIA_Hr(DllCall(this.__Vt(44), "ptr",this.__Value, "ptr",&rect, "ptr",UIA_Variant(out)))? UIA_VariantData(out):
  }
/*  VariantToRect(ByRef var, ByRef out="") { ; NOT WORKING
    ; in:  VT_VARIANT (SafeArray)
    ; out:  AHK Wrapped RECT Struct & ByRef Struct
    return UIA_Hr(DllCall(this.__Vt(45), "ptr",this.__Value, "ptr",var, "ptr",&(out,VarSetCapacity(out,16))))? UIA_RectToObject(out):
  }
*/
  ;~ SafeArrayToRectNativeArray   46
  ;~ CreateProxyFactoryEntry   47
  GetPropertyProgrammaticName(Id) {
    return UIA_Hr(DllCall(this.__Vt(49), "ptr",this.__Value, "int",Id, "ptr*",out))? StrGet(out):
  }
  GetPatternProgrammaticName(Id) {
    return UIA_Hr(DllCall(this.__Vt(50), "ptr",this.__Value, "int",Id, "ptr*",out))? StrGet(out):
  }
  PollForPotentialSupportedPatterns(e, Byref Ids="", Byref Names="") {
    return UIA_Hr(DllCall(this.__Vt(51), "ptr",this.__Value, "ptr",e.__Value, "ptr*",Ids, "ptr*",Names))? UIA_SafeArraysToObject(Names:=ComObj(0x2008,Names,1),Ids:=ComObj(0x2003,Ids,1)):
  }
  PollForPotentialSupportedProperties(e, Byref Ids="", Byref Names="") {
    return UIA_Hr(DllCall(this.__Vt(52), "ptr",this.__Value, "ptr",e.__Value, "ptr*",Ids, "ptr*",Names))? UIA_SafeArraysToObject(Names:=ComObj(0x2008,Names,1),Ids:=ComObj(0x2003,Ids,1)):
  }
  CheckNotSupported(value) { ; Useless in this Framework???
  /*  Checks a provided VARIANT to see if it contains the Not Supported identifier.
    After retrieving a property for a UI Automation element, call this method to determine whether the element supports the 
    retrieved property. CheckNotSupported is typically called after calling a property retrieving method such as GetCurrentPropertyValue.
  */
    return UIA_Hr(DllCall(this.__Vt(53), "ptr",this.__Value, "ptr",value, "int*",out))? out:
  }
  ElementFromIAccessible(IAcc, childId=0) {
  /* The method returns E_INVALIDARG - "One or more arguments are not valid" - if the underlying implementation of the
  Microsoft UI Automation element is not a native Microsoft Active Accessibility server; that is, if a client attempts to retrieve
  the IAccessible interface for an element originally supported by a proxy object from Oleacc.dll, or by the UIA-to-MSAA Bridge.
  */
    return UIA_Hr(DllCall(this.__Vt(56), "ptr",this.__Value, "ptr",ComObjValue(IAcc), "int",childId, "ptr*",out))? new UIA_Element(out):
  }
  ;~ ElementFromIAccessibleBuildCache   57
}

class UIA_Element extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671425(v=vs.85).aspx
  static __IID := "{d22108aa-8ac5-49a5-837b-37bbb3d7591e}"
    ,  __properties := "CurrentProcessId,20,int`r`nCurrentControlType,21,CONTROLTYPEID`r`nCurrentLocalizedControlType,22,BSTR`r`nCurrentName,23,BSTR`r`nCurrentAcceleratorKey,24,BSTR`r`nCurrentAccessKey,25,BSTR`r`nCurrentHasKeyboardFocus,26,BOOL`r`nCurrentIsKeyboardFocusable,27,BOOL`r`nCurrentIsEnabled,28,BOOL`r`nCurrentAutomationId,29,BSTR`r`nCurrentClassName,30,BSTR`r`nCurrentHelpText,31,BSTR`r`nCurrentCulture,32,int`r`nCurrentIsControlElement,33,BOOL`r`nCurrentIsContentElement,34,BOOL`r`nCurrentIsPassword,35,BOOL`r`nCurrentNativeWindowHandle,36,UIA_HWND`r`nCurrentItemType,37,BSTR`r`nCurrentIsOffscreen,38,BOOL`r`nCurrentOrientation,39,OrientationType`r`nCurrentFrameworkId,40,BSTR`r`nCurrentIsRequiredForForm,41,BOOL`r`nCurrentItemStatus,42,BSTR`r`nCurrentBoundingRectangle,43,RECT`r`nCurrentLabeledBy,44,IUIAutomationElement`r`nCurrentAriaRole,45,BSTR`r`nCurrentAriaProperties,46,BSTR`r`nCurrentIsDataValidForForm,47,BOOL`r`nCurrentControllerFor,48,IUIAutomationElementArray`r`nCurrentDescribedBy,49,IUIAutomationElementArray`r`nCurrentFlowsTo,50,IUIAutomationElementArray`r`nCurrentProviderDescription,51,BSTR`r`nCachedProcessId,52,int`r`nCachedControlType,53,CONTROLTYPEID`r`nCachedLocalizedControlType,54,BSTR`r`nCachedName,55,BSTR`r`nCachedAcceleratorKey,56,BSTR`r`nCachedAccessKey,57,BSTR`r`nCachedHasKeyboardFocus,58,BOOL`r`nCachedIsKeyboardFocusable,59,BOOL`r`nCachedIsEnabled,60,BOOL`r`nCachedAutomationId,61,BSTR`r`nCachedClassName,62,BSTR`r`nCachedHelpText,63,BSTR`r`nCachedCulture,64,int`r`nCachedIsControlElement,65,BOOL`r`nCachedIsContentElement,66,BOOL`r`nCachedIsPassword,67,BOOL`r`nCachedNativeWindowHandle,68,UIA_HWND`r`nCachedItemType,69,BSTR`r`nCachedIsOffscreen,70,BOOL`r`nCachedOrientation,71,OrientationType`r`nCachedFrameworkId,72,BSTR`r`nCachedIsRequiredForForm,73,BOOL`r`nCachedItemStatus,74,BSTR`r`nCachedBoundingRectangle,75,RECT`r`nCachedLabeledBy,76,IUIAutomationElement`r`nCachedAriaRole,77,BSTR`r`nCachedAriaProperties,78,BSTR`r`nCachedIsDataValidForForm,79,BOOL`r`nCachedControllerFor,80,IUIAutomationElementArray`r`nCachedDescribedBy,81,IUIAutomationElementArray`r`nCachedFlowsTo,82,IUIAutomationElementArray`r`nCachedProviderDescription,83,BSTR"
  
  SetFocus() {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
  }
  GetRuntimeId(ByRef stringId="") {
    return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr*",sa))? ComObj(0x2003,sa,1):
  }
  FindFirst(c="", scope=0x2) {
    static tc  ; TrueCondition
    if !tc
      tc:=this.__uia.CreateTrueCondition()
    return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "uint",scope, "ptr",(c=""?tc:c).__Value, "ptr*",out))? new UIA_Element(out):
  }
  FindAll(c="", scope=0x2) {
    static tc  ; TrueCondition
    if !tc
      tc:=this.__uia.CreateTrueCondition()
    return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "uint",scope, "ptr",(c=""?tc:c).__Value, "ptr*",out))? UIA_ElementArray(out):
  }
  ;~ Find (First/All, Element/Children/Descendants/Parent/Ancestors/Subtree, Conditions)
  ;~ FindFirstBuildCache   7  IUIAutomationElement
  ;~ FindAllBuildCache   8  IUIAutomationElementArray
  ;~ BuildUpdatedCache   9  IUIAutomationElement
  GetCurrentPropertyValue(propertyId, ByRef out="") {
    try {
        return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "uint",propertyId, "ptr",UIA_Variant(out)))? UIA_VariantData(out):
    }
  }
  GetCurrentPropertyValueEx(propertyId, ignoreDefaultValue=1, ByRef out="") {
  ; Passing FALSE in the ignoreDefaultValue parameter is equivalent to calling GetCurrentPropertyValue
    return UIA_Hr(DllCall(this.__Vt(11), "ptr",this.__Value, "uint",propertyId, "uint",ignoreDefaultValue, "ptr",UIA_Variant(out)))? UIA_VariantData(out):
  }
  ;~ GetCachedPropertyValue   12  VARIANT
  ;~ GetCachedPropertyValueEx   13  VARIANT
  GetCurrentPatternAs(pattern="") {
    if IsObject(UIA_%pattern%Pattern)&&(iid:=UIA_%pattern%Pattern.__iid)&&(pId:=UIA_%pattern%Pattern.__PatternID)
      return UIA_Hr(DllCall(this.__Vt(14), "ptr",this.__Value, "int",pId, "ptr",UIA_GUID(riid,iid), "ptr*",out))? new UIA_%pattern%Pattern(out):
    else throw Exception("Pattern not implemented.",-1, "UIA_" pattern "Pattern")
  }
  ;~ GetCachedPatternAs   15  void **ppv
  ;~ GetCurrentPattern   16  Iunknown **patternObject
  ;~ GetCachedPattern   17  Iunknown **patternObject
  ;~ GetCachedParent   18  IUIAutomationElement
  GetCachedChildren() { ; Haven't successfully tested
    return UIA_Hr(DllCall(this.__Vt(19), "ptr",this.__Value, "ptr*",out))&&out? UIA_ElementArray(out):
  }
  ;~ GetClickablePoint   84  POINT, BOOL
}

class UIA_ElementArray extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671426(v=vs.85).aspx
  static __IID := "{14314595-b4bc-4055-95f2-58f2e42c9855}"
    ,  __properties := "Length,3,int"
  
  GetElement(i) {
    try {
        return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "int",i, "ptr*",out))? new UIA_Element(out):
    }
  }
}

class UIA_TreeWalker extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671470(v=vs.85).aspx
  static __IID := "{4042c624-389c-4afc-a630-9df854a541fc}"
    ,  __properties := "Condition,15,IUIAutomationCondition"
  
  GetParentElement(e) {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))? new UIA_Element(out):
  }
  GetFirstChildElement(e) {
    return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? new UIA_Element(out):
  }
  GetLastChildElement(e) {
    return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? new UIA_Element(out):
  }
  GetNextSiblingElement(e) {
    return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? new UIA_Element(out):
  }
  GetPreviousSiblingElement(e) {
    return UIA_Hr(DllCall(this.__Vt(7), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? new UIA_Element(out):
  }
  NormalizeElement(e) {
    return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? new UIA_Element(out):
  }
/*  GetParentElementBuildCache(e, cacheRequest) {
    return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value.__Value, "ptr*",out))? new UIA_Element(out):
  }
  GetFirstChildElementBuildCache(e, cacheRequest) {
    return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? new UIA_Element(out):
  }
  GetLastChildElementBuildCache(e, cacheRequest) {
    return UIA_Hr(DllCall(this.__Vt(11), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? new UIA_Element(out):
  }
  GetNextSiblingElementBuildCache(e, cacheRequest) {
    return UIA_Hr(DllCall(this.__Vt(12), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? new UIA_Element(out):
  }
  GetPreviousSiblingElementBuildCache(e, cacheRequest) {
    return UIA_Hr(DllCall(this.__Vt(13), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? new UIA_Element(out):
  }
  NormalizeElementBuildCache(e, cacheRequest) {
    return UIA_Hr(DllCall(this.__Vt(14), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? new UIA_Element(out):
  }
*/
}

class UIA_Condition extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671420(v=vs.85).aspx
  static __IID := "{352ffba8-0973-437c-a61f-f64cafd81df9}"
}

class UIA_PropertyCondition extends UIA_Condition {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696121(v=vs.85).aspx
  static __IID := "{99ebf2cb-5578-4267-9ad4-afd6ea77e94b}"
    ,  __properties := "PropertyId,3,PROPERTYID`r`nPropertyValue,4,VARIANT`r`nPropertyConditionFlags,5,PropertyConditionFlags"
}
; should returned children have a condition type (property/and/or/bool/not), or be a generic uia_condition object?
class UIA_AndCondition extends UIA_Condition {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671407(v=vs.85).aspx
  static __IID := "{a7d0af36-b912-45fe-9855-091ddc174aec}"
    ,  __properties := "ChildCount,3,int"
  
  ;~ GetChildrenAsNativeArray  4  IUIAutomationCondition ***childArray
  GetChildren() {
    return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr*",out))&&out? ComObj(0x2003,out,1):
  }
}
class UIA_OrCondition extends UIA_Condition {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696108(v=vs.85).aspx
  static __IID := "{8753f032-3db1-47b5-a1fc-6e34a266c712}"
    ,  __properties := "ChildCount,3,int"
  
  ;~ GetChildrenAsNativeArray  4  IUIAutomationCondition ***childArray
  ;~ GetChildren  5  SAFEARRAY
}
class UIA_BoolCondition extends UIA_Condition {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671411(v=vs.85).aspx
  static __IID := "{8753f032-3db1-47b5-a1fc-6e34a266c712}"
    ,  __properties := "BooleanValue,3,boolVal"
}
class UIA_NotCondition extends UIA_Condition {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696106(v=vs.85).aspx
  static __IID := "{f528b657-847b-498c-8896-d52b565407a1}"
  
  ;~ GetChild  3  IUIAutomationCondition
}

class UIA_IUnknown extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ms680509(v=vs.85).aspx
  static __IID := "{00000000-0000-0000-C000-000000000046}"
}

class UIA_CacheRequest  extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671413(v=vs.85).aspx
  static __IID := "{b32a92b5-bc25-4078-9c08-d7ee95c48e03}"
}


class _UIA_EventHandler {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696044(v=vs.85).aspx
  static __IID := "{146c3c17-f12e-4e22-8c27-f894b9b79c69}"
  
/*  HandleAutomationEvent  3
    [in]  IUIAutomationElement *sender,
    [in]  EVENTID eventId
*/
}
class _UIA_FocusChangedEventHandler {    
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696051(v=vs.85).aspx
  static __IID := "{c270f6b5-5c69-4290-9745-7a7f97169468}"
  
/*  HandleFocusChangedEvent  3
    [in]  IUIAutomationElement *sender
*/
}
class _UIA_PropertyChangedEventHandler {    
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696119(v=vs.85).aspx
  static __IID := "{40cd37d4-c756-4b0c-8c6f-bddfeeb13b50}"
  
/*  HandlePropertyChangedEvent  3
    [in]  IUIAutomationElement *sender,
    [in]  PROPERTYID propertyId,
    [in]  VARIANT newValue
*/
}
class _UIA_StructureChangedEventHandler {    
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696197(v=vs.85).aspx
  static __IID := "{e81d1b4e-11c5-42f8-9754-e7036c79f054}"
  
/*  HandleStructureChangedEvent  3
    [in]  IUIAutomationElement *sender,
    [in]  StructureChangeType changeType,
    [in]  SAFEARRAY *runtimeId[int]
*/
}
class _UIA_TextEditTextChangedEventHandler { ; Windows 8.1 Preview [desktop apps only]
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/dn302202(v=vs.85).aspx
  static __IID := "{92FAA680-E704-4156-931A-E32D5BB38F3F}"
  
  ;~ HandleTextEditTextChangedEvent  3
}


;~     UIA_Patterns - http://msdn.microsoft.com/en-us/library/windows/desktop/ee684023
class UIA_DockPattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671421
  static  __IID := "{fde5ef97-1464-48f6-90bf-43d0948e86ec}"
    ,  __PatternID := 10011
    ,  __Properties := "CurrentDockPosition,4,int`r`nCachedDockPosition,5,int"

  SetDockPosition(Pos) {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "uint",pos))
  }
/*  DockPosition_Top  = 0,
  DockPosition_Left  = 1,
  DockPosition_Bottom  = 2,
  DockPosition_Right  = 3,
  DockPosition_Fill  = 4,
  DockPosition_None  = 5
*/
}
class UIA_ExpandCollapsePattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696046
  static  __IID := "{619be086-1f4e-4ee4-bafa-210128738730}"
    ,  __PatternID := 10005
    ,  __Properties := "CachedExpandCollapseState,6,int`r`nCurrentExpandCollapseState,5,int"
  
  Expand() {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
  }
  Collapse() {
    return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value))
  }  
/*  ExpandCollapseState_Collapsed  = 0,
  ExpandCollapseState_Expanded  = 1,
  ExpandCollapseState_PartiallyExpanded  = 2,
  ExpandCollapseState_LeafNode  = 3
*/
}
class UIA_GridItemPattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696053
  static  __IID := "{78f8ef57-66c3-4e09-bd7c-e79b2004894d}"
    ,  __PatternID := 10007
    ,  __Properties := "CurrentContainingGrid,3,IUIAutomationElement`r`nCurrentRow,4,int`r`nCurrentColumn,5,int`r`nCurrentRowSpan,6,int`r`nCurrentColumnSpan,7,int`r`nCachedContainingGrid,8,IUIAutomationElement`r`nCachedRow,9,int`r`nCachedColumn,10,int`r`nCachedRowSpan,11,int`r`nCachedColumnSpan,12,int"
}
class UIA_GridPattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696064
  static  __IID := "{414c3cdc-856b-4f5b-8538-3131c6302550}"
    ,  __PatternID := 10006
    ,  __Properties := "CurrentRowCount,4,int`r`nCurrentColumnCount,5,int`r`nCachedRowCount,6,int`r`nCachedColumnCount,7,int"

  GetItem(row,column) { ; Hr!=0 if no result, or blank output?
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "uint",row, "uint",column, "ptr*",out))? new UIA_Element(out):
  }
}
class UIA_InvokePattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696070
  static  __IID := "{fb377fbe-8ea6-46d5-9c73-6499642d3059}"
    ,  __PatternID := 10000
  
  Invoke() {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
  }
}
class UIA_ItemContainerPattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696072
  static  __IID := "{c690fdb2-27a8-423c-812d-429773c9084e}"
    ,  __PatternID := 10019

  FindItemByProperty(startAfter, propertyId, ByRef value, type=8) {  ; Hr!=0 if no result, or blank output?
    if (type!="Variant")
      UIA_Variant(value,type,value)
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",startAfter.__Value, "int",propertyId, "ptr",&value, "ptr*",out))? new UIA_Element(out):
  }
}
class UIA_LegacyIAccessiblePattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696074
  static  __IID := "{828055ad-355b-4435-86d5-3b51c14a9b1b}"
    ,  __PatternID := 10018
    ,  __Properties := "CurrentChildId,6,int`r`nCurrentName,7,BSTR`r`nCurrentValue,8,BSTR`r`nCurrentDescription,9,BSTR`r`nCurrentRole,10,DWORD`r`nCurrentState,11,DWORD`r`nCurrentHelp,12,BSTR`r`nCurrentKeyboardShortcut,13,BSTR`r`nCurrentDefaultAction,15,BSTR`r`nCachedChildId,16,int`r`nCachedName,17,BSTR`r`nCachedValue,18,BSTR`r`nCachedDescription,19,BSTR`r`nCachedRole,20,DWORD`r`nCachedState,21,DWORD`r`nCachedHelp,22,BSTR`r`nCachedKeyboardShortcut,23,BSTR`r`nCachedDefaultAction,25,BSTR"

  Select(flags=3) {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "int",flags))
  }
  DoDefaultAction() {
    return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value))
  }
  SetValue(value) {
    return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr",&value))
  }
  GetCurrentSelection() { ; Not correct
    ;~ if (hr:=DllCall(this.__Vt(14), "ptr",this.__Value, "ptr*",array))=0
      ;~ return new UIA_ElementArray(array)
    ;~ else
      ;~ MsgBox,, Error, %hr%
  }
  ;~ GetCachedSelection  24  IUIAutomationElementArray
  GetIAccessible() {
  /*  This method returns NULL if the underlying implementation of the UI Automation element is not a native 
  Microsoft Active Accessibility server; that is, if a client attempts to retrieve the IAccessible interface 
  for an element originally supported by a proxy object from OLEACC.dll, or by the UIA-to-MSAA Bridge.
  */
    return UIA_Hr(DllCall(this.__Vt(26), "ptr",this.__Value, "ptr*",pacc))&&pacc? ComObj(9,pacc,1):
  }
}
class UIA_MultipleViewPattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696099
  static  __IID := "{8d253c91-1dc5-4bb5-b18f-ade16fa495e8}"
    ,  __PatternID := 10008
    ,  __Properties := "CurrentCurrentView,5,int`r`nCachedCurrentView,7,int"

  GetViewName(view) { ; need to release BSTR?
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "int",view, "ptr*",name))? StrGet(name):
  }
  SetCurrentView(view) {
    return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "int",view))
  }
  GetCurrentSupportedViews() {
    return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr*",out))? ComObj(0x2003,out,1):
  }
  GetCachedSupportedViews() {
    return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value, "ptr*",out))? ComObj(0x2003,out,1):
  }
}
class UIA_RangeValuePattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696147
  static  __IID := "{59213f4f-7346-49e5-b120-80555987a148}"
    ,  __PatternID := 10003
    ,  __Properties := "CurrentValue,4,double`r`nCurrentIsReadOnly,5,BOOL`r`nCurrentMaximum,6,double`r`nCurrentMinimum,7,double`r`nCurrentLargeChange,8,double`r`nCurrentSmallChange,9,double`r`nCachedValue,10,double`r`nCachedIsReadOnly,11,BOOL`r`nCachedMaximum,12,double`r`nCachedMinimum,13,double`r`nCachedLargeChange,14,double`r`nCachedSmallChange,15,double"

  SetValue(val) {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "double",val))
  }
}
class UIA_ScrollItemPattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696165
  static  __IID := "{b488300f-d015-4f19-9c29-bb595e3645ef}"
    ,  __PatternID := 10017

  ScrollIntoView() {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
  }
}
class UIA_ScrollPattern extends UIA_Base {
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696167
  static  __IID := "{88f4d42a-e881-459d-a77c-73bbbb7e02dc}"
    ,  __PatternID := 10004
    ,  __Properties := "CurrentHorizontalScrollPercent,5,double`r`nCurrentVerticalScrollPercent,6,double`r`nCurrentHorizontalViewSize,7,double`r`CurrentHorizontallyScrollable,9,BOOL`r`nCurrentVerticallyScrollable,10,BOOL`r`nCachedHorizontalScrollPercent,11,double`r`nCachedVerticalScrollPercent,12,double`r`nCachedHorizontalViewSize,13,double`r`nCachedVerticalViewSize,14,double`r`nCachedHorizontallyScrollable,15,BOOL`r`nCachedVerticallyScrollable,16,BOOL"
    
  Scroll(horizontal, vertical) {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "uint",horizontal, "uint",vertical))
  }
  SetScrollPercent(horizontal, vertical) {
    return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "double",horizontal, "double",vertical))
  }
/*  UIA_ScrollPatternNoScroll  =  -1
  ScrollAmount_LargeDecrement  = 0,
  ScrollAmount_SmallDecrement  = 1,
  ScrollAmount_NoAmount  = 2,
  ScrollAmount_LargeIncrement  = 3,
  ScrollAmount_SmallIncrement  = 4
*/
}
;~ class UIA_SelectionItemPattern extends UIA_Base {10010
;~ class UIA_SelectionPattern extends UIA_Base {10001
;~ class UIA_SpreadsheetItemPattern extends UIA_Base {10027
;~ class UIA_SpreadsheetPattern extends UIA_Base {10026
;~ class UIA_StylesPattern extends UIA_Base {10025
;~ class UIA_SynchronizedInputPattern extends UIA_Base {10021
;~ class UIA_TableItemPattern extends UIA_Base {10013
;~ class UIA_TablePattern extends UIA_Base {10012
;~ class UIA_TextChildPattern extends UIA_Base {10029
;~ class UIA_TextEditPattern extends UIA_Base {10032
;~ class UIA_TextPattern extends UIA_Base {10014
;~ class UIA_TextPattern2 extends UIA_Base {10024
;~ class UIA_TogglePattern extends UIA_Base {10015
;~ class UIA_TransformPattern extends UIA_Base {10016
;~ class UIA_TransformPattern2 extends UIA_Base {10028
;~ class UIA_ValuePattern extends UIA_Base {10002
;~ class UIA_VirtualizedItemPattern extends UIA_Base {10020
;~ class UIA_WindowPattern extends UIA_Base {10009
;~ class UIA_AnnotationPattern extends UIA_Base {10023    ; Windows 8 [desktop apps only]
;~ class UIA_DragPattern extends UIA_Base {10030      ; Windows 8 [desktop apps only]
;~ class UIA_DropTargetPattern extends UIA_Base {10031    ; Windows 8 [desktop apps only]
/* class UIA_ObjectModelPattern extends UIA_Base {      ; Windows 8 [desktop apps only]
  ;~ http://msdn.microsoft.com/en-us/library/windows/desktop/hh437262(v=vs.85).aspx
  static  __IID := "{71c284b3-c14d-4d14-981e-19751b0d756d}"
    ,  __PatternID := 10022
  
  GetUnderlyingObjectModel() {
    return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
  }
}
*/

;~ class UIA_PatternHandler extends UIA_Base {
;~ class UIA_PatternInstance extends UIA_Base {
;~ class UIA_TextRange extends UIA_Base {
;~ class UIA_TextRange2 extends UIA_Base {
;~ class UIA_TextRangeArray extends UIA_Base {




{  ;~ UIA Functions
  UIA_Interface() {
    try {
      if uia:=ComObjCreate("{ff48dba4-60ef-4201-aa87-54103eef594e}","{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}")
        return uia:=new UIA_Interface(uia), uia.base.base.__UIA:=uia
      throw "UIAutomation Interface failed to initialize."
    } catch e
      MsgBox, 262160, UIA Startup Error, % IsObject(e)?"IUIAutomation Interface is not registered.":e.Message
  }
  UIA_Hr(hr) {
    ;~ http://blogs.msdn.com/b/eldar/archive/2007/04/03/a-lot-of-hresult-codes.aspx
    static err:={0x8000FFFF:"Catastrophic failure.",0x80004001:"Not implemented.",0x8007000E:"Out of memory.",0x80070057:"One or more arguments are not valid.",0x80004002:"Interface not supported.",0x80004003:"Pointer not valid.",0x80070006:"Handle not valid.",0x80004004:"Operation aborted.",0x80004005:"Unspecified error.",0x80070005:"General access denied.",0x800401E5:"The object identified by this moniker could not be found.",0x80040201:"UIA_E_ELEMENTNOTAVAILABLE",0x80040200:"UIA_E_ELEMENTNOTENABLED",0x80131509:"UIA_E_INVALIDOPERATION",0x80040202:"UIA_E_NOCLICKABLEPOINT",0x80040204:"UIA_E_NOTSUPPORTED",0x80040203:"UIA_E_PROXYASSEMBLYNOTLOADED"} ; //not completed
    if hr&&(hr&=0xFFFFFFFF) {
      RegExMatch(Exception("",-2).what,"(\w+).(\w+)",i)
      throw Exception(UIA_Hex(hr) " - " err[hr], -2, i2 "  (" i1 ")")
    }
    return !hr
  }
  UIA_NotImplemented() {
    RegExMatch(Exception("",-2).What,"(\D+)\.(\D+)",m)
    MsgBox, 262192, UIA Message, Class:`t%m1%`nMember:`t%m2%`n`nMethod has not been implemented yet.
  }
  UIA_ElementArray(p, uia="") { ; should AHK Object be 0 or 1 based?
    a:=new UIA_ElementArray(p),out:=[]
    Loop % a.Length
      out[A_Index]:=a.GetElement(A_Index-1)
    return out, out.base:={UIA_ElementArray:a}
  }
  UIA_RectToObject(ByRef r) { ; rect.__Value work with DllCalls?
    static b:={__Class:"object",__Type:"RECT",Struct:Func("UIA_RectStructure")}
    return {l:NumGet(r,0,"Int"),t:NumGet(r,4,"Int"),r:NumGet(r,8,"Int"),b:NumGet(r,12,"Int"),base:b}
  }
  UIA_RectStructure(this, ByRef r) {
    static sides:="ltrb"
    VarSetCapacity(r,16)
    Loop Parse, sides
      NumPut(this[A_LoopField],r,(A_Index-1)*4,"Int")
  }
  UIA_SafeArraysToObject(keys,values) {
  ;~  1 dim safearrays w/ same # of elements
    out:={}
    for key in keys
      out[key]:=values[A_Index-1]
    return out
  }
  UIA_Hex(p) {
    setting:=A_FormatInteger
    SetFormat,IntegerFast,H
    out:=p+0 ""
    SetFormat,IntegerFast,%setting%
    return out
  }
  UIA_GUID(ByRef GUID, sGUID) { ;~ Converts a string to a binary GUID and returns its address.
    VarSetCapacity(GUID,16,0)
    return DllCall("ole32\CLSIDFromString", "wstr",sGUID, "ptr",&GUID)>=0?&GUID:""
  }
  UIA_Variant(ByRef var,type=0,val=0) {
    ; Does a variant need to be cleared? If it uses SysAllocString? 
    return (VarSetCapacity(var,8+2*A_PtrSize)+NumPut(type,var,0,"short")+NumPut(type=8? DllCall("oleaut32\SysAllocString", "ptr",&val):val,var,8,"ptr"))*0+&var
  }
  UIA_IsVariant(ByRef vt, ByRef type="") {
    size:=VarSetCapacity(vt),type:=NumGet(vt,"UShort")
    return size>=16&&size<=24&&type>=0&&(type<=23||type|0x2000)
  }
  UIA_Type(ByRef item, ByRef info) {
  }
  UIA_VariantData(ByRef p, flag=1) {
    ; based on Sean's COM_Enumerate function
    ; need to clear varaint? what if you still need it (flag param)?
    return !UIA_IsVariant(p,vt)?"Invalid Variant"
        :vt=3?NumGet(p,8,"int")
        :vt=8?StrGet(NumGet(p,8))
        :vt=9||vt=13||vt&0x2000?ComObj(vt,NumGet(p,8),flag)
        :vt<0x1000&&UIA_VariantChangeType(&p,&p)=0?StrGet(NumGet(p,8)) UIA_VariantClear(&p)
        :NumGet(p,8)
  /*
    VT_EMPTY     =      0      ; No value
    VT_NULL      =      1     ; SQL-style Null
    VT_I2        =      2     ; 16-bit signed int
    VT_I4        =      3     ; 32-bit signed int
    VT_R4        =      4     ; 32-bit floating-point number
    VT_R8        =      5     ; 64-bit floating-point number
    VT_CY        =      6     ; Currency
    VT_DATE      =      7      ; Date
    VT_BSTR      =      8     ; COM string (Unicode string with length prefix)
    VT_DISPATCH  =      9     ; COM object 
    VT_ERROR     =    0xA  10  ; Error code (32-bit integer)
    VT_BOOL      =    0xB  11  ; Boolean True (-1) or False (0)
    VT_VARIANT   =    0xC  12  ; VARIANT (must be combined with VT_ARRAY or VT_BYREF)
    VT_UNKNOWN   =    0xD  13  ; IUnknown interface pointer
    VT_DECIMAL   =    0xE  14  ; (not supported)
    VT_I1        =   0x10  16  ; 8-bit signed int
    VT_UI1       =   0x11  17  ; 8-bit unsigned int
    VT_UI2       =   0x12  18  ; 16-bit unsigned int
    VT_UI4       =   0x13  19  ; 32-bit unsigned int
    VT_I8        =   0x14  20  ; 64-bit signed int
    VT_UI8       =   0x15  21  ; 64-bit unsigned int
    VT_INT       =   0x16  22  ; Signed machine int
    VT_UINT      =   0x17  23  ; Unsigned machine int
    VT_RECORD    =   0x24  36  ; User-defined type
    VT_ARRAY     = 0x2000      ; SAFEARRAY
    VT_BYREF     = 0x4000      ; Pointer to another type of value
           = 0x1000  4096

    COM_VariantChangeType(pvarDst, pvarSrc, vt=8) {
      return DllCall("oleaut32\VariantChangeTypeEx", "ptr",pvarDst, "ptr",pvarSrc, "Uint",1024, "Ushort",0, "Ushort",vt)
    }
    COM_VariantClear(pvar) {
      DllCall("oleaut32\VariantClear", "ptr",pvar)
    }
    COM_SysAllocString(str) {
      Return  DllCall("oleaut32\SysAllocString", "Uint", &str)
    }
    COM_SysFreeString(pstr) {
        DllCall("oleaut32\SysFreeString", "Uint", pstr)
    }
    COM_SysString(ByRef wString, sString) {
      VarSetCapacity(wString,4+nLen:=2*StrLen(sString))
      Return  DllCall("kernel32\lstrcpyW","Uint",NumPut(nLen,wString),"Uint",&sString)
    }
  */
  }
  UIA_VariantChangeType(pvarDst, pvarSrc, vt=8) { ; written by Sean
    return DllCall("oleaut32\VariantChangeTypeEx", "ptr",pvarDst, "ptr",pvarSrc, "Uint",1024, "Ushort",0, "Ushort",vt)
  }
  UIA_VariantClear(pvar) { ; Written by Sean
    DllCall("oleaut32\VariantClear", "ptr",pvar)
  }
}

/*
enum TreeScope
    {  TreeScope_Element  = 0x1,
  TreeScope_Children  = 0x2,
  TreeScope_Descendants  = 0x4,
  TreeScope_Parent  = 0x8,
  TreeScope_Ancestors  = 0x10,
  TreeScope_Subtree  = ( ( TreeScope_Element | TreeScope_Children )  | TreeScope_Descendants ) 
    } ;

DllCall("oleaut32\SafeArrayGetVartype", "ptr*",ComObjValue(SafeArray), "uint*",pvt)
HRESULT SafeArrayGetVartype(
  _In_   SAFEARRAY *psa,
  _Out_  VARTYPE *pvt
);

DllCall("oleaut32\SafeArrayDestroy", "ptr",ComObjValue(SafeArray))
HRESULT SafeArrayDestroy(
  _In_  SAFEARRAY *psa
);
*/

GetMenu(hWnd) {
  SendMessage, 0x1E1, 0, 0, , ahk_id %hWnd%  ;  MN_GETHMENU
  hMenu := ErrorLevel
  If !hMenu || (hMenu + 0 = "")
     Return
  Return RTrim(GetMenuText(hMenu))
}

GetMenuText(hMenu, child = 0) {
  Loop, % DllCall("GetMenuItemCount", "Ptr", hMenu)
  {
    idx := A_Index - 1
    nSize++ := DllCall("GetMenuString", "Ptr", hMenu, "int", idx, "Uint", 0, "int", 0, "Uint", 0x400)   ;  MF_BYPOSITION
    nSize := (nSize * (A_IsUnicode ? 2 : 1))
    VarSetCapacity(sString, nSize)
    DllCall("GetMenuString", "Ptr", hMenu, "int", idx, "str", sString, "int", nSize, "Uint", 0x400)   ;  MF_BYPOSITION
    idn := DllCall("GetMenuItemID", "Ptr", hMenu, "int", idx)
    hSubMenu := DllCall("GetSubMenu", "Ptr", hMenu, "int", idx)
    isSubMenu := (idn=-1 && hSubMenu) ? 1 : 0
    If isSubMenu
       string .= " [sub-menu]" ; GetMenuText(hSubMenu, ++child), --child
  }
  Return string
}
