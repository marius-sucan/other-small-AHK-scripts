; made by Marius Șucan
; based on functions extracted from Font Library 3 by Jballi
; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=4379
;
; and KeyPress OSD v4
; https://github.com/marius-sucan/KeyPress-OSD
;
; File source: https://github.com/marius-sucan/other-small-AHK-scripts

#Persistent
#NoEnv
#SingleInstance Force
SetBatchLines, -1
Global hFontDummy

OnExit, cleanup
createWindow()
Return

createWindow() {
  Global editField
  hFontDummy := Fnt_CreateFont("Arial","s22 Q0")   ; font used to detect WideChars using Fnt_TruncateStringToFit
  Gui, demoGuia: -MaximizeBox -MinimizeBox
  Gui, demoGuia: Margin, 15, 15
  Gui, Add, Text, w500, Please enter the text you want. Click on Count to know how many distinct graphemes were identified.
  Gui, Add, Edit, y+10 w500 r1 -multi -wantReturn -wantTab -wrap veditField, Demo: उनका ग्लैमरस लुक चर्चामें... बीच पर एक्ट्रेस का बोल्ड अवतार देखने को मिला. ब्लैक बिकिनी में उनका ग्लैमरस लुक चर्चा में बना हुआ है. बीच ☺ 😊 😘 
  Gui, Add, Button, y+10 h25 Default gCountGraphemesBtn, &Count graphemes
  Gui, Add, Button, x+10 h25 gCleanup, &Exit
  Gui, Show, AutoSize, Count graphemes from text line
}

CountGraphemesBtn() {
  GuiControlGet, editField
  MsgBox, % "Graphemes identified: " CountGraphemes(editField) "`nStrLen() result: " StrLen(editField) "`nRegEx Contracted: " simpleRegExContraction(editField)
}

filterNewLines(string) {
    string := StrReplace(string, A_TAB, A_SPACE)
    string := StrReplace(string, "`r`n", A_Space)
    string := StrReplace(string, "`n", A_Space)
    string := StrReplace(string, "`r", A_Space)
    string := StrReplace(string, "`f", A_Space)
    Return string
}

simpleRegExContraction(string) {
  string := RegExReplace(string, "\X", "1")
  string := RegExReplace(string, "[\p{Mn}\p{C}]")
  Return StrLen(string)
}

CountGraphemes(testChars) {
; function by Marius Șucan based on Font Library by jballi
; extracted from KeyPress OSD v4 in March 2019
; function suited for Abugida scripts [Devanagari,
; Telugu, Bengali, Tamil and so on], and pictograms, Emojis.

  Static increment := 4, prevTxt, wideCharsFoundLast
  If !testChars
     Return 0

  filterNewLines(testChars)
  oldTxtLen := testTxtLen := StrLen(testChars)
  If (testChars=prevTxt && oldTxtLen>300)
     Return wideCharsFoundLast

  If (oldTxtLen>350)
  {
     prevTxt := testChars
     ToolTip, Please wait - processing text line...
  }

  wideCharsFound := 0
  fit2This := InitialTxtWidth := Fnt_GetStringWidth(hFontDummy, testChars) + 5
  ModifiedFnt_TruncateStringToFit(hFontDummy,,,,1)
  testCharsNew := testChars
  Loop
  {
     fit2This := fit2This - increment
     truncated := ModifiedFnt_TruncateStringToFit(hFontDummy, testCharsNew, fit2This, newWidth, 0, newTxtLen, oldTxtLen)
     fit2This := newWidth + 2
    ;  countLoops++
     If (newTxtLen<oldTxtLen)
     {
        wideCharsFound++
        testCharsNew := truncated
        oldTxtLen := newTxtLen
     }
     If (fit2This<3 || newTxtLen<1)
        Break
  }
  ModifiedFnt_TruncateStringToFit(hFontDummy,,,,2)
  If (testTxtLen>300)
     wideCharsFoundLast := wideCharsFound

  ; ToolTip, %countLoops% -- %testTxtLen% -- %InitialTxtWidth% -- %wideCharsFound%
  ToolTip
  Return wideCharsFound
}

ModifiedFnt_TruncateStringToFit(hFont:=0,p_String:=0,p_MaxW:=0, ByRef newWidth:=0, initMode:=1, ByRef newTxtLen:=0, oldTxtLen:=0) {
; function based on Fnt_TruncateStringToFit from Font Library 3.0

    Static  Dummy94264906
          , DEFAULT_GUI_FONT:= 17
          , HWND_DESKTOP    := 0
          , OBJ_FONT        := 6
          , hDC, old_hFont, l_Fit, SIZE, SIZE2

    If (initMode=1)
    {
       hDC := DllCall("GetDC","Ptr",HWND_DESKTOP)
       old_hFont := DllCall("SelectObject","Ptr",hDC,"Ptr",hFont)
       VarSetCapacity(l_Fit,4,0)
       VarSetCapacity(SIZE,8,0)
       VarSetCapacity(SIZE2,8,0)
       Return
    } Else If (initMode=2)
    {
       DllCall("SelectObject","Ptr",hDC,"Ptr",old_hFont)
       DllCall("ReleaseDC","Ptr",HWND_DESKTOP,"Ptr",hDC)
       Return
    }

    RS:=DllCall("GetTextExtentExPoint"
        ,"Ptr",hDC                                      ;-- hdc
        ,"Str",p_String                                 ;-- lpszStr
        ,"Int",oldTxtLen                         ;-- cchString
        ,"Int",p_MaxW                                   ;-- nMaxExtent
        ,"IntP",l_Fit                                   ;-- lpnFit [out]
        ,"Ptr",0                                        ;-- alpDx [out]
        ,"Ptr",&SIZE)                                   ;-- lpSize

    result := SubStr(p_String,1,l_Fit)                    ; truncated string
    newTxtLen := RS ? StrLen(result) : 0
    RC:=DllCall("GetTextExtentPoint32"
        ,"Ptr",hDC                                      ;-- hDC
        ,"Str",result                                 ;-- lpString
        ,"Int",newTxtLen                         ;-- c (string length)
        ,"Ptr",&SIZE2)                                   ;-- lpSize
    newWidth := RC ? NumGet(SIZE2,0,"Int") : 0
    Return result
}

; ==================================================
; Functions from Font Library 3.0 by jballi
; source: https://autohotkey.com/boards/viewtopic.php?t=4379

Fnt_TruncateStringToFit(hFont,p_String,p_MaxW) {
    Static Dummy94264906
          ,DEFAULT_GUI_FONT:= 17
          ,HWND_DESKTOP    := 0
          ,OBJ_FONT        := 6

    ;-- Parameters
    If not StrLen(p_String)
       Return p_String

    If p_MaxW is not Integer
       Return p_String
    else
       If (p_MaxW<1)  ;-- Zero or negative
          Return ""

    ;-- If needed, get the handle to the default GUI font
    If (DllCall("GetObjectType","Ptr",hFont)<>OBJ_FONT)
       hFont := DllCall("GetStockObject","Int",DEFAULT_GUI_FONT)

    ;-- Select the font into the device context for the desktop
    hDC       := DllCall("GetDC","Ptr",HWND_DESKTOP)
    old_hFont := DllCall("SelectObject","Ptr",hDC,"Ptr",hFont)

    ;-- Determine string size for specified maximum width
    VarSetCapacity(l_Fit,4,0)
    VarSetCapacity(SIZE,8,0)
    DllCall("GetTextExtentExPoint"
        ,"Ptr",hDC                                      ;-- hdc
        ,"Str",p_String                                 ;-- lpszStr
        ,"Int",StrLen(p_String)                         ;-- cchString
        ,"Int",p_MaxW                                   ;-- nMaxExtent
        ,"IntP",l_Fit                                   ;-- lpnFit [out]
        ,"Ptr",0                                        ;-- alpDx [out]
        ,"Ptr",&SIZE)                                   ;-- lpSize

    ;-- Release the objects needed by the GetTextExtentPoint32 function
    DllCall("SelectObject","Ptr",hDC,"Ptr",old_hFont)
    ;-- Necessary to avoid memory leak
    DllCall("ReleaseDC","Ptr",HWND_DESKTOP,"Ptr",hDC)

    ;-- Return string
    result := SubStr(p_String,1,l_Fit)                    ; truncated string
    Return result
}

Fnt_GetStringWidth(hFont,p_String) {
    Static Dummy88611714
          ,DEFAULT_GUI_FONT:=17
          ,HWND_DESKTOP    :=0
          ,OBJ_FONT        :=6

    ;-- If needed, get the handle to the default GUI font
    if (DllCall("GetObjectType","Ptr",hFont)<>OBJ_FONT)
        hFont:=DllCall("GetStockObject","Int",DEFAULT_GUI_FONT)

    ;-- Select the font into the device context for the desktop
    hDC      :=DllCall("GetDC","Ptr",HWND_DESKTOP)
    old_hFont:=DllCall("SelectObject","Ptr",hDC,"Ptr",hFont)

    ;-- Get string size
    VarSetCapacity(SIZE,8,0)
    RC:=DllCall("GetTextExtentPoint32"
        ,"Ptr",hDC                                      ;-- hDC
        ,"Str",p_String                                 ;-- lpString
        ,"Int",StrLen(p_String)                         ;-- c (string length)
        ,"Ptr",&SIZE)                                   ;-- lpSize

    ;-- Release the objects needed by the GetTextExtentPoint32 function
    DllCall("SelectObject","Ptr",hDC,"Ptr",old_hFont)
        ;-- Necessary to avoid memory leak

    DllCall("ReleaseDC","Ptr",HWND_DESKTOP,"Ptr",hDC)

    ;-- Return width
    result := RC ? NumGet(SIZE,0,"Int"):0
    Return result
}

Fnt_CreateFont(p_Name:="",p_Options:="") {
    Static Dummy34361446

          ;-- Misc. font constants
          ,LOGPIXELSY         := 90
          ,CLIP_DEFAULT_PRECIS:= 0
          ,DEFAULT_CHARSET    := 1
          ,DEFAULT_GUI_FONT   := 17
          ,OUT_TT_PRECIS      := 4

          ;-- Font family
          ,FF_DONTCARE  := 0x0
          ,FF_ROMAN     := 0x1
          ,FF_SWISS     := 0x2
          ,FF_MODERN    := 0x3
          ,FF_SCRIPT    := 0x4
          ,FF_DECORATIVE:= 0x5

          ;-- Font pitch
          ,DEFAULT_PITCH := 0
          ,FIXED_PITCH   := 1
          ,VARIABLE_PITCH:= 2

          ;-- Font quality
          ,DEFAULT_QUALITY       := 0
          ,DRAFT_QUALITY         := 1
          ,PROOF_QUALITY         := 2  ;-- AutoHotkey default
          ,NONANTIALIASED_QUALITY:= 3
          ,ANTIALIASED_QUALITY   := 4
          ,CLEARTYPE_QUALITY     := 5

          ;-- Font weight
          ,FW_DONTCARE:= 0
          ,FW_NORMAL  := 400
          ,FW_BOLD    := 700

    ;-- Parameters
    ;   Remove all leading/trailing white space
    p_Name    := Trim(p_Name," `f`n`r`t`v")
    p_Options := Trim(p_Options," `f`n`r`t`v")

    ;-- If both parameters are null or unspecified, return the handle to the
    ;   default GUI font.
    If (p_Name="" and p_Options="")
    {
       Result := DllCall("gdi32\GetStockObject","Int",DEFAULT_GUI_FONT)
       Return Result
    }

    ;-- Initialize options
    o_Height   := ""             ;-- Undefined
    o_Italic   := False
    o_Quality  := PROOF_QUALITY  ;-- AutoHotkey default
    o_Size     := ""             ;-- Undefined
    o_Strikeout:= False
    o_Underline:= False
    o_Weight   := FW_DONTCARE

    ;-- Extract options (if any) from p_Options
    Loop Parse,p_Options,%A_Space%
    {
        If A_LoopField is Space
            Continue

        If (SubStr(A_LoopField,1,4)="bold")
            o_Weight := FW_BOLD
        Else If (SubStr(A_LoopField,1,6)="italic")
            o_Italic := True
        Else If (SubStr(A_LoopField,1,4)="norm")
        {
            o_Italic    := False
            o_Strikeout := False
            o_Underline := False
            o_Weight    := FW_DONTCARE
        }
        Else If (A_LoopField="-s")
            o_Size := 0
        Else If (SubStr(A_LoopField,1,6)="strike")
            o_Strikeout := True
        Else If (SubStr(A_LoopField,1,9)="underline")
            o_Underline := True
        Else If (SubStr(A_LoopField,1,1)="h")
        {
            o_Height := SubStr(A_LoopField,2)
            o_Size := ""  ;-- Undefined
        }
        Else If (SubStr(A_LoopField,1,1)="q")
            o_Quality := SubStr(A_LoopField,2)
        Else If (SubStr(A_LoopField,1,1)="s")
        {
            o_Size := SubStr(A_LoopField,2)
            o_Height := ""  ;-- Undefined
        }
        Else If (SubStr(A_LoopField,1,1)="w")
            o_Weight := SubStr(A_LoopField,2)
    }

    ;-- Convert/Fix invalid or
    ;-- unspecified parameters/options
    If p_Name is Space
        p_Name := Fnt_GetFontName()   ;-- Font name of the default GUI font

    If o_Height is not Integer
       o_Height := ""                ;-- Undefined

    If o_Quality is not Integer
       o_Quality := PROOF_QUALITY    ;-- AutoHotkey default

    If o_Size is Space              ;-- Undefined
        o_Size := Fnt_GetFontSize()   ;-- Font size of the default GUI font
     Else
        If o_Size is not Integer
           o_Size := ""              ;-- Undefined
         Else
            If (o_Size=0)
               o_Size := ""          ;-- Undefined

    If o_Weight is not Integer
       o_Weight := FW_DONTCARE       ;-- A font with a default weight is created

    ;-- If needed, convert point size to em height
    If o_Height is Space        ;-- Undefined
        If o_Size is Integer    ;-- Allows for a negative size (emulates AutoHotkey)
        {
           hDC := DllCall("gdi32\CreateDCW","Str","DISPLAY","Ptr",0,"Ptr",0,"Ptr",0)
           o_Height := -Round(o_Size*DllCall("gdi32\GetDeviceCaps","Ptr",hDC,"Int",LOGPIXELSY)/72)
           DllCall("gdi32\DeleteDC","Ptr",hDC)
        }

    If o_Height is not Integer
       o_Height := 0             ;-- A font with a default height is created

    ;-- Create font
    hFont := DllCall("gdi32\CreateFontW"
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
    If !hFont         ;-- Zero or null
       Return 1

    Result := DllCall("gdi32\DeleteObject","Ptr",hFont) ? 1 : 0
    Return Result
}

Fnt_GetFontName(hFont:="") {
    Static Dummy87890484
          ,DEFAULT_GUI_FONT     := 17
          ,HWND_DESKTOP         := 0
          ,OBJ_FONT             := 6
          ,MAX_FONT_NAME_LENGTH := 32     ;-- In TCHARS

    ;-- If needed, get the handle to the default GUI font
    If (DllCall("gdi32\GetObjectType","Ptr",hFont)<>OBJ_FONT)
       hFont := DllCall("gdi32\GetStockObject","Int",DEFAULT_GUI_FONT)

    ;-- Select the font into the device context for the desktop
    hDC       := DllCall("user32\GetDC","Ptr",HWND_DESKTOP)
    old_hFont := DllCall("gdi32\SelectObject","Ptr",hDC,"Ptr",hFont)

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
          ,HWND_DESKTOP    := 0
          ,LOGPIXELSY      := 90
          ,DEFAULT_GUI_FONT:= 17
          ,OBJ_FONT        := 6

    ;-- If needed, get the handle to the default GUI font
    If (DllCall("gdi32\GetObjectType","Ptr",hFont)<>OBJ_FONT)
       hFont := DllCall("gdi32\GetStockObject","Int",DEFAULT_GUI_FONT)

    ;-- Select the font into the device context for the desktop
    hDC       := DllCall("user32\GetDC","Ptr",HWND_DESKTOP)
    old_hFont := DllCall("gdi32\SelectObject","Ptr",hDC,"Ptr",hFont)

    ;-- Collect the number of pixels per logical inch along the screen height
    l_LogPixelsY := DllCall("gdi32\GetDeviceCaps","Ptr",hDC,"Int",LOGPIXELSY)

    ;-- Get text metrics for the font
    VarSetCapacity(TEXTMETRIC,A_IsUnicode ? 60:56,0)
    DllCall("gdi32\GetTextMetricsW","Ptr",hDC,"Ptr",&TEXTMETRIC)

    ;-- Convert em height to point size
    l_Size := Round((NumGet(TEXTMETRIC,0,"Int")-NumGet(TEXTMETRIC,12,"Int"))*72/l_LogPixelsY)
           ;-- (Height - Internal Leading) * 72 / LogPixelsY

    ;-- Release the objects needed by the GetTextMetrics function
    DllCall("gdi32\SelectObject","Ptr",hDC,"Ptr",old_hFont)
           ;-- Necessary to avoid memory leak

    DllCall("user32\ReleaseDC","Ptr",HWND_DESKTOP,"Ptr",hDC)
    Return l_Size
}

;; // end of Font Library 3.0 functions

cleanup:
  Gui, demoGuia: Destroy
  Fnt_DeleteFont(hFontDummy)
  ExitApp
Return
