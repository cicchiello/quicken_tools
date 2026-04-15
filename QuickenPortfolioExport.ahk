#Requires AutoHotkey v2.0
#SingleInstance Force

qdfPath := "C:\tmp\HOME_nightly.QDF"
exportPath := "\\10.0.0.214\pi-nas\openclaw\quicken_tools\portfolio_nightly.csv"

; Exit codes:
;   0  = success
;   10 = Quicken already running
;   11 = QDF file missing
;   12 = Quicken window did not appear
;   13 = Quicken window could not be activated

if ProcessExist("qw.exe") {
    ExitApp 10
}

if !FileExist(qdfPath) {
    ExitApp 11
}

Run qdfPath
Sleep 25000

if !WinExist("ahk_exe qw.exe") {
    ExitApp 12
}

WinActivate "ahk_exe qw.exe"
Sleep 1000

if !WinActive("ahk_exe qw.exe") {
    ExitApp 13
}

Send "^u"
Sleep 3000

Send "!s"
Sleep 500
Send "{Home}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Down}"
Sleep 300
Send "{Enter}"
Sleep 1500

Send "!g"
Sleep 500
Send "{Home}"
Sleep 300
Send "{Enter}"
Sleep 2000

Send "^p"
Sleep 2000

Send "!x"
Sleep 500
Send "{Space}"
Sleep 1000

Send "{Enter}"
Sleep 3000

SendText exportPath
Sleep 500
Send "{Enter}"
Sleep 3000

; Approve overwrite
Send "!y"
Sleep 3000

; Close Quicken
WinActivate "ahk_exe qw.exe"
Sleep 500
Send "!{F4}"
Sleep 3000

ExitApp 0
