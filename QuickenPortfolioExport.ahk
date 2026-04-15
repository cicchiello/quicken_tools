#Requires AutoHotkey v2.0
#SingleInstance Force

quickenPath := "c:\tmp\HOME_nightly.QDF"
exportPath := "\\10.0.0.214\pi-nas\openclaw\quicken_tools\portfolio_nightly.csv"

if ProcessExist("qw.exe") {
    ExitApp
}

if !FileExist(quickenPath) {
    MsgBox "Quicken executable not found at:`n" quickenPath
    ExitApp
}

Run quickenPath
Sleep 25000

if !WinExist("ahk_exe qw.exe") {
    MsgBox "Quicken process/window not found after 20 seconds."
    ExitApp
}

WinActivate "ahk_exe qw.exe"
Sleep 1000

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

ExitApp
