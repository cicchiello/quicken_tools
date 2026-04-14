# OpenClaw Quicken Portfolio Export Automation

Automates a daily export of the Quicken **Investing Portfolio** view to a CSV file for downstream OpenClaw processing.

## What this does

This automation:

1. Aborts immediately if Quicken is already running.
2. Opens a specific Quicken `.QDF` data file.
3. Waits for Quicken to finish launching.
4. Opens the **Investing Portfolio** view.
5. Sets:
   - **Show** = `Holdings`
   - **Group By** = `Accounts`
6. Opens the export workflow.
7. Selects the **Export to:** option.
8. Saves the export to a fixed CSV path.
9. Approves overwrite of the target file.
10. Closes Quicken.

The intended output file is:

```text
\\10.0.0.214\pi-nas\openclaw\quicken_exports\portfolio_nightly.csv
```

---

## Current assumptions

This automation currently assumes all of the following are true:

- It runs on a Windows PC that is always on.
- The Windows user account can launch and use Quicken.
- Auto-login is enabled or the user is already logged in.
- Quicken is installed at:

```text
C:\Program Files (x86)\Quicken\qw.exe
```

- The Quicken data file is known and fixed.
- The export destination UNC path is reachable from the PC.
- Quicken UI keyboard shortcuts and dropdown ordering remain stable.

Because this is UI automation, layout or workflow changes in Quicken may require script updates.

---

## Repository purpose

This repo is intended to hold:

- the AutoHotkey automation script
- any helper scripts or wrappers
- Task Scheduler setup notes
- logs or future enhancements for monitoring and error handling

---

## Prerequisites

### Software

- Windows 10 or 11
- Quicken for Windows
- AutoHotkey v2

### Access

- Read/write access to the Quicken `.QDF` file
- Write access to the export target directory on the NAS/share

### Recommended environment

- Dedicated Windows account for automation
- Auto-login enabled for unattended execution
- Task Scheduler configured to run only when the user is logged on
- Sleep/hibernation disabled on the automation PC

---

## Script behavior

The automation is currently implemented in AutoHotkey v2.

### Guard condition

If Quicken is already open when the script starts, the script exits immediately without taking any action.

Example guard:

```ahk
if ProcessExist("qw.exe") {
    ExitApp
}
```

This avoids interfering with an interactive Quicken session.

### Launch method

The automation launches the **specific `.QDF` file directly**, instead of launching `qw.exe` and letting Quicken choose the most recent file.

This has proven to be more reliable.

---

## Current script

> Update the `qdfPath` value for the local machine.

```ahk
#Requires AutoHotkey v2.0
#SingleInstance Force

qdfPath := "C:\path\to\your\file.QDF"
exportPath := "\\10.0.0.214\pi-nas\openclaw\quicken_exports\portfolio_nightly.csv"

if ProcessExist("qw.exe") {
    ExitApp
}

if !FileExist(qdfPath) {
    MsgBox "Quicken data file not found:`n" qdfPath
    ExitApp
}

Run qdfPath
Sleep 20000

if !WinExist("ahk_exe qw.exe") {
    MsgBox "Quicken process/window not found after 20 seconds."
    ExitApp
}

WinActivate "ahk_exe qw.exe"
Sleep 1000

; Open Investing Portfolio view
Send "^u"
Sleep 3000

; Move to "Show" and set it to Holdings
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

; Move to "Group By" and set it to Accounts
Send "!g"
Sleep 500
Send "{Home}"
Sleep 300
Send "{Enter}"
Sleep 2000

; Open export dialog
Send "^p"
Sleep 2000

; Select the "Export to:" option
Send "!x"
Sleep 500
Send "{Space}"
Sleep 1000

; Activate Export
Send "{Enter}"
Sleep 3000

; Save to target CSV path
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
```

---

## Configuration

These are the main values you are likely to customize:

### `qdfPath`

Local path to the Quicken data file to be opened.

Example:

```ahk
qdfPath := "C:\Finance\Quicken\PrimaryData.QDF"
```

### `exportPath`

Destination CSV file path.

Example:

```ahk
exportPath := "\\10.0.0.214\pi-nas\openclaw\quicken_exports\portfolio_nightly.csv"
```

### Timing values

The script currently uses fixed delays such as:

- `Sleep 20000` after launching Quicken
- `Sleep 3000` after opening views/dialogs

These may need adjustment depending on machine speed, Quicken updates, or network conditions.

---

## Running manually

To test manually:

1. Log into the Windows account used for automation.
2. Ensure Quicken is not already open.
3. Double-click the `.ahk` file.
4. Confirm that:
   - the correct data file opens
   - the portfolio view is selected
   - the export is written to the expected path
   - overwrite is handled
   - Quicken closes when done

---

## Running from Task Scheduler

Recommended Task Scheduler settings:

- **Trigger:** Daily at the desired time
- **Action:** Start AutoHotkey with the script as an argument, or run the `.ahk` file directly if associated
- **Run only when user is logged on:** Yes
- **Start in:** the script directory

Typical explicit configuration:

- **Program/script:**

```text
C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe
```

- **Add arguments:**

```text
"C:\path\to\quicken_export.ahk"
```

Do at least one manual **Run** from Task Scheduler before relying on the daily trigger.

---

## Recommended hardening

The current version works, but the next improvements worth adding are:

### 1. Logging

Add a log file for:

- script start time
- abort because Quicken is already running
- missing QDF file
- inability to find the Quicken window
- successful export completion

Example:

```ahk
logFile := "C:\OpenClaw\logs\quicken_export.log"
FileAppend(FormatTime(, "yyyy-MM-dd HH:mm:ss") " - Started`n", logFile)
```

### 2. Post-export validation

Optionally confirm that:

- the CSV file exists
- the modification time changed
- the file size is above some minimum threshold

### 3. Error handling for missing share access

If the NAS/share is unavailable, the Save step may fail. It would be useful to detect that and log it explicitly.

### 4. Safer window targeting

Today the script relies on keyboard navigation and fixed timing. Future versions could use more precise window/control targeting if needed.

---

## Known fragility points

This is UI automation, so it can break if any of the following change:

- Quicken keyboard shortcuts
- the order of items in the **Show** dropdown
- the default position in the **Group By** dropdown
- the export dialog layout
- overwrite confirmation wording or accelerator keys
- launch timing

After Quicken updates, rerun a manual test.

---

## Troubleshooting

### Quicken opens the wrong file

Make sure the script launches the `.QDF` file directly:

```ahk
Run qdfPath
```

Do not rely on launching `qw.exe` alone.

### Script starts while Quicken is already open

Confirm the guard is present:

```ahk
if ProcessExist("qw.exe") {
    ExitApp
}
```

### Export path fails

Verify that the Windows account running the script can browse and write to:

```text
\\10.0.0.214\pi-nas\openclaw\quicken_exports\
```

### The wrong dropdown item is selected

Quicken may have changed the item ordering. Re-test manually and update the number of `{Down}` key presses.

### Task Scheduler run behaves differently than manual run

Common causes:

- user not logged in
- network share not yet available
- Quicken startup slower under scheduled launch
- script working directory not set

---

## Suggested repo contents

A simple starting layout:

```text
.
├── README.md
├── scripts/
│   └── quicken_export.ahk
├── docs/
│   └── task_scheduler_notes.md
└── logs/
```

---

## Status

Current status: **working manually end-to-end**.

Validated behaviors:

- specific QDF file launch works
- portfolio export works
- overwrite approval works
- Quicken closes after export
- existing-running-Quicken guard can be added cleanly

Remaining operational validation:

- unattended reboot + auto-login test
- Task Scheduler invocation test
- repeated daily reliability test

