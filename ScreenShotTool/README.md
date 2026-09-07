# DocSnapper Mini

A high-speed, zero-install screen capture assistant for Windows packaged in a standalone hybrid `.bat` file. DocSnapper Mini floats as a dark-mode mini-toolbar that captures screens, active application windows, or freeform snipped regions, automatically streaming the images directly into Microsoft Word via the clipboard—without saving intermediate image files to disk.

---

## Key Features

* **Minimal Floating Toolbar:** An ultra-compact dark toolbar (approx. 260 × 130 px) pinned to the bottom-right corner of your primary screen.
* **Direct-to-Memory Pasting:** Screenshots bypass hard disk writes entirely, copying directly to the system clipboard and inserting into Word's active document stream.
* **Smart Image Formatting:** Automatically fits wide captures to standard document margins (max width 450 pt), centers the image, adds a clean 1 pt gray outline border, and creates sequential line breaks.
* **Word Document Targeting:** Routes captures to any currently open Word document chosen from a live dropdown, or automatically creates a new document if Word is closed.
* **Multi-Monitor Aware:** Enumerates all attached displays, allowing you to choose which monitor to snap when in full-screen capture mode.
* **Hardware Mouse & Highlight Ring:** Optionally records the hardware mouse cursor in place, complete with an accent red halo to highlight clicked UI targets.
* **Hands-Free Burst Mode:** Automates continuous screenshot capture on a repeating 5-second loop.
* **Instant Undo:** Deletes the most recently pasted image shape from the target Word document without switching windows.
* **One-Click PDF Export:** Automatically compiles and exports the active document into a `.pdf` file alongside your `.docx` source.
* **System Tray Minimization:** Hides the toolbar to the Windows notification tray near the system clock while keeping global hotkeys active.

---

## Keyboard Shortcuts

The following hotkeys function system-wide, even when the toolbar is minimized to the system tray:

| Shortcut | Capture Type | Behavior |
| --- | --- | --- |
| `Ctrl` + `Shift` + `S` | **Selected Display** | Captures the full screen of the monitor selected in Options. |
| `Ctrl` + `Shift` + `W` | **Active Window** | Captures only the foreground window/application in focus. |
| `Ctrl` + `Shift` + `R` | **Region Snip** | Displays a transparent crosshair overlay to drag a custom bounding box. |

---

## System Requirements

* **Operating System:** Windows 10 or Windows 11.
* **Office Suite:** Microsoft Word desktop installation (Office 2013 through 2021, or Microsoft 365).
* **Dependencies:** None. Powered by native Windows PowerShell 5.1 and .NET Framework (pre-installed on Windows).
* **Permissions:** Runs in standard user space (Administrator rights not required).

---

## Installation & Setup

1. **Save Script:** Save the script code as `ScreenCaptureToWord.bat`.
2. **File Encoding:** Save the file with **ANSI** or **UTF-8 (without BOM)** encoding to ensure command prompt syntax parses cleanly.
3. **Launch:** Double-click `ScreenCaptureToWord.bat`.
* A command prompt window will initialize and disappear, followed by the floating mini-toolbar in the lower-right corner of your desktop.



---

## Interface Reference

### Main Toolbar

* **Snap and Paste:** Triggers an immediate screen capture based on the active monitor and appends it to Word.
* **OPT:** Opens the Options flyout dialog to adjust targeting and annotation behaviors.
* **Undo:** Deletes the most recent image shape from the active document.
* **Burst:** Toggles the automatic 5-second capture interval (turns red when active).
* **Save:** Sends a direct save command to the targeted Microsoft Word document.
* **PDF:** Exports the current Word document directly as a `.pdf` file in the same local folder.
* **_ (Minimize):** Sends the tool to the Windows system tray. Double-click the tray icon to restore.

### Options Dialog (`OPT`)

* **Target Word Doc:** Dropdown menu listing all running Word document files. Select the document where captures should be placed.
* **Screen Monitor:** Selects which monitor is captured when using the full-screen mode.
* **Capture Mouse Cursor:** Enables or disables drawing the pointer on the screenshot.
* **Draw Red Halo Ring:** Toggles the translucent red highlight circle centered around the pointer tip.

---

## Troubleshooting

* **Script window closes immediately on error:** Launch the script from an already open `cmd.exe` terminal window to review any local PowerShell execution policy blocks.
* **PDF Export Fails:** The target Word document must be saved to your local drive at least once as a `.docx` before Word can generate an export path.
* **Hotkeys not firing:** Confirm that another background application (e.g., Discord, GeForce Experience, AMD Radeon Software, or Snipping Tool) does not have exclusive locks on `Ctrl + Shift + S/W/R`.
