# Doc Snapper Pro Studio

A lightweight, zero-dependency Windows automation tool packaged inside a single hybrid `.bat` file. It provides an always-on-top dashboard and system-wide hotkeys to capture screens, windows, or custom snipped regions, and immediately pastes them sequentially into Microsoft Word documents without saving temporary image files to disk.

---

## Key Features

* **Direct-to-Word Memory Streaming:** Captures are processed in memory and pasted directly into Microsoft Word documents via the clipboard, leaving no leftover image files on your drive.
* **Global Hotkey Interceptors:** Triggers screen captures from any active program or browser without needing to click the tool window.
* **Document Targeting:** Automatically detects all open Microsoft Word documents. Select your target document from a dropdown menu, or let the tool open a new document automatically.
* **Visual Pointer & Highlighter:** Captures the hardware mouse cursor and optionally draws a transparent red focus ring around the pointer tip to highlight clicked or hovered UI elements.
* **Multi-Monitor Awareness:** Automatically enumerates connected displays and lets you target primary, secondary, or tertiary displays.
* **Burst Mode Automation:** Snaps and pastes on a repeating 5-second loop for hands-free documentation of long processes.
* **One-Click Undo:** Remove the most recently pasted screenshot from the Word document without switching windows.
* **Direct PDF Export:** Converts and saves the active Word document directly to PDF format alongside the original file.
* **System Tray Mode:** Minimizes cleanly into the Windows Taskbar notification area (System Tray) while keeping all global hotkeys fully active.

---

## Keyboard Shortcuts

The following shortcuts work globally across Windows while the tool is running (even when minimized to the system tray):

| Shortcut | Mode | Description |
| --- | --- | --- |
| `Ctrl` + `Shift` + `S` | **Selected Monitor** | Captures the entire monitor chosen in the display selector. |
| `Ctrl` + `Shift` + `W` | **Active Window** | Captures only the focused window or application in the foreground. |
| `Ctrl` + `Shift` + `R` | **Drag Region** | Opens a transparent crosshair overlay to click-and-drag a custom bounding box. |

---

## System Requirements

* **Operating System:** Windows 10 or Windows 11 (64-bit / 32-bit).
* **Word Processor:** Microsoft Word (Desktop Edition: Office 2013, 2016, 2019, 2021, or Microsoft 365).
* **Runtime:** Built-in Windows PowerShell 5.1+ and .NET Framework 4.5+ (pre-installed on all modern Windows versions).
* **Administrative Privileges:** Not required. Runs entirely in user space.

---

## How to Set Up & Run

1. **Save the File:** Ensure the script code is saved as a batch file (for example, `Doc Snapper Pro Studio.bat`).
2. **Open Word (Optional):** You can open an existing Word document beforehand, or let the tool launch a fresh one automatically.
3. **Launch the Tool:** Double-click `Doc Snapper Pro Studio.bat`.
* A command prompt launcher will flash briefly, followed by the **Doc Snapper Pro Studio** dashboard window.


4. **Targeting Documents:**
* If you have multiple Word files open, click the **Target Word Document** drop-down to route screenshots to a specific document.
* Click **`R`** at any time to refresh the list if you opened or closed documents while the tool was running.



---

## Dashboard Reference

* **Target Word Document & [R]:** Dropdown list of running Word documents. The `R` button updates the list.
* **Display Monitor:** Selects which screen to snapshot when running in **Full Screen / Selected Monitor** mode.
* **Capture Area Group:** Sets the behavior used when clicking the **Snap & Paste** dashboard button.
* **Draw Mouse Cursor:** Includes the actual Windows pointer icon in the captured screenshot.
* **Highlight Cursor (Red Ring):** Draws an accent halo around the pointer tip to emphasize focus.
* **Snap & Paste:** Executes the capture action based on current settings and appends it to Word.
* **Undo Last Paste:** Deletes the most recent inline image inserted into the selected Word document.
* **Burst Mode (ON / OFF):** Toggles an automated 5-second repeating capture loop.
* **Save Word Doc:** Triggers a native document save on the targeted Word file.
* **Export to PDF:** Automatically compiles and exports the active document into a `.pdf` file in the same directory where the `.docx` is saved.
* **Minimize to System Tray:** Hides the dashboard interface from the screen and taskbar. Double-click the tray icon near the Windows system clock to bring it back.

---

## Image Formatting Specifications

Each image inserted into Microsoft Word is automatically formatted according to these rules:

* **Placement:** Appended sequentially at the bottom (`wdCollapseEnd`) of the active document.
* **Width Constraint:** If the image width exceeds 450 points (~6.25 inches), it is proportionally scaled down so it stays within standard Letter/A4 margin boundaries without clipping.
* **Alignment:** Centered within the paragraph block.
* **Border Line:** A subtle, clean 1 pt border (`#CCCCCC`) is applied to provide contrast against light backgrounds.
* **Spacing:** A blank paragraph spacing block is inserted immediately after the shape.

---

## Troubleshooting

* **Status displays "Word is not running!":** Ensure Microsoft Word desktop application is installed and not blocked by background sandbox restrictions.
* **Global hotkey does not respond:** Another background application (e.g., streaming tools, GPU overlays, or other screen-capture software) may already have claimed the shortcut combination (`Ctrl + Shift + S/W/R`).
* **PDF Export Fails:** The Word document must be saved to your local disk at least once as a `.docx` file before it can be exported as a `.pdf`.
