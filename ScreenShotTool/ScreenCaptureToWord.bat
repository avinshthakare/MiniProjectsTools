<# :batch_launcher
@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((Get-Content -LiteralPath '%~f0') -join [Environment]::NewLine)"
exit /b
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Win32 APIs for Hotkeys, Window Management, Cursor Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CURSORINFO {
        public Int32 cbSize;
        public Int32 flags;
        public IntPtr hCursor;
        public POINT ptScreenPos;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
        public Int32 x;
        public Int32 y;
    }

    public const Int32 CURSOR_SHOWING = 0x00000001;

    [DllImport("user32.dll")]
    public static extern bool GetCursorInfo(out CURSORINFO pci);

    [DllImport("user32.dll")]
    public static extern bool DrawIcon(IntPtr hDC, int X, int Y, IntPtr hIcon);
}
"@

# Form GUI Setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "Doc Snapper Pro Studio"
$form.Size = New-Object System.Drawing.Size(320, 570)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false

# 1. Target Word Doc Selector
$lblWord = New-Object System.Windows.Forms.Label
$lblWord.Text = "Target Word Document:"
$lblWord.Location = New-Object System.Drawing.Point(12, 10)
$lblWord.Size = New-Object System.Drawing.Size(200, 16)

$cmbWordDocs = New-Object System.Windows.Forms.ComboBox
$cmbWordDocs.Location = New-Object System.Drawing.Point(12, 28)
$cmbWordDocs.Size = New-Object System.Drawing.Size(220, 24)
$cmbWordDocs.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

$btnRefreshDocs = New-Object System.Windows.Forms.Button
$btnRefreshDocs.Text = "R"
$btnRefreshDocs.Location = New-Object System.Drawing.Point(238, 27)
$btnRefreshDocs.Size = New-Object System.Drawing.Size(26, 25)

# 2. Multi-Monitor Selector
$lblMonitor = New-Object System.Windows.Forms.Label
$lblMonitor.Text = "Display Monitor:"
$lblMonitor.Location = New-Object System.Drawing.Point(12, 58)
$lblMonitor.Size = New-Object System.Drawing.Size(200, 16)

$cmbMonitors = New-Object System.Windows.Forms.ComboBox
$cmbMonitors.Location = New-Object System.Drawing.Point(12, 75)
$cmbMonitors.Size = New-Object System.Drawing.Size(252, 24)
$cmbMonitors.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

for ($i = 0; $i -lt [System.Windows.Forms.Screen]::AllScreens.Count; $i++) {
    $s = [System.Windows.Forms.Screen]::AllScreens[$i]
    [void]$cmbMonitors.Items.Add("Monitor $($i + 1) ($($s.Bounds.Width)x$($s.Bounds.Height))")
}
if ($cmbMonitors.Items.Count -gt 0) { $cmbMonitors.SelectedIndex = 0 }

# 3. Mode Selection Group
$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = "Capture Area"
$grpMode.Location = New-Object System.Drawing.Point(12, 105)
$grpMode.Size = New-Object System.Drawing.Size(252, 90)

$rbFull = New-Object System.Windows.Forms.RadioButton
$rbFull.Text = "Selected Monitor (Ctrl+Shift+S)"
$rbFull.Location = New-Object System.Drawing.Point(10, 18)
$rbFull.Size = New-Object System.Drawing.Size(230, 20)
$rbFull.Checked = $true

$rbWindow = New-Object System.Windows.Forms.RadioButton
$rbWindow.Text = "Active Window (Ctrl+Shift+W)"
$rbWindow.Location = New-Object System.Drawing.Point(10, 40)
$rbWindow.Size = New-Object System.Drawing.Size(230, 20)

$rbSnip = New-Object System.Windows.Forms.RadioButton
$rbSnip.Text = "Drag Region (Ctrl+Shift+R)"
$rbSnip.Location = New-Object System.Drawing.Point(10, 62)
$rbSnip.Size = New-Object System.Drawing.Size(230, 20)

$grpMode.Controls.AddRange(@($rbFull, $rbWindow, $rbSnip))

# 4. Annotation Options
$chkDrawCursor = New-Object System.Windows.Forms.CheckBox
$chkDrawCursor.Text = "Draw Mouse Cursor"
$chkDrawCursor.Location = New-Object System.Drawing.Point(16, 200)
$chkDrawCursor.Size = New-Object System.Drawing.Size(135, 20)
$chkDrawCursor.Checked = $true

$chkHighlight = New-Object System.Windows.Forms.CheckBox
$chkHighlight.Text = "Highlight Cursor (Red Ring)"
$chkHighlight.Location = New-Object System.Drawing.Point(16, 222)
$chkHighlight.Size = New-Object System.Drawing.Size(200, 20)
$chkHighlight.Checked = $true

# 5. Primary Actions
$btnCapture = New-Object System.Windows.Forms.Button
$btnCapture.Text = "Snap && Paste"
$btnCapture.Location = New-Object System.Drawing.Point(12, 248)
$btnCapture.Size = New-Object System.Drawing.Size(252, 42)
$btnCapture.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$btnUndo = New-Object System.Windows.Forms.Button
$btnUndo.Text = "Undo Last Paste"
$btnUndo.Location = New-Object System.Drawing.Point(12, 296)
$btnUndo.Size = New-Object System.Drawing.Size(122, 30)

$btnBurst = New-Object System.Windows.Forms.Button
$btnBurst.Text = "Burst Mode: OFF"
$btnBurst.Location = New-Object System.Drawing.Point(140, 296)
$btnBurst.Size = New-Object System.Drawing.Size(124, 30)

$btnSaveDoc = New-Object System.Windows.Forms.Button
$btnSaveDoc.Text = "Save Word Doc"
$btnSaveDoc.Location = New-Object System.Drawing.Point(12, 332)
$btnSaveDoc.Size = New-Object System.Drawing.Size(122, 30)

$btnExportPdf = New-Object System.Windows.Forms.Button
$btnExportPdf.Text = "Export to PDF"
$btnExportPdf.Location = New-Object System.Drawing.Point(140, 332)
$btnExportPdf.Size = New-Object System.Drawing.Size(124, 30)

$btnMinTray = New-Object System.Windows.Forms.Button
$btnMinTray.Text = "Minimize to System Tray"
$btnMinTray.Location = New-Object System.Drawing.Point(12, 368)
$btnMinTray.Size = New-Object System.Drawing.Size(252, 30)

# Status Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.Location = New-Object System.Drawing.Point(12, 408)
$lblStatus.Size = New-Object System.Drawing.Size(252, 30)
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)

# Burst Mode Timer
$burstTimer = New-Object System.Windows.Forms.Timer
$burstTimer.Interval = 5000 # 5 seconds

# Tray Icon Setup
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = [System.Drawing.SystemIcons]::Application
$trayIcon.Text = "Doc Snapper Pro Studio"
$trayIcon.Visible = $false
$trayIcon.Add_DoubleClick({
    $form.Show()
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $trayIcon.Visible = $false
})

# Word Helpers
function Refresh-WordDocs {
    $cmbWordDocs.Items.Clear()
    try {
        $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        foreach ($d in $word.Documents) {
            [void]$cmbWordDocs.Items.Add($d.Name)
        }
    } catch {}

    if ($cmbWordDocs.Items.Count -eq 0) {
        [void]$cmbWordDocs.Items.Add("(New Document)")
    }
    $cmbWordDocs.SelectedIndex = 0
}

function Get-TargetDoc {
    try {
        $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    } catch {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
    }
    if ($word.Documents.Count -eq 0) {
        $word.Documents.Add() | Out-Null
    }

    $selectedName = $cmbWordDocs.SelectedItem
    if ($selectedName -and $selectedName -ne "(New Document)") {
        try {
            return $word.Documents.Item($selectedName)
        } catch {
            return $word.ActiveDocument
        }
    }
    return $word.ActiveDocument
}

# Drag-to-Snip Utility
function Get-SnippedBounds {
    $snipForm = New-Object System.Windows.Forms.Form
    $snipForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $snipForm.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
    $snipForm.BackColor = [System.Drawing.Color]::White
    $snipForm.Opacity = 0.25
    $snipForm.Cursor = [System.Windows.Forms.Cursors]::Cross
    $snipForm.TopMost = $true

    $rect = [System.Drawing.Rectangle]::Empty
    $startPt = [System.Drawing.Point]::Empty
    $isDown = $false

    $snipForm.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $script:startPt = $e.Location
            $script:isDown = $true
        }
    })

    $snipForm.Add_MouseMove({
        param($s, $e)
        if ($script:isDown) {
            $x = [Math]::Min($script:startPt.X, $e.X)
            $y = [Math]::Min($script:startPt.Y, $e.Y)
            $w = [Math]::Abs($script:startPt.X - $e.X)
            $h = [Math]::Abs($script:startPt.Y - $e.Y)
            $script:rect = New-Object System.Drawing.Rectangle($x, $y, $w, $h)
            $snipForm.Invalidate()
        }
    })

    $snipForm.Add_Paint({
        param($s, $e)
        if ($script:rect.Width -gt 0 -and $script:rect.Height -gt 0) {
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Red, 2)
            $e.Graphics.DrawRectangle($pen, $script:rect)
            $pen.Dispose()
        }
    })

    $snipForm.Add_MouseUp({
        param($s, $e)
        $script:isDown = $false
        $snipForm.Close()
    })

    [void]$snipForm.ShowDialog()
    return $script:rect
}

# Main Screen-Capture Engine
$actionExecute = {
    param([string]$ForcedMode = "")

    $form.Opacity = 0
    [System.Threading.Thread]::Sleep(200)

    $activeHwnd = [Win32]::GetForegroundWindow()
    $selectedMode = "FULL"
    if ($ForcedMode -ne "") {
        $selectedMode = $ForcedMode
    } elseif ($rbWindow.Checked) {
        $selectedMode = "WINDOW"
    } elseif ($rbSnip.Checked) {
        $selectedMode = "SNIP"
    }

    # Resolve area based on screen selection
    $targetScreen = [System.Windows.Forms.Screen]::AllScreens[0]
    if ($cmbMonitors.SelectedIndex -ge 0 -and $cmbMonitors.SelectedIndex -lt [System.Windows.Forms.Screen]::AllScreens.Count) {
        $targetScreen = [System.Windows.Forms.Screen]::AllScreens[$cmbMonitors.SelectedIndex]
    }

    $bounds = [System.Drawing.Rectangle]::Empty
    if ($selectedMode -eq "WINDOW") {
        $rect = New-Object Win32+RECT
        [Win32]::GetWindowRect($activeHwnd, [ref]$rect)
        $w = [Math]::Max(1, $rect.Right - $rect.Left)
        $h = [Math]::Max(1, $rect.Bottom - $rect.Top)
        $bounds = New-Object System.Drawing.Rectangle($rect.Left, $rect.Top, $w, $h)
    } elseif ($selectedMode -eq "SNIP") {
        $bounds = Get-SnippedBounds
    } else {
        $bounds = $targetScreen.Bounds
    }

    if ($bounds.Width -le 5 -or $bounds.Height -le 5) {
        $form.Opacity = 1
        return
    }

    # Capture raw screen
    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

    # Visual Annotations (Draw Cursor & Red-Ring Highlighter)
    if ($chkDrawCursor.Checked) {
        $cursorInfo = New-Object Win32+CURSORINFO
        $cursorInfo.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($cursorInfo)
        if ([Win32]::GetCursorInfo([ref]$cursorInfo) -and ($cursorInfo.flags -eq [Win32]::CURSOR_SHOWING)) {
            $curX = $cursorInfo.ptScreenPos.x - $bounds.Left
            $curY = $cursorInfo.ptScreenPos.y - $bounds.Top

            # Check if cursor is within screenshot bounding box
            if ($curX -ge 0 -and $curX -le $bounds.Width -and $curY -ge 0 -and $curY -le $bounds.Height) {
                if ($chkHighlight.Checked) {
                    $highlightBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60, 255, 0, 0))
                    $highlightPen = New-Object System.Drawing.Pen([System.Drawing.Color]::Red, 2)
                    $ringRadius = 18
                    $graphics.FillEllipse($highlightBrush, ($curX - $ringRadius), ($curY - $ringRadius), ($ringRadius * 2), ($ringRadius * 2))
                    $graphics.DrawEllipse($highlightPen, ($curX - $ringRadius), ($curY - $ringRadius), ($ringRadius * 2), ($ringRadius * 2))
                    $highlightBrush.Dispose()
                    $highlightPen.Dispose()
                }
                $hdc = $graphics.GetHdc()
                [Win32]::DrawIcon($hdc, $curX, $curY, $cursorInfo.hCursor)
                $graphics.ReleaseHdc($hdc)
            }
        }
    }

    [System.Windows.Forms.Clipboard]::SetImage($bitmap)
    $graphics.Dispose()
    $bitmap.Dispose()

    # Append to target Word document
    try {
        $doc = Get-TargetDoc
        $range = $doc.Content
        $range.Collapse(0) # Collapse to end
        $countBefore = $doc.InlineShapes.Count
        $range.Paste()

        if ($doc.InlineShapes.Count -gt $countBefore) {
            $shape = $doc.InlineShapes.Item($doc.InlineShapes.Count)
            if ($shape.Width -gt 450) {
                $aspect = $shape.Height / $shape.Width
                $shape.Width = 450
                $shape.Height = 450 * $aspect
            }
            $shape.Line.Visible = -1
            $shape.Line.ForeColor.RGB = 0xCCCCCC
            $shape.Line.Weight = 1.0
            $shape.Range.ParagraphFormat.Alignment = 1
        }
        $doc.Content.InsertParagraphAfter()
        [System.Media.SystemSounds]::Asterisk.Play()
        $lblStatus.Text = "Pasted to: $($doc.Name)"
    } catch {
        $lblStatus.Text = "Word error: $($_.Exception.Message)"
    }

    $form.Opacity = 1
}

# Attach Dashboard Events
$btnCapture.Add_Click({ & $actionExecute })
$btnRefreshDocs.Add_Click({ Refresh-WordDocs })

# Undo Last Paste Button
$btnUndo.Add_Click({
    try {
        $doc = Get-TargetDoc
        if ($doc.InlineShapes.Count -gt 0) {
            $lastShape = $doc.InlineShapes.Item($doc.InlineShapes.Count)
            $lastShape.Delete()
            $lblStatus.Text = "Last image undone."
            [System.Media.SystemSounds]::Beep.Play()
        } else {
            $lblStatus.Text = "No shapes to undo."
        }
    } catch {
        $lblStatus.Text = "Undo error: $($_.Exception.Message)"
    }
})

# Burst Mode Toggle
$btnBurst.Add_Click({
    if ($burstTimer.Enabled) {
        $burstTimer.Stop()
        $btnBurst.Text = "Burst Mode: OFF"
        $lblStatus.Text = "Burst Mode Stopped"
    } else {
        $burstTimer.Start()
        $btnBurst.Text = "Burst Mode: ON"
        $lblStatus.Text = "Burst active (every 5s)"
    }
})
$burstTimer.Add_Tick({ & $actionExecute })

# Save Active Doc
$btnSaveDoc.Add_Click({
    try {
        $doc = Get-TargetDoc
        $doc.Save()
        $lblStatus.Text = "Saved: $($doc.Name)"
        [System.Media.SystemSounds]::Beep.Play()
    } catch {
        $lblStatus.Text = "Could not save document."
    }
})

# Export to PDF
$btnExportPdf.Add_Click({
    try {
        $doc = Get-TargetDoc
        $docPath = $doc.FullName
        if ([string]::IsNullOrWhiteSpace($docPath) -or -not (Test-Path $docPath)) {
            $lblStatus.Text = "Save the doc to disk first!"
            return
        }
        $pdfPath = [System.IO.Path]::ChangeExtension($docPath, ".pdf")
        # 17 represents wdExportFormatPDF
        $doc.ExportAsFixedFormat($pdfPath, 17)
        $lblStatus.Text = "Exported to PDF!"
        [System.Media.SystemSounds]::Asterisk.Play()
    } catch {
        $lblStatus.Text = "PDF export failed."
    }
})

# Minimize to Tray
$btnMinTray.Add_Click({
    $form.Hide()
    $trayIcon.Visible = $true
    $trayIcon.ShowBalloonTip(2000, "Doc Snapper Active", "Running in background. Shortcuts still work!", [System.Windows.Forms.ToolTipIcon]::Info)
})

# Global Shortcuts Registration (Ctrl + Shift = 0x0006)
$MOD = 0x0006
$VK_S = 0x53; $VK_W = 0x57; $VK_R = 0x52
$ID_FULL = 101; $ID_WIN = 102; $ID_SNIP = 103

$form.Add_Load({
    [Win32]::RegisterHotKey($form.Handle, $ID_FULL, $MOD, $VK_S) | Out-Null
    [Win32]::RegisterHotKey($form.Handle, $ID_WIN,  $MOD, $VK_W) | Out-Null
    [Win32]::RegisterHotKey($form.Handle, $ID_SNIP, $MOD, $VK_R) | Out-Null
    Refresh-WordDocs
})

$form.Add_FormClosing({
    $burstTimer.Stop()
    [Win32]::UnregisterHotKey($form.Handle, $ID_FULL) | Out-Null
    [Win32]::UnregisterHotKey($form.Handle, $ID_WIN)  | Out-Null
    [Win32]::UnregisterHotKey($form.Handle, $ID_SNIP) | Out-Null
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
})

# Global Windows Message Interceptor
Add-Type -ReferencedAssemblies "System.Windows.Forms" @"
using System;
using System.Windows.Forms;

public class GlobalHotkeyFilter : IMessageFilter {
    private Action<int> _callback;
    public GlobalHotkeyFilter(Action<int> callback) { _callback = callback; }
    public bool PreFilterMessage(ref Message m) {
        const int WM_HOTKEY = 0x0312;
        if (m.Msg == WM_HOTKEY) {
            int id = m.WParam.ToInt32();
            if (_callback != null) { _callback(id); }
            return true;
        }
        return false;
    }
}
"@

$hotkeyCallback = [Action[int]]{
    param($id)
    if ($id -eq $ID_FULL) { & $actionExecute "FULL" }
    if ($id -eq $ID_WIN)  { & $actionExecute "WINDOW" }
    if ($id -eq $ID_SNIP) { & $actionExecute "SNIP" }
}

$msgFilter = [GlobalHotkeyFilter]::new($hotkeyCallback)
[System.Windows.Forms.Application]::AddMessageFilter($msgFilter)

# Render Form
$form.Controls.AddRange(@(
    $lblWord, $cmbWordDocs, $btnRefreshDocs,
    $lblMonitor, $cmbMonitors,
    $grpMode,
    $chkDrawCursor, $chkHighlight,
    $btnCapture, $btnUndo, $btnBurst,
    $btnSaveDoc, $btnExportPdf, $btnMinTray,
    $lblStatus
))

[System.Windows.Forms.Application]::Run($form)

# Cleanup
[System.Windows.Forms.Application]::RemoveMessageFilter($msgFilter)
