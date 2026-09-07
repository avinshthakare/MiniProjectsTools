<# :batch_launcher
@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$code = [System.IO.File]::ReadAllText('%~f0'); iex $code"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Script encountered an error.
    pause
)
exit /b
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Win32 APIs for Hotkeys, Windows, and Mouse Cursor
$cSource = @'
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

public class GlobalHotkeyFilter : System.Windows.Forms.IMessageFilter {
    private Action<int> _callback;
    public GlobalHotkeyFilter(Action<int> callback) { _callback = callback; }
    public bool PreFilterMessage(ref System.Windows.Forms.Message m) {
        if (m.Msg == 0x0312) {
            if (_callback != null) { _callback(m.WParam.ToInt32()); }
            return true;
        }
        return false;
    }
}
'@

Add-Type -TypeDefinition $cSource -ReferencedAssemblies "System.Windows.Forms","System.Drawing"

# Global Variables
$global:SelectedMonitorIndex = 0
$global:DrawCursor = $true
$global:HighlightCursor = $true
$global:SelectedDocName = ""

# Main Compact Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "DocSnapper"
$form.Size = New-Object System.Drawing.Size(260, 130)
$form.StartPosition = "Manual"
$screenArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(($screenArea.Right - 280), ($screenArea.Bottom - 150))
$form.TopMost = $true
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$form.BackColor = [System.Drawing.Color]::FromArgb(35, 36, 38)
$form.ForeColor = [System.Drawing.Color]::White

function Set-BtnStyle($b, $bg, $fg) {
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderSize = 0
    $b.BackColor = $bg
    $b.ForeColor = $fg
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
}

# Row 1: Snap & Options
$btnCapture = New-Object System.Windows.Forms.Button
$btnCapture.Text = "Snap and Paste"
$btnCapture.Location = New-Object System.Drawing.Point(8, 8)
$btnCapture.Size = New-Object System.Drawing.Size(192, 34)
$btnCapture.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
Set-BtnStyle $btnCapture ([System.Drawing.Color]::FromArgb(26, 115, 232)) ([System.Drawing.Color]::White)

$btnSettings = New-Object System.Windows.Forms.Button
$btnSettings.Text = "OPT"
$btnSettings.Location = New-Object System.Drawing.Point(206, 8)
$btnSettings.Size = New-Object System.Drawing.Size(32, 34)
$btnSettings.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
Set-BtnStyle $btnSettings ([System.Drawing.Color]::FromArgb(60, 64, 67)) ([System.Drawing.Color]::White)

# Row 2: Actions
$btnUndo = New-Object System.Windows.Forms.Button
$btnUndo.Text = "Undo"
$btnUndo.Location = New-Object System.Drawing.Point(8, 48)
$btnUndo.Size = New-Object System.Drawing.Size(42, 26)
$btnUndo.Font = New-Object System.Drawing.Font("Segoe UI", 8)
Set-BtnStyle $btnUndo ([System.Drawing.Color]::FromArgb(60, 64, 67)) ([System.Drawing.Color]::White)

$btnBurst = New-Object System.Windows.Forms.Button
$btnBurst.Text = "Burst"
$btnBurst.Location = New-Object System.Drawing.Point(54, 48)
$btnBurst.Size = New-Object System.Drawing.Size(44, 26)
$btnBurst.Font = New-Object System.Drawing.Font("Segoe UI", 8)
Set-BtnStyle $btnBurst ([System.Drawing.Color]::FromArgb(60, 64, 67)) ([System.Drawing.Color]::White)

$btnSaveDoc = New-Object System.Windows.Forms.Button
$btnSaveDoc.Text = "Save"
$btnSaveDoc.Location = New-Object System.Drawing.Point(102, 48)
$btnSaveDoc.Size = New-Object System.Drawing.Size(42, 26)
$btnSaveDoc.Font = New-Object System.Drawing.Font("Segoe UI", 8)
Set-BtnStyle $btnSaveDoc ([System.Drawing.Color]::FromArgb(60, 64, 67)) ([System.Drawing.Color]::White)

$btnExportPdf = New-Object System.Windows.Forms.Button
$btnExportPdf.Text = "PDF"
$btnExportPdf.Location = New-Object System.Drawing.Point(148, 48)
$btnExportPdf.Size = New-Object System.Drawing.Size(42, 26)
$btnExportPdf.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
Set-BtnStyle $btnExportPdf ([System.Drawing.Color]::FromArgb(60, 64, 67)) ([System.Drawing.Color]::White)

$btnMinTray = New-Object System.Windows.Forms.Button
$btnMinTray.Text = "_"
$btnMinTray.Location = New-Object System.Drawing.Point(194, 48)
$btnMinTray.Size = New-Object System.Drawing.Size(44, 26)
$btnMinTray.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
Set-BtnStyle $btnMinTray ([System.Drawing.Color]::FromArgb(60, 64, 67)) ([System.Drawing.Color]::White)

# Row 3: Status
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready - Shortcuts active"
$lblStatus.Location = New-Object System.Drawing.Point(8, 78)
$lblStatus.Size = New-Object System.Drawing.Size(230, 16)
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(160, 165, 170)

# Burst Timer
$burstTimer = New-Object System.Windows.Forms.Timer
$burstTimer.Interval = 5000

# Tray Icon
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = [System.Drawing.SystemIcons]::Application
$trayIcon.Text = "DocSnapper Mini"
$trayIcon.Visible = $false
$trayIcon.Add_DoubleClick({
    $form.Show()
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $trayIcon.Visible = $false
})

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

    if ($global:SelectedDocName -and $global:SelectedDocName -ne "(New Document)") {
        try {
            return $word.Documents.Item($global:SelectedDocName)
        } catch {
            return $word.ActiveDocument
        }
    }
    return $word.ActiveDocument
}

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

# Main Capture Workflow
$actionExecute = {
    param([string]$Mode = "FULL")

    $form.Opacity = 0
    [System.Threading.Thread]::Sleep(150)

    $activeHwnd = [Win32]::GetForegroundWindow()

    $targetScreen = [System.Windows.Forms.Screen]::AllScreens[0]
    if ($global:SelectedMonitorIndex -ge 0 -and $global:SelectedMonitorIndex -lt [System.Windows.Forms.Screen]::AllScreens.Count) {
        $targetScreen = [System.Windows.Forms.Screen]::AllScreens[$global:SelectedMonitorIndex]
    }

    $bounds = [System.Drawing.Rectangle]::Empty
    if ($Mode -eq "WINDOW") {
        $rect = New-Object Win32+RECT
        [Win32]::GetWindowRect($activeHwnd, [ref]$rect)
        $w = [Math]::Max(1, $rect.Right - $rect.Left)
        $h = [Math]::Max(1, $rect.Bottom - $rect.Top)
        $bounds = New-Object System.Drawing.Rectangle($rect.Left, $rect.Top, $w, $h)
    } elseif ($Mode -eq "SNIP") {
        $bounds = Get-SnippedBounds
    } else {
        $bounds = $targetScreen.Bounds
    }

    if ($bounds.Width -le 5 -or $bounds.Height -le 5) {
        $form.Opacity = 1
        return
    }

    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

    # Draw Pointer
    if ($global:DrawCursor) {
        $cursorInfo = New-Object Win32+CURSORINFO
        $cursorInfo.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($cursorInfo)
        if ([Win32]::GetCursorInfo([ref]$cursorInfo) -and ($cursorInfo.flags -eq [Win32]::CURSOR_SHOWING)) {
            $curX = $cursorInfo.ptScreenPos.x - $bounds.Left
            $curY = $cursorInfo.ptScreenPos.y - $bounds.Top

            if ($curX -ge 0 -and $curX -le $bounds.Width -and $curY -ge 0 -and $curY -le $bounds.Height) {
                if ($global:HighlightCursor) {
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

    # Append to Word
    try {
        $doc = Get-TargetDoc
        $range = $doc.Content
        $range.Collapse(0)
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
        $lblStatus.Text = "Pasted: $(Get-Date -Format 'HH:mm:ss')"
    } catch {
        $lblStatus.Text = "Word Err: $($_.Exception.Message)"
    }

    $form.Opacity = 1
}

$btnCapture.Add_Click({ & $actionExecute "FULL" })

$btnUndo.Add_Click({
    try {
        $doc = Get-TargetDoc
        if ($doc.InlineShapes.Count -gt 0) {
            $lastShape = $doc.InlineShapes.Item($doc.InlineShapes.Count)
            $lastShape.Delete()
            $lblStatus.Text = "Image undone"
            [System.Media.SystemSounds]::Beep.Play()
        }
    } catch {}
})

$btnBurst.Add_Click({
    if ($burstTimer.Enabled) {
        $burstTimer.Stop()
        $btnBurst.BackColor = [System.Drawing.Color]::FromArgb(60, 64, 67)
        $lblStatus.Text = "Burst Mode: OFF"
    } else {
        $burstTimer.Start()
        $btnBurst.BackColor = [System.Drawing.Color]::FromArgb(234, 67, 53)
        $lblStatus.Text = "Burst: Active (5s)"
    }
})
$burstTimer.Add_Tick({ & $actionExecute "FULL" })

$btnSaveDoc.Add_Click({
    try {
        $doc = Get-TargetDoc
        $doc.Save()
        $lblStatus.Text = "Doc saved"
        [System.Media.SystemSounds]::Beep.Play()
    } catch {}
})

$btnExportPdf.Add_Click({
    try {
        $doc = Get-TargetDoc
        $docPath = $doc.FullName
        if ($docPath -and (Test-Path $docPath)) {
            $pdfPath = [System.IO.Path]::ChangeExtension($docPath, ".pdf")
            $doc.ExportAsFixedFormat($pdfPath, 17)
            $lblStatus.Text = "Exported to PDF"
            [System.Media.SystemSounds]::Asterisk.Play()
        } else {
            $lblStatus.Text = "Save doc first"
        }
    } catch {
        $lblStatus.Text = "PDF failed"
    }
})

$btnMinTray.Add_Click({
    $form.Hide()
    $trayIcon.Visible = $true
    $trayIcon.ShowBalloonTip(1500, "DocSnapper Active", "Running in background.", [System.Windows.Forms.ToolTipIcon]::Info)
})

# Settings Dialog
$btnSettings.Add_Click({
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "Options"
    $settingsForm.Size = New-Object System.Drawing.Size(250, 230)
    $settingsForm.StartPosition = "Manual"
    $settingsForm.Location = New-Object System.Drawing.Point($form.Left, ($form.Top - 240))
    $settingsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
    $settingsForm.TopMost = $true

    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Text = "Target Word Doc:"
    $lblD.Location = New-Object System.Drawing.Point(12, 10)
    $lblD.Size = New-Object System.Drawing.Size(210, 16)
    
    $cmbD = New-Object System.Windows.Forms.ComboBox
    $cmbD.Location = New-Object System.Drawing.Point(12, 28)
    $cmbD.Size = New-Object System.Drawing.Size(210, 22)
    $cmbD.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    try {
        $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        foreach ($d in $word.Documents) { [void]$cmbD.Items.Add($d.Name) }
    } catch {}
    if ($cmbD.Items.Count -eq 0) { [void]$cmbD.Items.Add("(New Document)") }
    
    if ($global:SelectedDocName -and $cmbD.Items.Contains($global:SelectedDocName)) {
        $cmbD.SelectedItem = $global:SelectedDocName
    } else {
        $cmbD.SelectedIndex = 0
    }

    $lblM = New-Object System.Windows.Forms.Label
    $lblM.Text = "Screen Monitor:"
    $lblM.Location = New-Object System.Drawing.Point(12, 58)
    $lblM.Size = New-Object System.Drawing.Size(210, 16)

    $cmbM = New-Object System.Windows.Forms.ComboBox
    $cmbM.Location = New-Object System.Drawing.Point(12, 76)
    $cmbM.Size = New-Object System.Drawing.Size(210, 22)
    $cmbM.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    for ($i = 0; $i -lt [System.Windows.Forms.Screen]::AllScreens.Count; $i++) {
        $s = [System.Windows.Forms.Screen]::AllScreens[$i]
        [void]$cmbM.Items.Add("Monitor $($i + 1) ($($s.Bounds.Width)x$($s.Bounds.Height))")
    }
    $cmbM.SelectedIndex = [Math]::Min($global:SelectedMonitorIndex, ($cmbM.Items.Count - 1))

    $chkCur = New-Object System.Windows.Forms.CheckBox
    $chkCur.Text = "Capture Mouse Cursor"
    $chkCur.Location = New-Object System.Drawing.Point(14, 110)
    $chkCur.Size = New-Object System.Drawing.Size(200, 20)
    $chkCur.Checked = $global:DrawCursor

    $chkHl = New-Object System.Windows.Forms.CheckBox
    $chkHl.Text = "Draw Red Halo Ring"
    $chkHl.Location = New-Object System.Drawing.Point(14, 132)
    $chkHl.Size = New-Object System.Drawing.Size(200, 20)
    $chkHl.Checked = $global:HighlightCursor

    $btnDone = New-Object System.Windows.Forms.Button
    $btnDone.Text = "Apply"
    $btnDone.Location = New-Object System.Drawing.Point(14, 160)
    $btnDone.Size = New-Object System.Drawing.Size(208, 25)
    $btnDone.Add_Click({
        $global:SelectedDocName = $cmbD.SelectedItem
        $global:SelectedMonitorIndex = $cmbM.SelectedIndex
        $global:DrawCursor = $chkCur.Checked
        $global:HighlightCursor = $chkHl.Checked
        $settingsForm.Close()
    })

    $settingsForm.Controls.AddRange(@($lblD, $cmbD, $lblM, $cmbM, $chkCur, $chkHl, $btnDone))
    $settingsForm.ShowDialog()
})

# Hotkeys (Ctrl + Shift = 0x0006)
$MOD = 0x0006
$VK_S = 0x53; $VK_W = 0x57; $VK_R = 0x52
$ID_FULL = 101; $ID_WIN = 102; $ID_SNIP = 103

$form.Add_Load({
    [Win32]::RegisterHotKey($form.Handle, $ID_FULL, $MOD, $VK_S) | Out-Null
    [Win32]::RegisterHotKey($form.Handle, $ID_WIN,  $MOD, $VK_W) | Out-Null
    [Win32]::RegisterHotKey($form.Handle, $ID_SNIP, $MOD, $VK_R) | Out-Null
})

$form.Add_FormClosing({
    $burstTimer.Stop()
    [Win32]::UnregisterHotKey($form.Handle, $ID_FULL) | Out-Null
    [Win32]::UnregisterHotKey($form.Handle, $ID_WIN)  | Out-Null
    [Win32]::UnregisterHotKey($form.Handle, $ID_SNIP) | Out-Null
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
})

$hotkeyCallback = [Action[int]]{
    param($id)
    if ($id -eq $ID_FULL) { & $actionExecute "FULL" }
    if ($id -eq $ID_WIN)  { & $actionExecute "WINDOW" }
    if ($id -eq $ID_SNIP) { & $actionExecute "SNIP" }
}

$msgFilter = [GlobalHotkeyFilter]::new($hotkeyCallback)
[System.Windows.Forms.Application]::AddMessageFilter($msgFilter)

$form.Controls.AddRange(@($btnCapture, $btnSettings, $btnUndo, $btnBurst, $btnSaveDoc, $btnExportPdf, $btnMinTray, $lblStatus))

[System.Windows.Forms.Application]::Run($form)

[System.Windows.Forms.Application]::RemoveMessageFilter($msgFilter)
