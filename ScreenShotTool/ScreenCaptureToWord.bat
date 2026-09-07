<# :batch_launcher
@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((Get-Content -LiteralPath '%~f0') -join [Environment]::NewLine)"
exit /b
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Add Win32 API support for system-wide Hotkeys
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
"@

# Create floating dashboard
$form = New-Object System.Windows.Forms.Form
$form.Text = "Doc Snapper"
$form.Size = New-Object System.Drawing.Size(230, 155)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow

# Capture Button
$btnCapture = New-Object System.Windows.Forms.Button
$btnCapture.Text = "Capture & Paste`n(Ctrl + Shift + S)"
$btnCapture.Size = New-Object System.Drawing.Size(190, 55)
$btnCapture.Location = New-Object System.Drawing.Point(12, 12)
$btnCapture.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# Status Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Press Ctrl + Shift + S"
$lblStatus.Size = New-Object System.Drawing.Size(190, 20)
$lblStatus.Location = New-Object System.Drawing.Point(12, 75)
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

# Core capture and paste function
$actionCapture = {
    try {
        $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        if ($word.Documents.Count -eq 0) {
            $lblStatus.Text = "No Word doc open!"
            return
        }
        $doc = $word.ActiveDocument
    }
    catch {
        $lblStatus.Text = "Word is not running!"
        return
    }

    # Hide dashboard briefly so it doesn't show in screenshot
    $form.Opacity = 0
    [System.Threading.Thread]::Sleep(200)

    # Capture Screen
    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

    # Temporary file
    $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "doc_snap_$([System.Guid]::NewGuid()).png")
    $bitmap.Save($tempFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()

    # Append to Word
    try {
        $range = $doc.Content
        $range.Collapse(0) # 0 = wdCollapseEnd
        $doc.InlineShapes.AddPicture($tempFile, $false, $true, $range) | Out-Null
        $doc.Content.InsertParagraphAfter()
        $lblStatus.Text = "Pasted! (Ready)"
    }
    catch {
        $lblStatus.Text = "Failed to paste."
    }
    finally {
        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force
        }
    }

    $form.Opacity = 1
}

# Attach to button
$btnCapture.Add_Click($actionCapture)

# Hotkey settings:
# Modifiers: MOD_ALT = 0x0001, MOD_CONTROL = 0x0002, MOD_SHIFT = 0x0004
# Ctrl (0x02) + Shift (0x04) = 0x06
$MOD_CTRL_SHIFT = 0x0006
$VK_S = 0x53          # Virtual key code for letter 'S'
$HOTKEY_ID = 1001

$form.Add_Load({
    $registered = [Win32]::RegisterHotKey($form.Handle, $HOTKEY_ID, $MOD_CTRL_SHIFT, $VK_S)
    if (-not $registered) {
        $lblStatus.Text = "Shortcut conflict! Use button."
    }
})

$form.Add_FormClosing({
    [Win32]::UnregisterHotKey($form.Handle, $HOTKEY_ID)
})

# Hotkey message listener
Add-Type -ReferencedAssemblies "System.Windows.Forms" @"
using System;
using System.Windows.Forms;

public class HotkeyFilter : IMessageFilter {
    private Action _callback;
    public HotkeyFilter(Action callback) {
        _callback = callback;
    }
    public bool PreFilterMessage(ref Message m) {
        const int WM_HOTKEY = 0x0312;
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == 1001) {
            if (_callback != null) {
                _callback();
            }
            return true;
        }
        return false;
    }
}
"@

$filter = [HotkeyFilter]::new([Action]$actionCapture)
[System.Windows.Forms.Application]::AddMessageFilter($filter)

$form.Controls.Add($btnCapture)
$form.Controls.Add($lblStatus)

[System.Windows.Forms.Application]::Run($form)

# Cleanup
[System.Windows.Forms.Application]::RemoveMessageFilter($filter)
[Win32]::UnregisterHotKey($form.Handle, $HOTKEY_ID)