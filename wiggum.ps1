Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing, System.Windows.Forms -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class Automation {
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT p);

    [DllImport("user32.dll")]
    public static extern IntPtr WindowFromPoint(POINT p);

    [DllImport("user32.dll")]
    public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder sb, int max);

    [DllImport("user32.dll")]
    public static extern int SetLayeredWindowAttributes(IntPtr hwnd, int crKey, byte alpha, int flags);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hwnd, int index);

    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hwnd, int index, int newLong);

    [DllImport("user32.dll")]
    public static extern bool UpdateLayeredWindow(
        IntPtr hwnd, IntPtr hdcDst, ref POINTAPI pptDst, ref SIZE psize,
        IntPtr hdcSrc, ref POINTAPI pptSrc, int crKey, ref BLENDFUNCTION pblend, int dwFlags);

    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);

    [DllImport("gdi32.dll")]
    public static extern bool DeleteDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern bool DeleteObject(IntPtr obj);

    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int val, int size);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINTAPI { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    public struct SIZE { public int cx; public int cy; }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    public struct BLENDFUNCTION {
        public byte BlendOp;
        public byte BlendFlags;
        public byte SourceConstantAlpha;
        public byte AlphaFormat;
    }

    public const uint WM_CHAR    = 0x0102;
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP   = 0x0101;
    public const int  VK_RETURN  = 0x0D;
    public const int  GWL_EXSTYLE = -20;
    public const int  WS_EX_LAYERED = 0x80000;
    public const int  ULW_ALPHA = 0x02;
    public const int  AC_SRC_OVER = 0x00;
    public const int  AC_SRC_ALPHA = 0x01;

    public static bool IsDown(int vk) {
        return (GetAsyncKeyState(vk) & 0x8000) != 0;
    }

    public static POINT GetCursor() {
        POINT p; GetCursorPos(out p); return p;
    }

    public static string GetTitle(IntPtr hwnd) {
        var sb = new System.Text.StringBuilder(256);
        GetWindowText(hwnd, sb, 256);
        return sb.ToString();
    }

    public static void SendText(IntPtr hwnd, string text) {
        foreach (char c in text) {
            PostMessage(hwnd, WM_CHAR, (IntPtr)c, IntPtr.Zero);
        }
    }

    public static void SendEnter(IntPtr hwnd) {
        PostMessage(hwnd, WM_KEYDOWN, (IntPtr)VK_RETURN, IntPtr.Zero);
        PostMessage(hwnd, WM_KEYUP, (IntPtr)VK_RETURN, IntPtr.Zero);
    }

    public static void MakeLayered(IntPtr hwnd) {
        int ex = GetWindowLong(hwnd, GWL_EXSTYLE);
        SetWindowLong(hwnd, GWL_EXSTYLE, ex | WS_EX_LAYERED);
    }

    public static void SetAlpha(IntPtr hwnd, byte alpha) {
        SetLayeredWindowAttributes(hwnd, 0, alpha, 0x2);
    }

    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    public static System.Collections.Generic.List<IntPtr> GetAllVisibleWindows() {
        var list = new System.Collections.Generic.List<IntPtr>();
        EnumWindows(delegate(IntPtr hwnd, IntPtr lParam) {
            if (IsWindowVisible(hwnd)) {
                var sb = new System.Text.StringBuilder(256);
                GetWindowText(hwnd, sb, 256);
                if (sb.Length > 0) list.Add(hwnd);
            }
            return true;
        }, IntPtr.Zero);
        return list;
    }

    public static void TryMica(IntPtr hwnd) {
        int val = 2;
        DwmSetWindowAttribute(hwnd, 38, ref val, 4);
        int dark = 1;
        DwmSetWindowAttribute(hwnd, 20, ref dark, 4);
    }

    public static void ApplyGlowBitmap(IntPtr formHwnd, Bitmap bmp, int x, int y) {
        IntPtr hBmp = bmp.GetHbitmap(Color.FromArgb(0));
        IntPtr memDc = CreateCompatibleDC(IntPtr.Zero);
        IntPtr old = SelectObject(memDc, hBmp);

        POINTAPI ptDst = new POINTAPI { X = x, Y = y };
        SIZE sz = new SIZE { cx = bmp.Width, cy = bmp.Height };
        POINTAPI ptSrc = new POINTAPI { X = 0, Y = 0 };
        BLENDFUNCTION blend = new BLENDFUNCTION {
            BlendOp = AC_SRC_OVER,
            BlendFlags = 0,
            SourceConstantAlpha = 255,
            AlphaFormat = AC_SRC_ALPHA
        };

        int ex = GetWindowLong(formHwnd, GWL_EXSTYLE);
        SetWindowLong(formHwnd, GWL_EXSTYLE, ex | WS_EX_LAYERED);
        UpdateLayeredWindow(formHwnd, IntPtr.Zero, ref ptDst, ref sz, memDc, ref ptSrc, 0, ref blend, ULW_ALPHA);

        SelectObject(memDc, old);
        DeleteObject(hBmp);
        DeleteDC(memDc);
    }

    // Inset glow - draws glow INSIDE the window rect
    public static Bitmap CreateInsetGlowBitmap(int w, int h, int glowSize, Color c1, Color c2) {
        Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp)) {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);

            for (int i = 0; i < glowSize; i++) {
                float t = (float)i / glowSize;
                int alpha = (int)(200 * (1 - t) * (1 - t));
                if (alpha > 255) alpha = 255;
                if (alpha < 0) alpha = 0;

                Color outer = Color.FromArgb(alpha, c1.R, c1.G, c1.B);
                Color inner = Color.FromArgb(alpha, c2.R, c2.G, c2.B);

                // Top edge inset
                using (var brush = new LinearGradientBrush(
                    new Point(0, 0), new Point(w, 0), outer, inner))
                    g.FillRectangle(brush, 0, i, w, 1);

                // Bottom edge inset
                using (var brush = new LinearGradientBrush(
                    new Point(0, 0), new Point(w, 0), outer, inner))
                    g.FillRectangle(brush, 0, h - i - 1, w, 1);

                // Left edge inset
                using (var brush = new LinearGradientBrush(
                    new Point(0, 0), new Point(0, h), outer, inner))
                    g.FillRectangle(brush, i, 0, 1, h);

                // Right edge inset
                using (var brush = new LinearGradientBrush(
                    new Point(0, 0), new Point(0, h), outer, inner))
                    g.FillRectangle(brush, w - i - 1, 0, 1, h);
            }
        }
        return bmp;
    }
}
"@

# --- Inset glow border ---
function Show-GlowBorder {
    param([IntPtr]$hwnd, [string]$tag, [switch]$NoFade)

    $rect = New-Object Automation+RECT
    [Automation]::GetWindowRect($hwnd, [ref]$rect) | Out-Null

    $x = $rect.Left
    $y = $rect.Top
    $w = $rect.Right - $rect.Left
    $h = $rect.Bottom - $rect.Top
    $glowSize = 20

    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.StartPosition = 'Manual'
    $form.Location = [System.Drawing.Point]::new($x, $y)
    $form.Size = [System.Drawing.Size]::new($w, $h)
    # Build bitmap BEFORE showing to prevent white flash
    $bmp = [Automation]::CreateInsetGlowBitmap($w, $h, $glowSize, [System.Drawing.Color]::Cyan, [System.Drawing.Color]::LimeGreen)

    # Tag label top-left inside the glow
    if ($tag) {
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $size = $g.MeasureString($tag, $font)
        $labelX = 6
        $labelY = 6
        $g.FillRectangle(
            (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(180, 0, 0, 0))),
            $labelX, $labelY, $size.Width + 12, $size.Height + 6)
        $g.DrawString($tag, $font,
            (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 0, 255, 255))),
            ($labelX + 6), ($labelY + 3))
        $font.Dispose()
        $g.Dispose()
    }

    # Show form and immediately apply the bitmap (minimizes flash)
    $form.Show()
    [Automation]::ApplyGlowBitmap($form.Handle, $bmp, $x, $y)
    $bmp.Dispose()

    if (-not $NoFade) {
        Start-Sleep -Milliseconds 200

        $steps = 14
        $stepMs = 50
        for ($i = 1; $i -le $steps; $i++) {
            $alpha = [double](1.0 - ($i / $steps))
            $fadeBmp = [Automation]::CreateInsetGlowBitmap($w, $h, $glowSize, [System.Drawing.Color]::Cyan, [System.Drawing.Color]::LimeGreen)

            $data = $fadeBmp.LockBits(
                (New-Object System.Drawing.Rectangle(0, 0, $w, $h)),
                [System.Drawing.Imaging.ImageLockMode]::ReadWrite,
                [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $bytes = [byte[]]::new($data.Stride * $h)
            [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $bytes, 0, $bytes.Length)
            for ($p = 3; $p -lt $bytes.Length; $p += 4) {
                $bytes[$p] = [byte]([int]($bytes[$p] * $alpha))
            }
            [System.Runtime.InteropServices.Marshal]::Copy($bytes, 0, $data.Scan0, $bytes.Length)
            $fadeBmp.UnlockBits($data)

            [Automation]::ApplyGlowBitmap($form.Handle, $fadeBmp, $x, $y)
            $fadeBmp.Dispose()

            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds $stepMs
        }

        $form.Close()
        $form.Dispose()
        return $null
    }

    return $form
}

# --- Naming dialog ---
function Get-NagName {
    param([string]$windowTitle)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Name this nag"
    $dlg.Size = [System.Drawing.Size]::new(340, 130)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.StartPosition = 'CenterScreen'
    $dlg.TopMost = $true
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
    $dlg.ForeColor = [System.Drawing.Color]::White

    try { [Automation]::TryMica($dlg.Handle) } catch {}

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $windowTitle
    $lbl.Location = [System.Drawing.Point]::new(12, 10)
    $lbl.Size = [System.Drawing.Size]::new(300, 20)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = [System.Drawing.Point]::new(12, 35)
    $txt.Size = [System.Drawing.Size]::new(220, 28)
    $txt.Font = New-Object System.Drawing.Font("Segoe UI", 12)
    $txt.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $txt.ForeColor = [System.Drawing.Color]::White
    $txt.BorderStyle = 'FixedSingle'

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "OK"
    $btn.Location = [System.Drawing.Point]::new(240, 35)
    $btn.Size = [System.Drawing.Size]::new(70, 28)
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = [System.Drawing.Color]::FromArgb(0, 170, 170)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btn.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlg.Close() })

    $dlg.AcceptButton = $btn
    $dlg.Controls.Add($lbl)
    $dlg.Controls.Add($txt)
    $dlg.Controls.Add($btn)

    $dlg.Add_Shown({ $txt.Focus() })

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txt.Text.Trim()) {
        return $txt.Text.Trim()
    }
    return $null
}

# --- Pie timer widget ---
$script:pieForm = $null

function Show-PieTimer {
    param([int]$slicesRemaining, [int]$totalSlices)

    $size = 64
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

    if ($script:pieForm) {
        $script:pieForm.Close()
        $script:pieForm.Dispose()
        $script:pieForm = $null
    }

    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.StartPosition = 'Manual'
    $form.Size = [System.Drawing.Size]::new($size, $size)
    # Float over the taskbar, near the clock
    $form.Location = [System.Drawing.Point]::new(
        $screen.Right - $size - 120,
        $screen.Bottom - $size + 2
    )
    $form.BackColor = [System.Drawing.Color]::Magenta
    $form.TransparencyKey = [System.Drawing.Color]::Magenta

    $remaining = $slicesRemaining
    $total = $totalSlices

    $form.Add_Paint({
        param($s, $e)
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $pad = 4
        $d = $s.ClientSize.Width - ($pad * 2)
        $r = [System.Drawing.Rectangle]::new($pad, $pad, $d, $d)

        # Dark background circle
        $bgBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(50, 50, 50))
        $g.FillEllipse($bgBrush, $r)
        $bgBrush.Dispose()

        # Draw remaining slices
        if ($remaining -gt 0) {
            $sliceAngle = 360.0 / $total
            for ($i = 0; $i -lt $remaining; $i++) {
                $startAngle = -90 + ($i * $sliceAngle)
                $ratio = $i / [Math]::Max($total - 1, 1)
                $cr = [int](0 + 50 * $ratio)
                $cg = [int](220 + 35 * $ratio)
                $cb = [int](220 - 100 * $ratio)
                $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, $cr, $cg, $cb))
                $g.FillPie($brush, $r, $startAngle, $sliceAngle - 1.5)
                $brush.Dispose()
            }
        }

        # Thin cyan border ring
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(180, 0, 255, 255), 1.5)
        $g.DrawEllipse($pen, $r)
        $pen.Dispose()

        # Center dot
        $dotSize = 6
        $cx = $pad + ($d / 2) - ($dotSize / 2)
        $cy = $pad + ($d / 2) - ($dotSize / 2)
        $dotBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200, 0, 0, 0))
        $g.FillEllipse($dotBrush, $cx, $cy, $dotSize, $dotSize)
        $dotBrush.Dispose()
    }.GetNewClosure())

    $form.Show()
    $form.Refresh()
    $script:pieForm = $form
}

function Hide-PieTimer {
    if ($script:pieForm) {
        $script:pieForm.Close()
        $script:pieForm.Dispose()
        $script:pieForm = $null
    }
}

# --- Status overlay (bottom of screen) ---
function Show-StatusOverlay {
    param($tabs)

    $lineHeight = 28
    $boxW = 520
    $boxH = 60 + ($tabs.Count * $lineHeight)
    $fullScreen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.StartPosition = 'Manual'
    $form.Size = [System.Drawing.Size]::new($boxW, $boxH)
    $form.Location = [System.Drawing.Point]::new(
        [int](($fullScreen.Width - $boxW) / 2) + $fullScreen.Left,
        $fullScreen.Bottom - $boxH
    )
    $form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

    try { [Automation]::TryMica($form.Handle) } catch {}

    $glowForms = @()
    $tabsCopy = $tabs

    $form.Add_Paint({
        param($s, $e)
        $g = $e.Graphics
        $w = $s.ClientSize.Width
        $h = $s.ClientSize.Height

        $pen = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            [System.Drawing.Point]::new(0, 0),
            [System.Drawing.Point]::new($w, 0),
            [System.Drawing.Color]::Cyan,
            [System.Drawing.Color]::LimeGreen
        )
        $g.FillRectangle($pen, 0, 0, $w, 2)
        $g.FillRectangle($pen, 0, $h - 2, $w, 2)
        $g.FillRectangle($pen, 0, 0, 2, $h)
        $g.FillRectangle($pen, $w - 2, 0, 2, $h)
        $pen.Dispose()

        $titleFont = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
        $g.DrawString("NAGGING TARGETS", $titleFont, [System.Drawing.Brushes]::Cyan, 15, 12)
        $titleFont.Dispose()

        $font = New-Object System.Drawing.Font("Consolas", 11)
        $y = 45
        for ($i = 0; $i -lt $tabsCopy.Count; $i++) {
            $t = $tabsCopy[$i]
            $alive = [Automation]::IsWindow($t.Hwnd)
            $freshTitle = if ($alive) { [Automation]::GetTitle($t.HwndTop) } else { $t.Title }
            $status = if ($alive) { "ALIVE" } else { "DEAD" }
            $color = if ($alive) { [System.Drawing.Color]::LimeGreen } else { [System.Drawing.Color]::Red }
            $brush = New-Object System.Drawing.SolidBrush($color)
            $line = "  #$($i+1) [$status] $($t.Tag) - $freshTitle"
            $g.DrawString($line, $font, $brush, 10, $y)
            $brush.Dispose()
            $y += $lineHeight
        }
        $font.Dispose()
    }.GetNewClosure())

    foreach ($t in $tabs) {
        if ([Automation]::IsWindow($t.HwndTop)) {
            $glow = Show-GlowBorder $t.HwndTop $t.Tag -NoFade
            $glowForms += $glow
        }
    }

    $form.Show()
    $form.Refresh()
    Start-Sleep -Seconds 3

    foreach ($g in $glowForms) { if ($g) { $g.Close(); $g.Dispose() } }
    $form.Close()
    $form.Dispose()
}

# --- Countdown overlay (dark + ornate) ---
function Show-Countdown {
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $boxW = 140
    $boxH = 110

    foreach ($n in 3, 2, 1) {
        $form = New-Object System.Windows.Forms.Form
        $form.FormBorderStyle = 'None'
        $form.TopMost = $true
        $form.ShowInTaskbar = $false
        $form.StartPosition = 'Manual'
        $form.Size = [System.Drawing.Size]::new($boxW, $boxH)
        $form.Location = [System.Drawing.Point]::new(
            $screen.Right - $boxW - 20,
            $screen.Bottom - $boxH - 20
        )
        $form.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)

        $num = $n
        $cw = [int]$boxW
        $ch = [int]$boxH
        $form.Add_Paint({
            param($s, $e)
            $g = $e.Graphics
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

            # Gradient border (cyan to lime)
            $borderBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                [System.Drawing.Point]::new(0, 0),
                [System.Drawing.Point]::new($cw, $ch),
                [System.Drawing.Color]::Cyan,
                [System.Drawing.Color]::LimeGreen
            )
            $borderPen = New-Object System.Drawing.Pen($borderBrush, 3)
            $g.DrawRectangle($borderPen, 2, 2, ($cw - 4), ($ch - 4))
            $borderPen.Dispose()
            $borderBrush.Dispose()

            # Inner subtle border
            $innerPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 60, 60), 1)
            $g.DrawRectangle($innerPen, 5, 5, ($cw - 10), ($ch - 10))
            $innerPen.Dispose()

            # Burnt orange number
            $burntOrange = [System.Drawing.Color]::FromArgb(204, 85, 0)
            $font = New-Object System.Drawing.Font("Segoe UI", 48, [System.Drawing.FontStyle]::Bold)
            $brush = New-Object System.Drawing.SolidBrush($burntOrange)
            $sf = New-Object System.Drawing.StringFormat
            $sf.Alignment = [System.Drawing.StringAlignment]::Center
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
            $numRect = New-Object System.Drawing.RectangleF(0, -2, $cw, $ch)
            $g.DrawString("$num", $font, $brush, $numRect, $sf)
            $font.Dispose(); $brush.Dispose()

            # Small label underneath
            $labelFont = New-Object System.Drawing.Font("Segoe UI", 8)
            $labelBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(100, 100, 100))
            $sf2 = New-Object System.Drawing.StringFormat
            $sf2.Alignment = [System.Drawing.StringAlignment]::Center
            $labelRect = New-Object System.Drawing.RectangleF(0, ($ch - 24), $cw, 20)
            $g.DrawString("STARTING", $labelFont, $labelBrush, $labelRect, $sf2)
            $labelFont.Dispose(); $labelBrush.Dispose()
            $sf.Dispose(); $sf2.Dispose()
        }.GetNewClosure())

        $form.Show()
        $form.Refresh()
        Start-Sleep -Seconds 1
        $form.Close()
        $form.Dispose()
    }
}

# --- Auto-discovery fingerprints ---
function Find-ChatWindows {
    # Title patterns - easy, high confidence fingerprints
    # Reject patterns - skip windows matching these (file viewers, images, etc)
    $rejectPatterns = @(
        "\.png", "\.jpg", "\.jpeg", "\.gif", "\.bmp", "\.svg",
        "\.pdf", "\.mp4", "\.mp3", "\.zip",
        "Photos", "Image Viewer", "Paint"
    )

    $titlePatterns = @(
        @{ Pattern = "Claude Code";    Tag = "Claude Code" }
        @{ Pattern = "claude\.ai";     Tag = "Claude Web" }
        @{ Pattern = "claude\.com";    Tag = "Claude Web" }
        @{ Pattern = "Claude -";       Tag = "Claude" }
        @{ Pattern = "chatgpt\.com";   Tag = "ChatGPT" }
        @{ Pattern = "ChatGPT -";      Tag = "ChatGPT" }
        @{ Pattern = "gemini\.google"; Tag = "Gemini" }
        @{ Pattern = "Gemini -";       Tag = "Gemini" }
        @{ Pattern = "Copilot";        Tag = "Copilot" }
        @{ Pattern = "Cursor -";       Tag = "Cursor" }
        @{ Pattern = "cursor\.com";    Tag = "Cursor" }
        @{ Pattern = "Midjourney -";   Tag = "Midjourney" }
        @{ Pattern = "midjourney\.com"; Tag = "Midjourney" }
        @{ Pattern = "Ubuntu";         Tag = "Terminal" }
        @{ Pattern = "WSL";            Tag = "Terminal" }
    )

    $found = @()
    $windows = [Automation]::GetAllVisibleWindows()

    foreach ($hwnd in $windows) {
        $title = [Automation]::GetTitle($hwnd)
        if (-not $title) { continue }

        # Skip our own nag script window
        $myPid = $PID
        $winPid = [uint32]0
        [Automation]::GetWindowThreadProcessId($hwnd, [ref]$winPid) | Out-Null
        if ($winPid -eq $myPid) { continue }

        # Skip obvious non-chat windows
        $rejected = $false
        foreach ($rp in $rejectPatterns) {
            if ($title -match $rp) { $rejected = $true; break }
        }
        if ($rejected) { continue }

        foreach ($fp in $titlePatterns) {
            if ($title -match $fp.Pattern) {
                $found += [PSCustomObject]@{
                    Hwnd    = $hwnd
                    HwndTop = $hwnd
                    Title   = $title
                    Tag     = $fp.Tag
                    Source  = "auto"
                }
                break
            }
        }
    }

    return $found
}

# ============================================
# MAIN
# ============================================

$tabs = [System.Collections.ArrayList]::new()
$lastClickState = $false
$seenHwnds = @{}
$tagCounter = 0

# Phase 0: Auto-discover chat windows
Write-Host "=== SCANNING ==="
$autoFound = Find-ChatWindows

if ($autoFound.Count -gt 0) {
    Write-Host "  Found $($autoFound.Count) chat windows:"
    foreach ($w in $autoFound) {
        $tagCounter++
        $w.Tag = "$($w.Tag)-$tagCounter"
        $key = $w.Hwnd.ToInt64()
        if (-not $seenHwnds.ContainsKey($key)) {
            $seenHwnds[$key] = $true
            $tabs.Add($w) | Out-Null
            Write-Host "  + $($w.Tag)  [$($w.Title)]"
            Show-GlowBorder $w.HwndTop $w.Tag
        }
    }
} else {
    Write-Host "  No chat windows found automatically."
}

Write-Host ""
Write-Host "=== MANUAL TAGGING ==="
Write-Host "Right-Alt + Click         = tag window (auto-named)"
Write-Host "Right-Alt + '.' + Click   = tag window (you name it)"
Write-Host "Ctrl+K                    = start nagging"
Write-Host ""
Write-Host "During run: Ctrl+K+K to show status overlay."
Write-Host ""

# Phase 1: Manual record (supplement auto-discovery)
while ($true) {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 20

    $ctrl = [Automation]::IsDown(0xA2) -or [Automation]::IsDown(0xA3)
    $k = [Automation]::IsDown(0x4B)
    if ($ctrl -and $k) { break }

    $altDown = [Automation]::IsDown(0xA5)
    $dotDown = [Automation]::IsDown(0xBE)
    $clickDown = [Automation]::IsDown(0x01)

    if ($altDown -and $clickDown -and -not $lastClickState) {
        $pos = [Automation]::GetCursor()
        $hwndChild = [Automation]::WindowFromPoint($pos)
        $hwndTop = [Automation]::GetAncestor($hwndChild, 2)
        $title = [Automation]::GetTitle($hwndTop)
        $key = $hwndChild.ToInt64()

        if (-not $seenHwnds.ContainsKey($key)) {
            $tagCounter++

            if ($dotDown) {
                $customName = Get-NagName $title
                if ($customName) {
                    $tag = $customName
                } else {
                    $tag = "NAG-$tagCounter"
                }
            } else {
                $tag = "NAG-$tagCounter"
            }

            $seenHwnds[$key] = $true
            $tabs.Add([PSCustomObject]@{
                Hwnd    = $hwndChild
                HwndTop = $hwndTop
                Title   = $title
                Tag     = $tag
            }) | Out-Null
            Write-Host "  $tag  [$title]  hwnd=$hwndChild"

            Show-GlowBorder $hwndTop $tag
        } else {
            Write-Host "  (already tagged, skipping)"
        }
    }
    $lastClickState = $clickDown
}

if ($tabs.Count -eq 0) {
    Write-Host "No windows tagged. Exiting."
    exit 1
}

Write-Host ""
Write-Host "  Starting in"
Show-Countdown

Write-Host ""
Write-Host "=== LOOPING ($($tabs.Count) targets, every 3 min) ==="
Write-Host "    Ctrl+K+K to see status overlay"
Write-Host ""

$message = "status update please"
$lastKKTime = [datetime]::MinValue
$totalSlices = 6
$intervalSec = 180
$sliceInterval = $intervalSec / $totalSlices

while ($true) {
    # Send nag
    for ($i = 0; $i -lt $tabs.Count; $i++) {
        $t = $tabs[$i]
        if (-not [Automation]::IsWindow($t.Hwnd)) {
            Write-Host "  $($t.Tag) [$($t.Title)] - DEAD, skipping"
            continue
        }
        $t.Title = [Automation]::GetTitle($t.HwndTop)

        [Automation]::SendText($t.Hwnd, $message)
        Start-Sleep -Milliseconds 50
        [Automation]::SendEnter($t.Hwnd)
        Write-Host "  $($t.Tag) [$($t.Title)] - sent"
    }

    $now = Get-Date -Format "HH:mm:ss"
    Write-Host "$now - done, sleeping 3 min"

    # Pie countdown
    $cycleStart = Get-Date
    $currentSlice = $totalSlices
    Show-PieTimer $currentSlice $totalSlices

    while ($true) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100

        $elapsed = ((Get-Date) - $cycleStart).TotalSeconds

        # Update pie
        $newSlice = $totalSlices - [Math]::Floor($elapsed / $sliceInterval)
        if ($newSlice -lt 0) { $newSlice = 0 }
        if ($newSlice -ne $currentSlice) {
            $currentSlice = $newSlice
            if ($currentSlice -gt 0) {
                Show-PieTimer $currentSlice $totalSlices
            } else {
                Hide-PieTimer
            }
        }

        # Ctrl+K+K check
        $ctrl = [Automation]::IsDown(0xA2) -or [Automation]::IsDown(0xA3)
        $k = [Automation]::IsDown(0x4B)
        if ($ctrl -and $k) {
            $kkNow = Get-Date
            if (($kkNow - $lastKKTime).TotalSeconds -gt 4) {
                $lastKKTime = $kkNow
                Write-Host "  [Ctrl+K+K] Showing status"
                Show-StatusOverlay $tabs
            }
        }

        if ($elapsed -ge $intervalSec) {
            Hide-PieTimer
            break
        }
    }
}
