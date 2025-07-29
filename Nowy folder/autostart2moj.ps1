Param(
    [string] $proc="C:\skrypty\znakowarka_EzCadA",
    [string] $proca="C:\skrypty\laser.ezd",
    [string] $adm
)
Clear-Host

Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WinAp {
      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool SetForegroundWindow(IntPtr hWnd);

      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

      [DllImport("user32.dll")]
      public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    }
"@

$znakowarkaProc = Get-Process | Where-Object {$_.MainWindowTitle -match "znakowarka_EzCadA"}
$laserProc = Get-Process | Where-Object {$_.MainWindowTitle -match "laser.ezd"}

if ($laserProc -ne $null) {
    $laserLastInput = [System.Windows.Forms.SystemInformation]::IdleTime
    if ($laserLastInput.TotalMinutes -ge 5) {
        $laserProc.CloseMainWindow()
        $laserProc.WaitForExit()
    }
}

if ($znakowarkaProc -eq $null) {
    Start-Process "$proc"
} else {
    [WinAp]::ShowWindow($znakowarkaProc.MainWindowHandle, 9)  # SW_RESTORE
    [WinAp]::SetForegroundWindow($znakowarkaProc.MainWindowHandle)
}
