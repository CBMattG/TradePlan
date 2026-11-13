# Creates "Tradeplan" shortcuts (Desktop + Start menu) that open the desktop app (Tradeplan.pyw)
# with the Tradeplan icon. Safe to re-run (e.g. after moving the folder or on a new PC).
#   powershell -ExecutionPolicy Bypass -File create_shortcuts.ps1
# The shortcut carries the same AppUserModelID the app sets on itself, so a shortcut pinned to
# the taskbar and the running window share one taskbar button.
$ErrorActionPreference = 'Stop'
$app   = Split-Path -Parent $MyInvocation.MyCommand.Path
$pyw   = Join-Path $app 'Tradeplan.pyw'
$icon  = Join-Path $app 'tradeplan.ico'
$appId = 'CharlesBentley.Tradeplan'

# pythonw.exe: the Python launcher alias if present (survives Python upgrades), else next to python.exe
$pythonw = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\pythonw.exe'
if (-not (Test-Path $pythonw)) {
  $py = (& python -c "import sys; print(sys.executable)").Trim()
  $pythonw = Join-Path (Split-Path $py) 'pythonw.exe'
}
if (-not (Test-Path $pythonw)) { throw "pythonw.exe not found - is Python installed?" }

Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
[ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IPropertyStore {
  int GetCount(out uint c); int GetAt(uint i, out PROPERTYKEY k);
  int GetValue(ref PROPERTYKEY k, [Out] PROPVARIANT v); int SetValue(ref PROPERTYKEY k, [In] PROPVARIANT v); int Commit();
}
[StructLayout(LayoutKind.Sequential, Pack = 4)] public struct PROPERTYKEY { public Guid fmtid; public uint pid; }
[StructLayout(LayoutKind.Sequential)] public class PROPVARIANT { public ushort vt; ushort r1, r2, r3; public IntPtr p; IntPtr p2; }
public static class LnkAppId {
  [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
  static extern int SHGetPropertyStoreFromParsingName(string path, IntPtr pbc, int flags, ref Guid iid, out IPropertyStore ps);
  public static void Set(string lnk, string id) {
    Guid iid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");
    IPropertyStore ps; Marshal.ThrowExceptionForHR(SHGetPropertyStoreFromParsingName(lnk, IntPtr.Zero, 2 /*GPS_READWRITE*/, ref iid, out ps));
    var key = new PROPERTYKEY { fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), pid = 5 };   // System.AppUserModel.ID
    var v = new PROPVARIANT { vt = 31 /*VT_LPWSTR*/, p = Marshal.StringToCoTaskMemUni(id) };
    try { Marshal.ThrowExceptionForHR(ps.SetValue(ref key, v)); Marshal.ThrowExceptionForHR(ps.Commit()); }
    finally { Marshal.FreeCoTaskMem(v.p); Marshal.ReleaseComObject(ps); }
  }
}
"@

$shell = New-Object -ComObject WScript.Shell
$targets = @(
  (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Tradeplan.lnk'),
  (Join-Path ([Environment]::GetFolderPath('Programs')) 'Tradeplan.lnk')
)
foreach ($lnk in $targets) {
  $s = $shell.CreateShortcut($lnk)
  $s.TargetPath = $pythonw
  $s.Arguments = '"' + $pyw + '"'
  $s.WorkingDirectory = $app
  $s.IconLocation = "$icon,0"
  $s.Description = 'Tradeplan - buying & forecast plan'
  $s.Save()
  try { [LnkAppId]::Set($lnk, $appId) } catch { Write-Warning "Couldn't tag $lnk for taskbar grouping: $_" }
  "created $lnk"
}
