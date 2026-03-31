using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace FlaConsole
{
    internal class WindowPids
    {
        public static IReadOnlyCollection<uint> GetPidsWithTopLevelWindows(bool onlyVisible = true)
        {
            var pids = new HashSet<uint>();

            EnumWindows((hWnd, lParam) =>
            {
                if (!IsCandidateTopLevelWindow(hWnd))
                    return true; // continue enumeration

                // Optional: filter out "cloaked" windows (Win10+), tool windows, etc.
                GetWindowThreadProcessId(hWnd, out uint pid);
                if (pid != 0)
                {
                    pids.Add(pid);
                }

                return true; // continue enumeration
            }, IntPtr.Zero);

            return pids;
        }

        private static bool IsCandidateTopLevelWindow(IntPtr hWnd)
        {
            // Must be visible (minimized windows are still "visible" here)
            if (!IsWindowVisible(hWnd))
                return false;

            // Exclude message-only windows (defensive; EnumWindows shouldn't return them)
            if (GetParent(hWnd) == HWND_MESSAGE)
                return false;

            // Optional: exclude tool windows (often not "main app windows")
            var exStyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE).ToInt64();
            if ((exStyle & WS_EX_TOOLWINDOW) != 0)
                return false;

            // Optional: exclude cloaked windows (e.g., UWP background windows)
            if (IsCloaked(hWnd))
                return false;

            // Optional: require a non-empty rect (some odd windows can be 0x0)
            if (!GetWindowRect(hWnd, out var rc))
                return false;
            if (rc.Right <= rc.Left || rc.Bottom <= rc.Top)
                return false;

            // Optional: exclude "owned" popups; keep only unowned top-level windows.
            // Comment this out if you DO want owned top-level windows.
            if (GetWindow(hWnd, GW_OWNER) != IntPtr.Zero)
                return false;

            // Exclude standard dialog windows (MessageBox is typically class "#32770")
            string cls = GetClassNameString(hWnd);
            if (string.Equals(cls, "#32770", StringComparison.Ordinal))
                return false;

            return true;
        }

        private static string GetClassNameString(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            int len = GetClassName(hWnd, sb, sb.Capacity);
            return len > 0 ? sb.ToString(0, len) : string.Empty;
        }


        private static bool IsCloaked(IntPtr hWnd)
        {
            // DWM might not be available in some environments; treat failures as "not cloaked".
            int cloaked = 0;
            int hr = DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, out cloaked, sizeof(int));
            return hr == 0 && cloaked != 0;
        }


        // --- Win32 interop ---

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        private static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
            => IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : (IntPtr)GetWindowLong32(hWnd, nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);


        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out int pvAttribute, int cbAttribute);

        private const int GWL_EXSTYLE = -20;
        private const long WS_EX_TOOLWINDOW = 0x00000080L;

        private const uint GW_OWNER = 4;

        private static readonly IntPtr HWND_MESSAGE = new IntPtr(-3);

        private const int DWMWA_CLOAKED = 14;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }
    }
}
