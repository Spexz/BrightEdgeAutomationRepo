using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BrightEdgeAutomationTool
{
    public static class HWNDHelper
    {

        public static bool FindAndBringFwd(string title)
        {
            var target_hwnd = FindWindowByCaption(IntPtr.Zero, title);

            if (target_hwnd == IntPtr.Zero)
            {
                //Console.WriteLine("not found");
                return false;
            }

            BringWindowToFront(target_hwnd);
            return true;
        }

        public static bool IsRankTrackerOpen()
        {
            var handles = FindWindowsWithText("Rank Tracker v");
            //Console.WriteLine(handles.Count());

            IntPtr target_hwnd = handles.FirstOrDefault();

            if (target_hwnd == IntPtr.Zero)
            {
                return false;
            }

            return true;
        }

        public static IntPtr GetRankTrackerWindow()
        {
            var handles = FindWindowsWithText("Rank Tracker v");
            //Console.WriteLine(handles.Count());

            IntPtr target_hwnd = handles.FirstOrDefault();

            return target_hwnd;
        }

        public static void BringWindowToFront(IntPtr target_hwnd)
        {
            ShowWindowAsync(new HandleRef(null, target_hwnd), SW_RESTORE);
            SetForegroundWindow(target_hwnd);
        }

        public static void SetRankTrackerSizeAndPosition(IntPtr target_hwnd, Func<string, bool> updateStatus)
        {
            // Set the window's position.
            int width = 1370;
            int height = 700;
            int x = 0;
            int y = 0;

            var results = SetWindowPos(target_hwnd, IntPtr.Zero, x, y, width, height, 0);
            updateStatus("SetWindowPos success: " + results.ToString());
            updateStatus("SetWindowPos Error#: " + Marshal.GetLastWin32Error().ToString());


            results = MoveWindow(target_hwnd, x, y, width, height, false);
            updateStatus("MoveWindow success: " + results.ToString());
            updateStatus("MoveWindow Error#: " + Marshal.GetLastWin32Error().ToString());

            Debug.WriteLine(results);
            Debug.WriteLine(Marshal.GetLastWin32Error());
        }

        public static void WaitForWindowToRespond(IntPtr hwnd, int timeoutMinutes = 5)
        {
            LoopUntil(() => {
                if (IsHungAppWindow(hwnd) == false)
                    return true;

                return false;
            }, TimeSpan.FromMinutes(timeoutMinutes));
        }


        public static bool LoopUntil(Func<bool> task, TimeSpan timeSpan)
        {
            bool success = false;
            int elapsed = 0;
            while ((!success) && (elapsed < timeSpan.TotalMilliseconds))
            {
                Thread.Sleep(100);
                elapsed += 100;
                success = task();
            }
            return success;
        }





        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsHungAppWindow(IntPtr hWnd);



        [DllImport("kernel32.dll")]
        static extern uint GetLastError();



        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        // Delegate to filter which windows to include 
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        /// <summary> Get the text for the window pointed to by hWnd </summary>
        public static string GetWindowText(IntPtr hWnd)
        {
            int size = GetWindowTextLength(hWnd);
            if (size > 0)
            {
                var builder = new StringBuilder(size + 1);
                GetWindowText(hWnd, builder, builder.Capacity);
                return builder.ToString();
            }

            return String.Empty;
        }

        /// <summary> Find all windows that match the given filter </summary>
        /// <param name="filter"> A delegate that returns true for windows
        ///    that should be returned and false for windows that should
        ///    not be returned </param>
        public static IEnumerable<IntPtr> FindWindows(EnumWindowsProc filter)
        {
            IntPtr found = IntPtr.Zero;
            List<IntPtr> windows = new List<IntPtr>();

            EnumWindows(delegate (IntPtr wnd, IntPtr param)
            {
                if (filter(wnd, param))
                {
                    // only add the windows that pass the filter
                    windows.Add(wnd);
                }

                // but return true here so that we iterate all windows
                return true;
            }, IntPtr.Zero);

            return windows;
        }

        /// <summary> Find all windows that contain the given title text </summary>
        /// <param name="titleText"> The text that the window title must contain. </param>
        public static IEnumerable<IntPtr> FindWindowsWithText(string titleText)
        {
            return FindWindows(delegate (IntPtr wnd, IntPtr param)
            {
                return GetWindowText(wnd).Contains(titleText);
            });
        }

        public static IEnumerable<IntPtr> FindChildWindowsWithText(string titleText, IntPtr parentWnd)
        {
            var parent = GetParent(parentWnd);

            return FindWindows(delegate (IntPtr wnd, IntPtr param)
            {
                return GetWindowText(wnd).Contains(titleText) && wnd == parent;
            });
        }

        // Define the FindWindow API function.
        [DllImport("user32.dll", EntryPoint = "FindWindow",
            SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly,
            string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(HandleRef hWnd, int nCmdShow);

        public const int SW_RESTORE = 9;

        // Define the SetWindowPos API function.
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetWindowPos(IntPtr hWnd,
            IntPtr hWndInsertAfter, int X, int Y, int cx, int cy,
            SetWindowPosFlags uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        // Define the SetWindowPosFlags enumeration.
        [Flags()]
        private enum SetWindowPosFlags : uint
        {
            SynchronousWindowPosition = 0x4000,
            DeferErase = 0x2000,
            DrawFrame = 0x0020,
            FrameChanged = 0x0020,
            HideWindow = 0x0080,
            DoNotActivate = 0x0010,
            DoNotCopyBits = 0x0100,
            IgnoreMove = 0x0002,
            DoNotChangeOwnerZOrder = 0x0200,
            DoNotRedraw = 0x0008,
            DoNotReposition = 0x0200,
            DoNotSendChangingEvent = 0x0400,
            IgnoreResize = 0x0001,
            IgnoreZOrder = 0x0004,
            ShowWindow = 0x0040,
        }

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        public static extern IntPtr GetParent(IntPtr hWnd);
    }
}
