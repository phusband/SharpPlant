using System;
using System.Runtime.InteropServices;
using System.Text;

using SharpPlant.SmartPlantReview;

namespace SharpPlant
{
    internal class NativeWin32
    {
        public delegate int EnumWindowsProcDelegate(int hWnd, int lParam);
        public delegate bool EnumChildCallback(int hwnd, ref int lParam);

        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_CLOSE = 0xF060;
        public const uint BM_CLICK = 0x00F5;

        public const int GW_HWNDFIRST = 0;
        public const int GW_HWNDLAST = 1;
        public const int GW_HWNDNEXT = 2;
        public const int GW_HWNDPREV = 3;
        public const int GW_OWNER = 4;
        public const int GW_CHILD = 5;

        [DllImport("user32.dll")]
        public static extern int FindWindow(
            string lpClassName, // class name 
            string lpWindowName // window name 
            );

        [DllImport("user32.dll")]
        public static extern IntPtr GetDlgItem(
            IntPtr hDlg, // parent hwnd
            int nIdDlgItem // controlID
            );

        [DllImport("Oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(
            int hwnd,
            uint dwObjectID,
            byte[] riid,
            ref dynamic ptr
            );

        [DllImport("User32.dll")]
        public static extern bool EnumChildWindows(
            int hWndParent,
            EnumChildCallback lpEnumFunc,
            ref int lParam
            );


        // Gets the 'nth' Window of the parent dialog based on the class name
        public static IntPtr FindWindowByIndex(IntPtr hWndParent, int index)
        {
            if (index == 0)
                return hWndParent;
            var ct = 0;
            var result = IntPtr.Zero;
            do
            {
                result = FindWindowEx(hWndParent, result, "Button", null);
                if (result != IntPtr.Zero)
                    ++ct;
            } while (ct < index && result != IntPtr.Zero);
            return result;
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(
            IntPtr parentHandle, // parent HWND
            IntPtr childAfter,
            string className, // name of the class
            string windowTitle // title of the child window
            );

        [DllImport("user32.dll")]
        public static extern int SendMessage(
            int hWnd, // handle to destination window 
            uint msg, // message 
            int wParam, // first message parameter 
            int lParam // second message parameter 
            );

        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(
            int hWnd // handle to window
            );

        [DllImport("User32.dll")]
        public static extern int GetClassName(
            int hWnd,
            StringBuilder lpClassName,
            int nMaxCount
            );

        [DllImport("user32")]
        public static extern int EnumWindows(EnumWindowsProcDelegate lpEnumFunc, int lParam);

        [DllImport("User32.Dll")]
        public static extern void GetWindowText(int h, StringBuilder s, int nMaxCount);

        [DllImport("user32", EntryPoint = "GetWindowLongA")]
        public static extern int GetWindowLongPtr(int hwnd, int nIndex);

        [DllImport("user32")]
        public static extern int GetParent(int hwnd);

        [DllImport("user32")]
        public static extern int GetWindow(int hwnd, int wCmd);

        [DllImport("user32")]
        public static extern int IsWindowVisible(int hwnd);

        [DllImport("user32")]
        public static extern int GetDesktopWindow();
    }
}