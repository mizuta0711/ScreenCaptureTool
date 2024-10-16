using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace ScreenCaptureTool.Utilities
{
    internal class Win32API
    {
        #region Win32 APIのインポート

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            internal int Left;
            internal int Top;
            internal int Right;
            internal int Bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool IsWindowVisible(IntPtr hWnd);

        internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        internal static extern int BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, IntPtr hdcSrc, int xSrc, int ySrc, int rop);

        [DllImport("gdi32.dll")]
        internal static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        internal static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        internal static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        internal static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        internal static extern bool DeleteDC(IntPtr hdc);

        #endregion Win32 APIのインポート

        #region Methods(Static)

        /// <summary>
        /// 部分一致でウィンドウを検索
        /// </summary>
        /// <param name="partialTitle">ウィンドウ名</param>
        /// <returns>ウィンドウハンドル(見つからなかった場合は0)</returns>
        internal static IntPtr FindWindowByTitle(string partialTitle)
        {
            IntPtr foundWindow = IntPtr.Zero;

            Win32API.EnumWindows((hWnd, lParam) =>
            {
                StringBuilder windowTitle = new StringBuilder(256);
                Win32API.GetWindowText(hWnd, windowTitle, 256);

                if (windowTitle.ToString().Contains(partialTitle, StringComparison.OrdinalIgnoreCase) && Win32API.IsWindowVisible(hWnd))
                {
                    foundWindow = hWnd;
                    return false; // ウィンドウが見つかったので列挙を終了
                }
                return true; // まだ見つかっていないので続行
            }, IntPtr.Zero);

            return foundWindow;
        }

        /// <summary>
        /// ウィンドウをキャプチャーしたBitmapを生成
        /// </summary>
        /// <param name="hWnd">ウィンドウハンドル</param>
        /// <param name="width">幅</param>
        /// <param name="height">高さ</param>
        /// <returns>Bitmap</returns>
        internal static Bitmap CaptureWindow(IntPtr hWnd, int width, int height)
        {
            const int SRCCOPY = 0x00CC0020;

            // ウィンドウのDCを取得
            IntPtr hdcWindow = Win32API.GetDC(hWnd);
            IntPtr hdcMemDC = Win32API.CreateCompatibleDC(hdcWindow);

            // ウィンドウのビットマップを作成
            IntPtr hBitmap = Win32API.CreateCompatibleBitmap(hdcWindow, width, height);
            IntPtr hOld = Win32API.SelectObject(hdcMemDC, hBitmap);

            // ウィンドウのビットブロック転送 (BitBlt) を実行
            Win32API.BitBlt(hdcMemDC, 0, 0, width, height, hdcWindow, 0, 0, SRCCOPY);

            // ビットマップを取得
            Bitmap bmp = Image.FromHbitmap(hBitmap);

            // リソース解放
            Win32API.SelectObject(hdcMemDC, hOld);
            Win32API.DeleteObject(hBitmap);
            Win32API.DeleteDC(hdcMemDC);

            // ウィンドウのDCを解放
            Win32API.DeleteDC(hdcWindow);

            return bmp;
        }

        #endregion Methods(Static)
    }
}