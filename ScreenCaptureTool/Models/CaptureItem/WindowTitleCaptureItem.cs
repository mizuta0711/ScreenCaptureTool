using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace ScreenCaptureTool.Models.CaptureItem
{
    /// <summary>
    /// ウィンドウタイトルキャプチャーアイテム
    /// </summary>
    internal class WindowTitleCaptureItem : CaptureItem
    {
        #region Win32 APIのインポート

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        private static extern int BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, IntPtr hdcSrc, int xSrc, int ySrc, int rop);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hdc);

        #endregion Win32 APIのインポート

        #region Properties

        /// <summary>
        /// キャプチャーするウィンドウタイトル
        /// </summary>
        public string WindowTitle { get; set; }

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="windowTitle">ウィンドウタイトル</param>
        public WindowTitleCaptureItem(string windowTitle)
        {
            WindowTitle = windowTitle;
        }

        #endregion Constructor

        #region Methods

        #region Methods(Private)

        // 部分一致でウィンドウを検索
        private IntPtr FindWindowByTitle(string partialTitle)
        {
            IntPtr foundWindow = IntPtr.Zero;

            EnumWindows((hWnd, lParam) =>
            {
                StringBuilder windowTitle = new StringBuilder(256);
                GetWindowText(hWnd, windowTitle, 256);

                if (windowTitle.ToString().Contains(partialTitle, StringComparison.OrdinalIgnoreCase) && IsWindowVisible(hWnd))
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
        private Bitmap CaptureWindow(IntPtr hWnd, int width, int height)
        {
            const int SRCCOPY = 0x00CC0020;

            // ウィンドウのDCを取得
            IntPtr hdcWindow = GetDC(hWnd);
            IntPtr hdcMemDC = CreateCompatibleDC(hdcWindow);

            // ウィンドウのビットマップを作成
            IntPtr hBitmap = CreateCompatibleBitmap(hdcWindow, width, height);
            IntPtr hOld = SelectObject(hdcMemDC, hBitmap);

            // ウィンドウのビットブロック転送 (BitBlt) を実行
            BitBlt(hdcMemDC, 0, 0, width, height, hdcWindow, 0, 0, SRCCOPY);

            // ビットマップを取得
            Bitmap bmp = Image.FromHbitmap(hBitmap);

            // リソース解放
            SelectObject(hdcMemDC, hOld);
            DeleteObject(hBitmap);
            DeleteDC(hdcMemDC);

            // ウィンドウのDCを解放
            DeleteDC(hdcWindow);

            return bmp;
        }

        #endregion Methods(Private)

        #region Methods(Override)

        /// <summary>
        /// キャプチャーを行う
        /// </summary>
        /// <returns>画像(失敗時はnull)</returns>
        public override Bitmap? Capture()
        {
            if (string.IsNullOrWhiteSpace(WindowTitle))
            {
                return null;
            }

            IntPtr hWnd = FindWindowByTitle(WindowTitle);
            if (hWnd == IntPtr.Zero)
            {
                return null;
            }

            // ウィンドウの位置とサイズを取得
            if (GetWindowRect(hWnd, out RECT rect) == false)
            {
                return null;
            }

            // RECTから幅と高さを計算
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            // ウィンドウ全体のビットマップを作成
            return CaptureWindow(hWnd, width, height);
        }

        #endregion Methods(Override)

        #endregion Methods
    }
}