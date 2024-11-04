using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ScreenCaptureTool.Utilities
{
    /// <summary>
    /// キャプチャーのヘルパークラス
    /// </summary>
    internal class CaptureHelper
    {
        #region Fields

        /// <summary>
        /// デフォルトDPI
        /// </summary>
        private const int DefaultDpi = 96;

        #endregion Fields

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

        /// <summary>
        /// ウィンドウのDPIを取得
        /// </summary>
        /// <param name="hwnd">ウィンドウハンドル</param>
        /// <returns>DPI(デフォルト:96dpi)</returns>
        internal static uint GetDpiForCurrentWindow(IntPtr hwnd)
        {
            if (Environment.OSVersion.Version.Major >= 10)
            {
                try
                {
                    return Win32API.GetDpiForWindow(hwnd);
                }
                catch
                {
                    return DefaultDpi;
                }
            }
            return DefaultDpi;
        }

        /// <summary>
        /// DPIスケール取得メソッド
        /// </summary>
        /// <param name="window">WPFウィンドウ</param>
        /// <returns>DPIスケール</returns>
        internal static (double DpiX, double DpiY) GetDpiScale(Window window)
        {
            var source = PresentationSource.FromVisual(window);
            if (source?.CompositionTarget == null)
            {
                return (1.0, 1.0); // デフォルトスケール
            }

            // DPIスケールを取得
            double dpiX = source.CompositionTarget.TransformToDevice.M11;
            double dpiY = source.CompositionTarget.TransformToDevice.M22;
            return (dpiX, dpiY);
        }

        /// <summary>
        /// DPIスケール取得メソッド
        /// </summary>
        /// <param name="window">WPFウィンドウ</param>
        /// <returns>DPIスケール</returns>
        internal static (double DpiX, double DpiY) GetDpiScale(IntPtr hwnd)
        {
            uint dpi = GetDpiForCurrentWindow(hwnd);
            double scale = (double)dpi / (double)DefaultDpi;
            return (scale, scale);
        }

        /// <summary>
        /// DPIスケールを取得
        /// </summary>
        /// <returns>DPIスケール</returns>
        internal static (float scaleX, float scaleY) GetDpiScale()
        {
            const int DisplayDPI = 96;
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) // デスクトップの Graphics を取得
            {
                return (g.DpiX / DisplayDPI, g.DpiY / DisplayDPI);
            }
        }

        #endregion Methods(Static)
    }
}