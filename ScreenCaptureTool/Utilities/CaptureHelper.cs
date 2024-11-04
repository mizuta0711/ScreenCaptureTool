using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using static ScreenCaptureTool.Utilities.Win32API;

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

        /// <summary>
        /// ウィンドウタイトルの最大長
        /// </summary>
        private const int WindowTitleMaxLength = 256;

        #endregion Fields

        #region Methods(Static)

        /// <summary>
        /// 部分一致でウィンドウを検索
        /// </summary>
        /// <param name="title">ウィンドウタイトル</param>
        /// <param name="isEqual">完全一致かどうか</param>
        /// <returns>ウィンドウハンドル(見つからなかった場合は0)</returns>
        internal static IntPtr FindWindowByTitle(string title, bool isEqual = false)
        {
            IntPtr foundWindow = IntPtr.Zero;

            Win32API.EnumWindows((Win32API.EnumWindowsProc)((hWnd, lParam) =>
            {
                if (Win32API.IsWindowVisible(hWnd) == false)
                {
                    return true; // 非表示のウィンドウは無視して続行
                }

                // ウィンドウタイトルを取得
                StringBuilder windowTitle = new StringBuilder(WindowTitleMaxLength);
                Win32API.GetWindowText(hWnd, windowTitle, WindowTitleMaxLength);
                if (string.IsNullOrEmpty(windowTitle.ToString()))
                {
                    return true; // タイトルが空の場合は続行
                }

                if (isEqual)
                {
                    // 完全一致
                    if (windowTitle.ToString().Equals(title, StringComparison.OrdinalIgnoreCase))
                    {
                        foundWindow = hWnd;
                        return false; // ウィンドウが見つかったので列挙を終了
                    }
                }
                else
                {
                    // 部分一致
                    if (windowTitle.ToString().Contains(title, StringComparison.OrdinalIgnoreCase))
                    {
                        foundWindow = hWnd;
                        return false; // ウィンドウが見つかったので列挙を終了
                    }
                }

                return true; // まだ見つかっていないので続行
            }), IntPtr.Zero);

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
            // ウィンドウのデバイスコンテキストを取得
            IntPtr hWindowDC = Win32API.GetWindowDC(hWnd);
            Bitmap bitmap = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                IntPtr hDC = g.GetHdc();
                BitBlt(hDC, 0, 0, width, height, hWindowDC, 0, 0, (int)CopyPixelOperation.SourceCopy);
                g.ReleaseHdc(hDC);
            }

            // デバイスコンテキストの解放
            Win32API.ReleaseDC(hWnd, hWindowDC);
            return bitmap;
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