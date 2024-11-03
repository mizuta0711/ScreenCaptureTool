using System;
using System.Drawing;

using ScreenCaptureTool.Utilities;

namespace ScreenCaptureTool.Models.CaptureItem
{
    /// <summary>
    /// ウィンドウタイトルキャプチャーアイテム
    /// </summary>
    [Serializable]
    public class WindowTitleCaptureItem : CaptureItem
    {
        #region Properties

        /// <summary>
        /// キャプチャーするウィンドウタイトル
        /// </summary>
        public string TargetWindowTitle { get; private set; }

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="windowTitle">ウィンドウタイトル</param>
        public WindowTitleCaptureItem(string windowTitle)
        {
            TargetWindowTitle = windowTitle;
        }

        #endregion Constructor

        #region Methods

        #region Methods(Override)

        /// <summary>
        /// キャプチャーを行う
        /// </summary>
        /// <returns>画像(失敗時はnull)</returns>
        public override Bitmap? Capture()
        {
            // ウィンドウタイトルが空の場合はキャプチャーしない
            if (string.IsNullOrWhiteSpace(TargetWindowTitle))
            {
                return null;
            }

            // ウィンドウハンドルを取得
            IntPtr hWnd = CaptureHelper.FindWindowByTitle(TargetWindowTitle);
            if (hWnd == IntPtr.Zero)
            {
                return null;
            }

            // ウィンドウの位置とサイズを取得
            if (Win32API.GetWindowRect(hWnd, out Win32API.RECT rect) == false)
            {
                return null;
            }

            // RECTから幅と高さを計算
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            // ウィンドウ全体のビットマップを作成
            return CaptureHelper.CaptureWindow(hWnd, width, height);
        }

        #endregion Methods(Override)

        #endregion Methods
    }
}