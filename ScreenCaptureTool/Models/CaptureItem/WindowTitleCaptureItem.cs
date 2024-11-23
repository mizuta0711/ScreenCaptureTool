using System;
using System.Drawing;
using System.Drawing.Imaging;

using ScreenCaptureTool.Utilities;

namespace ScreenCaptureTool.Models.CaptureItem
{
    /// <summary>
    /// ウィンドウタイトルキャプチャーアイテム
    /// </summary>
    [Serializable]
    public class WindowTitleCaptureItem : CaptureItemBase
    {
        #region Properties

        /// <summary>
        /// キャプチャーするウィンドウタイトル
        /// </summary>
        public string TargetWindowTitle { get; private set; }

        /// <summary>
        /// ウィンドウのクライアント領域をキャプチャーするかどうか
        /// </summary>
        public bool UseDirectCapture { get; private set; } = false;

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="windowTitle">ウィンドウタイトル</param>
        /// <param name="useDirectCapture">ウィンドウのクライアント領域をキャプチャーするかどうか</param>
        public WindowTitleCaptureItem(string windowTitle, bool useDirectCapture = false)
        {
            TargetWindowTitle = windowTitle;
            UseDirectCapture = useDirectCapture;
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

            // キャプチャー
            int width = rect.Right - rect.Left + 1;
            int height = rect.Bottom - rect.Top + 1;
            if (UseDirectCapture)
            {
                // ウィンドウのクライアント領域をキャプチャー
                return CaptureHelper.CaptureWindow(hWnd, width, height);
            }
            else
            {
                // ウィンドウ位置のスクリーンをキャプチャー
                Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height));
                }
                return bitmap;
            }
        }

        #endregion Methods(Override)

        #endregion Methods
    }
}