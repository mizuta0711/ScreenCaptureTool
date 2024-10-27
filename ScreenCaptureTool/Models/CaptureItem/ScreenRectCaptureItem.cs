using System;
using System.Drawing;
using System.Windows;

using PixelFormat = System.Drawing.Imaging.PixelFormat;

namespace ScreenCaptureTool.Models.CaptureItem
{
    /// <summary>
    /// 画面矩形キャプチャーアイテム
    /// </summary>
    [Serializable]
    public class ScreenRectCaptureItem : CaptureItem
    {
        #region Properties

        /// <summary>
        /// キャプチャー範囲
        /// </summary>
        public Rectangle TargetRect { get; set; }

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="left">キャプチャー範囲:X座標</param>
        /// <param name="top">キャプチャー範囲:Y座標</param>
        /// <param name="width">キャプチャー範囲:幅</param>
        /// <param name="height">キャプチャー範囲:高さ</param>
        public ScreenRectCaptureItem(int left, int top, int width, int height)
        {
            TargetRect = new Rectangle(left, top, width, height);
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="rect">キャプチャー範囲</param>
        public ScreenRectCaptureItem(Rectangle rect) : this(rect.Left, rect.Top, rect.Width, rect.Height)
        {
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="rect">キャプチャー範囲</param>
        public ScreenRectCaptureItem(Rect rect) : this((int)Math.Floor(rect.Left), (int)Math.Floor(rect.Top), (int)Math.Floor(rect.Width), (int)Math.Floor(rect.Height))
        {
        }

        #endregion Constructor

        #region Methods(Override)

        /// <summary>
        /// デスクトップの指定範囲をキャプチャーする
        /// </summary>
        /// <returns>true: 成功　/ false: 失敗</returns>
        public override Bitmap? Capture()
        {
            // デスクトップの解像度を取得
            int screenWidth = (int)SystemParameters.VirtualScreenWidth;
            int screenHeight = (int)SystemParameters.VirtualScreenHeight;

            int x = TargetRect.Left;
            int y = TargetRect.Top;
            int width = TargetRect.Width;
            int height = TargetRect.Height;

            // 矩形のサイズをチェックして調整
            if (x < 0 || y < 0 || width <= 0 || height <= 0 ||
                x + width > screenWidth || y + height > screenHeight)
            {
                return null;
            }

            // 矩形のビットマップを作成
            Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
            }
            return bitmap;
        }

        #endregion Methods(Override)
    }
}