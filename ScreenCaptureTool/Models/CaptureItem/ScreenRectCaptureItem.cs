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
    internal class ScreenRectCaptureItem : CaptureItem
    {
        #region Properties

        /// <summary>
        /// キャプチャー範囲：X
        /// </summary>
        public int Left { get; set; }

        /// <summary>
        /// キャプチャー範囲：Y
        /// </summary>
        public int Top { get; set; }

        /// <summary>
        /// キャプチャー範囲：幅
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// キャプチャー範囲：高さ
        /// </summary>
        public int Height { get; set; }

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="left">X座標</param>
        /// <param name="top">Y座標</param>
        /// <param name="width">幅</param>
        /// <param name="height">高さ</param>
        public ScreenRectCaptureItem(int left, int top, int width, int height)
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
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

            int x = Left;
            int y = Top;
            int width = Width;
            int height = Height;

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