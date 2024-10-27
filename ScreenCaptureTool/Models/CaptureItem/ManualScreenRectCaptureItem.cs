using System;

namespace ScreenCaptureTool.Models.CaptureItem
{
    /// <summary>
    /// 画面矩形キャプチャーアイテム
    /// </summary>
    [Serializable]
    public class ManualScreenRectCaptureItem : ScreenRectCaptureItem
    {
        #region Properties

        /// <summary>
        /// 座標を固定するかどうか
        /// </summary>
        public bool LocationFixed { get; private set; }

        /// <summary>
        /// サイズを固定するかどうか
        /// </summary>
        public bool SizeFixed { get; private set; }

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ManualScreenRectCaptureItem() : base(0, 0, 0, 0)
        {
            LocationFixed = false;
            SizeFixed = false;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="size">キャプチャーサイズ</param>
        public ManualScreenRectCaptureItem(System.Drawing.Size size) : base(0, 0, size.Width, size.Height)
        {
            LocationFixed = false;
            SizeFixed = true;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="size">キャプチャーサイズ</param>
        public ManualScreenRectCaptureItem(System.Windows.Size size) : base(0, 0, (int)Math.Floor(size.Width), (int)Math.Floor(size.Height))
        {
            LocationFixed = false;
            SizeFixed = true;
        }

        #endregion Constructor
    }
}