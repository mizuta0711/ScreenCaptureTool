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

        #region Constructors(Drawing)

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="point">キャプチャー座標</param>
        public ManualScreenRectCaptureItem(System.Drawing.Point point) : base(point.X, point.Y, 0, 0)
        {
            LocationFixed = true;
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
        /// <param name="point">キャプチャー座標</param>
        /// <param name="size">キャプチャーサイズ</param>
        public ManualScreenRectCaptureItem(System.Drawing.Point point, System.Drawing.Size size) : base(point.X, point.Y, size.Width, size.Height)
        {
            LocationFixed = true;
            SizeFixed = true;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="rect">キャプチャー矩形</param>
        public ManualScreenRectCaptureItem(System.Drawing.Rectangle rect) : this(rect.Location, rect.Size)
        {
        }

        #endregion Constructors(Drawing)

        #region Constructors(Windows)

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="point">キャプチャー座標</param>
        public ManualScreenRectCaptureItem(System.Windows.Point point) : base((int)Math.Floor(point.X), (int)Math.Floor(point.Y), 0, 0)
        {
            LocationFixed = true;
            SizeFixed = false;
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

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="point">キャプチャー座標</param>
        /// <param name="size">キャプチャーサイズ</param>
        public ManualScreenRectCaptureItem(System.Windows.Point point, System.Windows.Size size) : base((int)Math.Floor(point.X), (int)Math.Floor(point.Y), (int)Math.Floor(size.Width), (int)Math.Floor(size.Height))
        {
            LocationFixed = true;
            SizeFixed = true;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="rect">キャプチャー矩形</param>
        public ManualScreenRectCaptureItem(System.Windows.Rect rect) : this(rect.Location, rect.Size)
        {
        }

        #endregion Constructors(Windows)

        #endregion Constructor
    }
}