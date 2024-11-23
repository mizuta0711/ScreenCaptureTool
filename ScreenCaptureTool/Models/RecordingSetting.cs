using ScreenCaptureTool.Models.CaptureItem;

using System.Drawing;

namespace ScreenCaptureTool.Models
{
    public class RecordingSetting
    {
        #region Struct

        /// <summary>
        /// エッジのトリミング
        /// </summary>
        public struct EdgeInsets
        {
            public int Top;
            public int Bottom;
            public int Left;
            public int Right;

            public EdgeInsets(int top, int bottom, int left, int right)
            {
                Top = top;
                Bottom = bottom;
                Left = left;
                Right = right;
            }
        }

        #endregion Struct

        #region Enum

        /// <summary>
        /// 録画方法
        /// </summary>
        public enum RecordingType
        {
            Window,     // ウィンドウ
            ScreenRect  // 画面の一部
        }

        /// <summary>
        /// 画像保存形式
        /// </summary>
        public enum ImageSaveType
        {
            clipbord,   // クリップボード
            filePNG,    // 画像(PNG形式)
            fileBMP,    // 画像(BMP形式)
            fileJPEG    // 画像(JPEG形式)
        }

        #endregion Enum

        #region Properties

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; } = "名称未設定";

        /// <summary>
        /// 録画方法
        /// </summary>
        public RecordingType Type { get; set; } = RecordingType.Window;

        /// <summary>
        /// ウィンドウタイトル
        /// </summary>
        public string WindowTitle { get; set; } = "";

        /// <summary>
        /// ウィンドウ位置固定
        /// </summary>
        public bool LocationFixed { get; set; } = false;

        /// <summary>
        /// ウィンドウ位置
        /// </summary>
        public Point Location { get; set; } = new Point(0, 0);

        /// <summary>
        /// ウィンドウサイズ固定
        /// </summary>
        public bool SizeFixed { get; set; } = false;

        /// <summary>
        /// ウィンドウサイズ
        /// </summary>
        public Size Size { get; set; } = new Size(0, 0);

        /// <summary>
        /// 画像加工：縁トリミング
        /// </summary>
        public bool TrimEdgeEnabled { get; set; } = false;

        /// <summary>
        /// 画像加工：縁トリミングサイズ
        /// </summary>
        public EdgeInsets TrimEdgeInset { get; set; } = new EdgeInsets(0, 0, 0, 0);

        /// <summary>
        /// 画像加工：リサイズ
        /// </summary>
        public bool ResizeEnabled { get; set; } = false;

        /// <summary>
        /// 画像加工：リサイズサイズ
        /// </summary>
        public Size ResizeSize { get; set; } = new Size(640, 480);

        /// <summary>
        /// 画像保存形式
        /// </summary>
        public ImageSaveType SaveType { get; set; } = ImageSaveType.filePNG;

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public RecordingSetting(string name)
        {
            Name = name;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="name">名称</param>
        /// <param name="captureItem">キャプチャーアイテム</param>
        public RecordingSetting(string name, CaptureItemBase captureItem) : this(name)
        {
            // キャプチャーアイテムから設定を取得
            // ウィンドウキャプチャー
            if (captureItem is WindowTitleCaptureItem windowTitleCaptureItem)
            {
                Type = RecordingType.Window;
                WindowTitle = windowTitleCaptureItem.TargetWindowTitle;
            }

            // 画面の一部キャプチャー
            if (captureItem is ManualScreenRectCaptureItem screenRectCaptureItem)
            {
                Type = RecordingType.ScreenRect;
                LocationFixed = screenRectCaptureItem.LocationFixed;
                Location = new Point(screenRectCaptureItem.TargetRect.Left, screenRectCaptureItem.TargetRect.Top);
                SizeFixed = screenRectCaptureItem.SizeFixed;
                Size = new Size(screenRectCaptureItem.TargetRect.Width, screenRectCaptureItem.TargetRect.Height);
            }
        }

        #endregion Constructor
    }
}