using ScreenCaptureTool.Models.CaptureItem;

using System;
using System.Drawing;

namespace ScreenCaptureTool.Models
{
    /// <summary>
    /// 撮影設定
    /// </summary>
    [Serializable]
    public class RecordingSettings
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
            Clipboard,  // クリップボード
            FilePNG,    // 画像(PNG形式)
            FileBMP,    // 画像(BMP形式)
            FileJPEG    // 画像(JPEG形式)
        }

        #endregion Enum

        #region Properties

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; } = "名称未設定";

        /// <summary>
        /// 説明
        /// </summary>
        public string Description { get; set; } = "";

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
        public ImageSaveType SaveType { get; set; } = ImageSaveType.FilePNG;

        /// <summary>
        /// ファイルの上書き確認
        /// </summary>
        public bool ConfirmOverrideFile { get; set; } = true;

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public RecordingSettings(string name)
        {
            Name = name;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="name">名称</param>
        /// <param name="captureItem">キャプチャーアイテム</param>
        public RecordingSettings(string name, CaptureItemBase captureItem) : this(name)
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

        /// <summary>
        /// 画像保存形式に対応するファイル拡張子を取得
        /// </summary>
        public string FileExtension
        {
            get
            {
                switch (SaveType)
                {
                    case ImageSaveType.FilePNG:
                        return "png";

                    case ImageSaveType.FileBMP:
                        return "bmp";

                    case ImageSaveType.FileJPEG:
                        return "jpg";

                    default:
                        return "";
                }
            }
        }

        #endregion Constructor

        #region Methods(Public)

        /// <summary>
        /// 設定に基づいてキャプチャーアイテムを取得する
        /// </summary>
        /// <returns>キャプチャーアイテム</returns>
        /// <exception cref="InvalidOperationException">生成できなかった場合</exception>
        public CaptureItemBase GetCaptureItem()
        {
            // キャプチャーアイテムを生成
            switch (Type)
            {
                // ウィンドウキャプチャー
                case RecordingType.Window:
                    return new WindowTitleCaptureItem(WindowTitle);

                // 画面の一部キャプチャー
                case RecordingType.ScreenRect:
                    if (LocationFixed && SizeFixed)
                    {
                        // 位置とサイズが固定されている場合
                        return new ScreenRectCaptureItem(Location, Size);
                    }
                    else if (LocationFixed)
                    {
                        // 位置が固定されている場合
                        return new ManualScreenRectCaptureItem(Location);
                    }
                    else if (SizeFixed)
                    {
                        // サイズが固定されている場合
                        return new ManualScreenRectCaptureItem(Size);
                    }
                    // 位置とサイズが固定されていない場合
                    return new ManualScreenRectCaptureItem();
            }

            // 生成できなかった場合は例外を投げる
            throw new InvalidOperationException();
        }

        #endregion Methods(Public)
    }
}