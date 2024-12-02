using ScreenCaptureTool.Models.CaptureItem;
using ScreenCaptureTool.Utilities;

using System;
using System.Drawing;
using System.IO;
using System.Xml.Serialization;

namespace ScreenCaptureTool.Models
{
    /// <summary>
    /// 撮影設定
    /// </summary>
    public class RecordingSetting
    {
        #region Struct

        /// <summary>
        /// エッジのトリミング
        /// </summary>
        public struct EdgeInsets
        {
            public int Top;
            public int Left;
            public int Bottom;
            public int Right;

            /// <summary>
            /// 空かどうか
            /// </summary>
            public bool IsEmpty => Top == 0 && Left == 0 && Bottom == 0 && Right == 0;

            public EdgeInsets(int top, int left, int bottom, int right)
            {
                Top = top;
                Left = left;
                Bottom = bottom;
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
        [XmlIgnore]
        public Point Location { get; set; } = new Point(0, 0);

        /// <summary>
        /// ウィンドウサイズ固定
        /// </summary>
        public bool SizeFixed { get; set; } = false;

        /// <summary>
        /// ウィンドウサイズ
        /// </summary>
        [XmlIgnore]
        public Size Size { get; set; } = new Size(0, 0);

        /// <summary>
        /// 画像加工：縁トリミング
        /// </summary>
        public bool TrimEdgeEnabled { get; set; } = false;

        /// <summary>
        /// 画像加工：縁トリミングサイズ
        /// </summary>
        [XmlIgnore]
        public EdgeInsets TrimEdgeInset { get; set; } = new EdgeInsets(0, 0, 0, 0);

        /// <summary>
        /// 画像加工：リサイズ
        /// </summary>
        public bool ResizeEnabled { get; set; } = false;

        /// <summary>
        /// 画像加工：リサイズサイズ
        /// </summary>
        [XmlIgnore]
        public Size ResizeSize { get; set; } = new Size(640, 480);

        /// <summary>
        /// 画像保存形式
        /// </summary>
        public ImageSaveType SaveType { get; set; } = ImageSaveType.FilePNG;

        /// <summary>
        /// ファイル名フォーマット
        /// </summary>
        public string FilenameFormat { get; set; } = "%NAME%";

        /// <summary>
        /// 名前を付けて保存を有効にする
        /// </summary>
        public bool SaveAsEnable { get; set; } = false;

        /// <summary>
        /// ファイルの上書き確認
        /// </summary>
        public bool ConfirmOverrideFile { get; set; } = true;

        #region Properties(Serialize)

        /// <summary>
        /// シリアライズ用：ウィンドウ位置
        /// </summary>
        [XmlElement("TrimEdgeInset")]
        public string TrimEdgeInsetSerialized
        {
            get => $"{TrimEdgeInset.Top},{TrimEdgeInset.Left},{TrimEdgeInset.Bottom},{TrimEdgeInset.Right}";
            set
            {
                try
                {
                    var parts = value.Split(',');
                    TrimEdgeInset = new EdgeInsets(int.Parse(parts[0]), int.Parse(parts[1]), int.Parse(parts[2]), int.Parse(parts[3]));
                }
                catch (Exception e)
                {
                    TrimEdgeInset = new EdgeInsets(0, 0, 0, 0);
                }
            }
        }

        /// <summary>
        /// シリアライズ用：ウィンドウ位置
        /// </summary>
        [XmlElement("Location")]
        public string LocationSerialized
        {
            get => $"{Location.X},{Location.Y}";
            set
            {
                try
                {
                    var parts = value.Split(',');
                    Location = new Point(int.Parse(parts[0]), int.Parse(parts[1]));
                }
                catch (Exception e)
                {
                    Location = new Point(0, 0);
                }
            }
        }

        /// <summary>
        /// シリアライズ用：ウィンドウサイズ
        /// </summary>
        [XmlElement("Size")]
        public string SizeSerialized
        {
            get => $"{Size.Width},{Size.Height}";
            set
            {
                try
                {
                    var parts = value.Split(',');
                    Size = new Size(int.Parse(parts[0]), int.Parse(parts[1]));
                }
                catch (Exception e)
                {
                    Size = new Size(0, 0);
                }
            }
        }

        /// <summary>
        /// シリアライズ用：リサイズサイズ
        /// </summary>
        [XmlElement("ResizeSize")]
        public string ResizeSizeSerialized
        {
            get => $"{ResizeSize.Width},{ResizeSize.Height}";
            set
            {
                try
                {
                    var parts = value.Split(',');
                    ResizeSize = new Size(int.Parse(parts[0]), int.Parse(parts[1]));
                }
                catch (Exception e)
                {
                    ResizeSize = new Size(0, 0);
                }
            }
        }

        #endregion Properties(Serialize)

        #region Properties(Other)

        /// <summary>
        /// 撮影情報を取得
        /// </summary>
        [XmlIgnore]
        public string Information
        {
            get
            {
                if (Type == RecordingType.Window)
                {
                    return $"ウィンドウ:{WindowTitle}";
                }
                if (LocationFixed && SizeFixed)
                {
                    return $"位置:({Location.X},{Location.Y}) / サイズ:({Size.Width},{Size.Height})";
                }
                if (LocationFixed)
                {
                    return $"位置:({Location.X},{Location.Y}) / サイズ:撮影時に選択";
                }
                if (SizeFixed)
                {
                    return $"位置:撮影時に選択 / サイズ:({Size.Width},{Size.Height})";
                }
                return "撮影時に範囲を選択";
            }
        }

        /// <summary>
        /// 保存先情報を取得
        /// </summary>
        [XmlIgnore]
        public string SaveInformation
        {
            get
            {
                if (SaveType == ImageSaveType.Clipboard)
                {
                    return $"クリップボードに保存";
                }
                // ファイルに保存
                return FilenameFormat + FileExtension;
            }
        }

        /// <summary>
        /// 画像保存形式に対応するファイル拡張子を取得
        /// </summary>
        [XmlIgnore]
        private string FileExtension
        {
            get
            {
                switch (SaveType)
                {
                    case ImageSaveType.FilePNG:
                        return ".png";

                    case ImageSaveType.FileBMP:
                        return ".bmp";

                    case ImageSaveType.FileJPEG:
                        return ".jpg";

                    default:
                        return "";
                }
            }
        }

        /// <summary>
        /// 保存ファイル名(ファイル拡張子も含む)
        /// </summary>
        [XmlIgnore]
        public string SaveFileName
        {
            get
            {
                var formatter = new FileNameFormatter(FilenameFormat, Name, FileExtension);
                return formatter.FormattedName;
            }
        }

        #endregion Properties(Other)

        #endregion Properties

        #region Constructor

        /// <summary>
        /// デフォルトコンストラクタ
        /// </summary>
        public RecordingSetting()
        {
        }

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

        #region Methods(Public)

        /// <summary>
        /// ファイル名の連番をインクリメントして書式文字列を更新
        /// </summary>
        public void IncrimentFilenameFormatNumber()
        {
            var formatter = new FileNameFormatter(FilenameFormat, Name, FileExtension);
            FilenameFormat = formatter.IncriementNumber();
        }

        /// <summary>
        /// 画像保存形式を取得
        /// </summary>
        /// <param name="path">ファイルパス</param>
        /// <returns>画像保存形式(PNG/BMP/JPEG)</returns>
        /// <remarks>不明な拡張子の場合はPNGと判定</remarks>
        public static ImageSaveType GetImageSaveType(string path)
        {
            var ext = Path.GetExtension(path).ToLower();
            switch (ext)
            {
                case ".png":
                    return ImageSaveType.FilePNG;

                case ".bmp":
                    return ImageSaveType.FileBMP;

                case ".jpg":
                case ".jpeg":
                    return ImageSaveType.FileJPEG;

                default:
                    return ImageSaveType.FilePNG;
            }
        }

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