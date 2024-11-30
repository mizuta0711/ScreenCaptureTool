using System.ComponentModel;
using System.Windows.Media.Imaging;

namespace ScreenCaptureTool.Models
{
    /// <summary>
    /// 画像ファイル情報を保持するクラス
    /// </summary>
    public class ImageFile : INotifyPropertyChanged
    {
        #region Variables

        /// <summary>
        /// 選択状態
        /// </summary>
        private bool isSelected;

        /// <summary>
        /// ファイル名
        /// </summary>
        private string? fileName;

        /// <summary>
        /// ファイルパス
        /// </summary>
        private string? imageUrl;

        /// <summary>
        /// サムネイル画像：幅
        /// </summary>
        private int thumbnailWidth;

        /// <summary>
        /// サムネイル画像：高さ
        /// </summary>
        private int thumbnailHeight;

        /// <summary>
        /// 値変更のイベントハンドラ
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        #endregion Variables

        #region Properties

        /// <summary>
        /// 選択状態
        /// </summary>
        public bool IsSelected
        {
            get => isSelected;
            set
            {
                isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        /// <summary>
        /// ファイル名
        /// </summary>
        public string? FileName
        {
            get => fileName;
            set
            {
                if (fileName != value)
                {
                    fileName = value;
                    OnPropertyChanged(nameof(FileName));
                }
            }
        }

        /// <summary>
        /// ファイルパス
        /// </summary>
        public string? ImageUrl
        {
            get => imageUrl;
            set
            {
                if (imageUrl != value)
                {
                    imageUrl = value;
                    OnPropertyChanged(nameof(imageUrl));
                }
            }
        }

        /// <summary>
        /// サムネイル画像：幅
        /// </summary>
        public int ThumbnailWidth
        {
            get { return thumbnailWidth; }
            set
            {
                if (thumbnailWidth != value)
                {
                    thumbnailWidth = value;
                    OnPropertyChanged(nameof(ThumbnailWidth));
                }
            }
        }

        /// <summary>
        /// サムネイル画像：高さ
        /// </summary>
        public int ThumbnailHeight
        {
            get { return thumbnailHeight; }
            set
            {
                if (thumbnailHeight != value)
                {
                    thumbnailHeight = value;
                    OnPropertyChanged(nameof(ThumbnailHeight));
                }
            }
        }

        #endregion Properties

        #region Events

        /// <summary>
        /// プロパティの値変更イベント
        /// </summary>
        /// <param name="propertyName">プロパティ名</param>
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion Events
    }
}