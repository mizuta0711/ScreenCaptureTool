using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace ScreenCaptureTool
{
    /// <summary>
    /// 画像ファイル情報を保持するクラス
    /// </summary>
    public class ImageFile : INotifyPropertyChanged
    {
        #region Variables

        /// <summary>
        /// ファイル名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 画像イメージ
        /// </summary>
        public BitmapImage Thumbnail { get; set; }

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
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Variables

        #region Properties

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