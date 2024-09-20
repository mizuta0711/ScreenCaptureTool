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
        public string FileName { get; set; }
        public BitmapImage Thumbnail { get; set; }

        private int thumbnailWidth;

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

        private int thumbnailHeight;

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

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}