using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace ScreenCaptureTool.Windows.Controls
{
    /// <summary>
    /// 非同期で画像を読み込むカスタムImageコントロール
    /// </summary>
    public class AsyncImage : Image
    {
        #region DependencyProperty: ImageUrl

        /// <summary>
        /// 画像のURLを指定するDependencyProperty
        /// </summary>
        public static readonly DependencyProperty ImageUrlProperty =
            DependencyProperty.Register("ImageUrl", typeof(string), typeof(AsyncImage),
                new UIPropertyMetadata(string.Empty, OnImageUrlChanged));

        /// <summary>
        /// 画像のURL（DependencyPropertyとしてバインド可能）
        /// </summary>
        public string ImageUrl
        {
            get { return (string)GetValue(ImageUrlProperty); }
            set { SetValue(ImageUrlProperty, value); }
        }

        #endregion DependencyProperty: ImageUrl

        #region Event Handlers

        /// <summary>
        /// ImageUrlプロパティが変更されたときに呼び出されるコールバック
        /// </summary>
        /// <param name="d">DependencyObjectインスタンス</param>
        /// <param name="e">プロパティ変更イベントのデータ</param>
        private static void OnImageUrlChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (AsyncImage)d;
            // 非同期で新しい画像を読み込む処理を開始
            control.LoadImageAsync((string)e.NewValue);
        }

        #endregion Event Handlers

        #region Private Methods

        /// <summary>
        /// 指定されたURLから非同期で画像を読み込む
        /// </summary>
        /// <param name="url">画像URL</param>
        private async void LoadImageAsync(string url)
        {
            // プレースホルダー画像を一時的に表示
            this.Source = GetPlaceholderImage();

            try
            {
                // 画像を非同期で読み込む
                var bitmap = await Task.Run(() =>
                {
                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.UriSource = new Uri(url, UriKind.Absolute);
                    bi.DecodePixelWidth = 200; // TODO: とりあえず固定値…後で設定可能にする
                    bi.EndInit();
                    bi.Freeze(); // 他のスレッドからも安全にアクセス可能にする
                    return bi;
                });

                // 読み込んだ画像をImageコントロールのSourceに設定
                this.Source = bitmap;
            }
            catch
            {
                // エラー時にはエラー画像を表示
                this.Source = GetErrorImage();
            }
        }

        /// <summary>
        /// プレースホルダー画像を取得する
        /// </summary>
        /// <returns>プレースホルダー用BitmapImage</returns>
        private BitmapImage GetPlaceholderImage()
        {
            // TODO: プレースホルダー画像の生成処理を実装
            return null;
        }

        /// <summary>
        /// エラー時に表示する画像を取得する
        /// </summary>
        /// <returns>エラー表示用BitmapImage</returns>
        private BitmapImage GetErrorImage()
        {
            // TODO: エラー画像の生成処理を実装
            return null;
        }

        #endregion Private Methods
    }
}