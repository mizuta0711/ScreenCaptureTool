using System;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using ScreenCaptureTool.Models;
using System.Windows.Interop;
using static ScreenCaptureTool.Models.RecordingSetting;

namespace ScreenCaptureTool.Utilities
{
    /// <summary>
    /// Bitmapのヘルパークラス
    /// </summary>
    internal class BitmapHelper
    {
        #region Methods(Static)

        /// <summary>
        /// BitmapSourceからBitmapImageへの変換
        /// </summary>
        /// <param name="bitmapSource">BitmapSource</param>
        /// <returns>変換後のBitmapImage</returns>
        internal static BitmapImage ConvertBitmapSourceToBitmapImage(BitmapSource bitmapSource)
        {
            // BitmapSourceをMemoryStreamに保存
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // BitmapEncoderでBitmapSourceをエンコード
                BitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                encoder.Save(memoryStream);

                // MemoryStreamからBitmapImageを作成
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = new MemoryStream(memoryStream.ToArray());
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }

        /// <summary>
        /// 指定されたファイルをBitmapImage形式で読み込む
        /// </summary>
        /// <param name="filePath">パス</param>
        /// <returns>BitmapImage</returns>
        internal static BitmapImage LoadBitmapImage(string filePath)
        {
            using (Stream stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete
            ))
            {
                // ロックしないように指定したstreamを使用する。
                BitmapDecoder decoder = BitmapDecoder.Create(
                    stream,
                    BitmapCreateOptions.None, // この辺のオプションは適宜
                    BitmapCacheOption.Default // これも
                );
                BitmapSource bmp = new WriteableBitmap(decoder.Frames[0]);
                bmp.Freeze();

                // BitmapImage形式に変換して返す
                return ConvertBitmapSourceToBitmapImage(bmp);
            }
        }

        /// <summary>
        /// Bitmapを保存する
        /// </summary>
        /// <param name="bitmap">Bitmap</param>
        /// <param name="filePath">保存先のパス</param>
        /// <param name="saveType">保存形式</param>
        /// <param name="confirm">上書き確認を行うか</param>
        /// <returns>true: 保存 / false: 失敗</returns>
        internal static bool SaveToFile(Bitmap bitmap, string filePath, RecordingSetting.ImageSaveType saveType, bool confirm)
        {
            // フォルダが存在しない場合は作成する
            if (!Directory.Exists(Path.GetDirectoryName(filePath)))
            {
                if (Path.GetDirectoryName(filePath) is string parentFolderPath)
                {
                    Directory.CreateDirectory(parentFolderPath);
                }
            }

            // ファイルが既に存在する場合、上書き確認ダイアログを表示
            if (confirm && File.Exists(filePath))
            {
                var result = MessageBox.Show(
                    "このファイルは既に存在します。上書きしますか？",
                    "上書き確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.No)
                {
                    // 上書きをキャンセル
                    return false;
                }
            }

            // 指定された形式で保存
            switch (saveType)
            {
                case RecordingSetting.ImageSaveType.FilePNG:
                    bitmap.Save(filePath, ImageFormat.Png);
                    return true;

                case RecordingSetting.ImageSaveType.FileBMP:
                    bitmap.Save(filePath, ImageFormat.Bmp);
                    return true;

                case RecordingSetting.ImageSaveType.FileJPEG:
                    bitmap.Save(filePath, ImageFormat.Jpeg);
                    return true;
            }

            return false;
        }

        /// <summary>
        /// クリップボードにBitmapをコピー
        /// </summary>
        /// <param name="bitmap">Bitmap</param>
        internal static void CopyToClipboard(Bitmap bitmap)
        {
            IntPtr hBitmap = bitmap.GetHbitmap(); // HBitmap を取得
            try
            {
                // HBitmap を BitmapSource に変換
                var bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

                // クリップボードに画像をコピー
                Clipboard.SetImage(bitmapSource);
            }
            finally
            {
                // HBitmap のリソースを解放
                Win32API.DeleteObject(hBitmap);
            }
        }

        /// <summary>
        /// 渡されたBitmapからEdgeInsetsで指定された縁をトリミングします。
        /// </summary>
        /// <param name="bitmap">元のBitmap</param>
        /// <param name="insets">トリミングするエッジのピクセル量</param>
        /// <returns>トリミングされたBitmap</returns>
        /// <exception cref="ArgumentNullException">bitmapがnullの場合</exception>
        /// <exception cref="ArgumentException">トリミング後のサイズが無効な場合</exception>
        internal static Bitmap TrimEdge(Bitmap bitmap, EdgeInsets insets)
        {
            if (bitmap == null)
                throw new ArgumentNullException(nameof(bitmap));

            // トリミング後のサイズを計算
            int trimmedWidth = bitmap.Width - insets.Left - insets.Right;
            int trimmedHeight = bitmap.Height - insets.Top - insets.Bottom;

            if (trimmedWidth <= 0 || trimmedHeight <= 0)
                throw new ArgumentException("トリミング後のサイズが無効です。EdgeInsetsが大きすぎる可能性があります。");

            // トリミング領域を計算
            Rectangle trimRectangle = new Rectangle(insets.Left, insets.Top, trimmedWidth, trimmedHeight);

            // 新しいBitmapを作成し、指定領域を描画
            Bitmap trimmedBitmap = new Bitmap(trimmedWidth, trimmedHeight);
            using (Graphics g = Graphics.FromImage(trimmedBitmap))
            {
                g.DrawImage(bitmap, new Rectangle(0, 0, trimmedWidth, trimmedHeight), trimRectangle, GraphicsUnit.Pixel);
            }

            return trimmedBitmap;
        }

        /// <summary>
        /// 指定された幅と高さにBitmapをリサイズします。
        /// </summary>
        /// <param name="bitmap">元のBitmap</param>
        /// <param name="width">リサイズ後の幅</param>
        /// <param name="height">リサイズ後の高さ</param>
        /// <param name="keepAspect">アスペクト比を保つかどうか</param>
        /// <returns>リサイズされたBitmap</returns>
        /// <remarks>画像よりも大きいサイズが指定された場合、拡大はされません</remarks>
        /// <exception cref="ArgumentNullException">bitmapがnullの場合</exception>
        /// <exception cref="ArgumentException">幅と高さが正の値でない場合</exception>
        internal static Bitmap Resize(Bitmap bitmap, int width, int height, bool keepAspect = true)
        {
            if (bitmap == null)
                throw new ArgumentNullException(nameof(bitmap));

            if (width <= 0 || height <= 0)
                throw new ArgumentException("幅と高さは正の値である必要があります。");

            // アスペクト比を計算
            int targetWidth = width;
            int targetHeight = height;

            if (keepAspect)
            {
                double aspectRatio = (double)bitmap.Width / bitmap.Height;
                if (width / (double)height > aspectRatio)
                {
                    // 幅より高さに制約がある場合
                    targetWidth = (int)(height * aspectRatio);
                }
                else
                {
                    // 高さより幅に制約がある場合
                    targetHeight = (int)(width / aspectRatio);
                }
            }

            // リサイズが不要な場合はそのまま返す(拡大はしない)
            if (bitmap.Width <= targetWidth && bitmap.Height <= targetHeight)
            {
                return bitmap;
            }

            // 新しいBitmapを作成
            Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight);

            // 描画設定を行いリサイズを実行
            using (Graphics g = Graphics.FromImage(resizedBitmap))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.DrawImage(bitmap, 0, 0, targetWidth, targetHeight);
            }

            return resizedBitmap;
        }

        #endregion Methods(Static)
    }
}