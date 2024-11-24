using System;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using ScreenCaptureTool.Models;
using System.Windows.Interop;

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
        internal static bool SaveToFile(Bitmap bitmap, string filePath, RecordingSettings.ImageSaveType saveType, bool confirm)
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
                case RecordingSettings.ImageSaveType.FilePNG:
                    bitmap.Save(filePath, ImageFormat.Png);
                    return true;

                case RecordingSettings.ImageSaveType.FileBMP:
                    bitmap.Save(filePath, ImageFormat.Bmp);
                    return true;

                case RecordingSettings.ImageSaveType.FileJPEG:
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

        #endregion Methods(Static)
    }
}