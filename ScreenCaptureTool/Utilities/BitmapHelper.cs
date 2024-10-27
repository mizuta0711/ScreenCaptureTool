using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace ScreenCaptureTool.Utilities
{
    /// <summary>
    /// Bitmapのヘルパークラス
    /// </summary>
    internal class BitmapHelper
    {
        #region BitmapUtility

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

        #endregion BitmapUtility
    }
}