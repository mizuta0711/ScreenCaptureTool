using ScreenCaptureTool.Models.CaptureItem;

using System;

namespace ScreenCaptureTool.Controllers.CaptureProvider
{
    public class CaptureProviderFactory
    {
        /// <summary>
        /// キャプチャープロバイダーのインスタンスを生成する
        /// </summary>
        /// <param name="captureItem">キャプチャーアイテム</param>
        /// <returns>キャプチャープロバイダー</returns>
        /// <exception cref="ArgumentException">対応していないアイテム</exception>
        public static CaptureProvider InstantiateCaptureProvider(CaptureItem captureItem)
        {
            if (captureItem.GetType() == typeof(ScreenRectCaptureItem))
            {
                return new ScreenRectCaptureProvider(captureItem);
            }
            if (captureItem.GetType() == typeof(ManualScreenRectCaptureItem))
            {
                return new ManualScreenRectCaptureProvider(captureItem);
            }
            if (captureItem.GetType() == typeof(WindowTitleCaptureItem))
            {
                return new WindowTitleCaptureProvider(captureItem);
            }

            // 対応していないCaptureItemの場合は例外を投げる
            throw new ArgumentException();
        }
    }
}