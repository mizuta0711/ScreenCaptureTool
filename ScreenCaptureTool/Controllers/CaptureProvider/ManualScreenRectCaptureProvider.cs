using ScreenCaptureTool.Models.CaptureItem;
using ScreenCaptureTool.Windows;

using System.Drawing;

namespace ScreenCaptureTool.Controllers.CaptureProvider
{
    /// <summary>
    /// ユーザー指定の矩形範囲キャプチャー機能
    /// </summary>
    public class ManualScreenRectCaptureProvider : CaptureProvider
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="item">キャプチャーアイテム</param>
        public ManualScreenRectCaptureProvider(CaptureItemBase item) : base(item)
        {
        }

        /// <summary>
        /// キャプチャーを行う
        /// </summary>
        /// <returns>画像(失敗時はnull)</returns>
        public override Bitmap? Capture()
        {
            var overlayWindow = new CaptureRectOverlayWindow();
            if (overlayWindow.ShowDialog() == true)
            {
                // TODO: インスタンスを置き換えているが、矩形情報だけを更新するように変更する
                // TODO: 画面の拡大率が反映されないので、キャプチャー時に拡大率を考慮するように変更する
                CaptureItem = new ScreenRectCaptureItem(overlayWindow.SelectedRect);
                return base.Capture();
            }
            else
            {
                return null;
            }
        }
    }
}