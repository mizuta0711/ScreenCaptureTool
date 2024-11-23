using ScreenCaptureTool.Models.CaptureItem;

using System.Drawing;

namespace ScreenCaptureTool.Controllers.CaptureProvider
{
    /// <summary>
    /// 画面固定矩形キャプチャー機能
    /// </summary>
    public class ScreenRectCaptureProvider : CaptureProvider
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="item">キャプチャーアイテム</param>
        public ScreenRectCaptureProvider(CaptureItemBase item) : base(item)
        {
        }
    }
}