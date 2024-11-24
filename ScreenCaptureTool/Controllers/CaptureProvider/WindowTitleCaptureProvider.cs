using ScreenCaptureTool.Models.CaptureItem;

namespace ScreenCaptureTool.Controllers.CaptureProvider
{
    /// <summary>
    /// ウィンドウタイトル指定のキャプチャー機能
    /// </summary>
    public class WindowTitleCaptureProvider : CaptureProvider
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="item">キャプチャーアイテム</param>
        public WindowTitleCaptureProvider(CaptureItemBase item) : base(item)
        {
            // キャプチャーアイテムの型チェック
            if (!(item is WindowTitleCaptureItem))
            {
                throw new System.ArgumentException("item must be WindowTitleCaptureItem", nameof(item));
            }
        }
    }
}