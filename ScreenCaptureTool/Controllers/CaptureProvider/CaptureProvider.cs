using ScreenCaptureTool.Models.CaptureItem;

using System.Drawing;

namespace ScreenCaptureTool.Controllers.CaptureProvider
{
    /// <summary>
    /// キャプチャー機能の基底クラス
    /// </summary>
    public abstract class CaptureProvider
    {
        /// <summary>
        /// キャプチャーアイテム
        /// </summary>
        protected CaptureItem CaptureItem { set; get; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="item">キャプチャーアイテム</param>
        public CaptureProvider(CaptureItem item)
        {
            CaptureItem = item;
        }

        /// <summary>
        /// キャプチャー実行
        /// </summary>
        /// <returns>画像(失敗時はnull)</returns>
        public virtual Bitmap? Capture()
        {
            return CaptureItem?.Capture();
        }
    }
}